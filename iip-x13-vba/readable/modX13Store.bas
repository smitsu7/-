Attribute VB_Name = "modX13Store"
Option Explicit

' 役割      : 比較候補（試行）と最終決定したモデルを、ファイルの照合付きで保存し、読み戻して確かめる。
'             試行と最終結果は別のフォルダ・別のシートに置き、名前で区別する。
' 主な入口  : SaveTrial（試行の保存）/ ValidateTrial（試行の検査）/ TrialsOfRun / TrialFolder
'             FinalizeTrial（最終保存）/ ValidateFinal / VerifyFinalAgainstSheet（再読込み後の照合）/ CurrentAdoptions
' 依存先    : modX13IO、modMain、modPackagePaths、modSheetNames、modHolidayIIP、modX13Results、modX13Calendar、
'             modX13Options（実績144か月の定数）、modLog
' 保存の順序: <保存先>.pending へコピー → 照合一覧（ファイル名・バイト数・CRC32）を作る → 照合 →
'             本来の名前へ移す → もう一度照合・出力表を検査 → 完了印（_trial.complete／_final.complete）を書く → 表へ記録
'             途中で失敗した場合は、そのフォルダに _failed.txt を書き、この処理で追記した行を消す。
'             過去の試行・最終結果のフォルダは消さない・上書きしない。
' 保存先    : 試行  … X13_CANDIDATE_DIR\<実行ID>\<系列ID>\<試行ID>\
'             最終  … X13_FINAL_DIR\<系列ID>\<保存ID>\
' 補足      : CRC32は偶発的な変化の検出に使う（改ざん防止の認証ではない）。祝日CSVの同一性はSHA-256で確かめる。
' 出力形式  : X-13が書いたファイルは、ASCII版（.out・.err・jobs.log）もHTML版（.html・_err.html・jobs_log.html）も
'             名前と中身を変えずに保存する。HTML版のときは、VBAがHTMLから作った解析用テキスト
'             （_report_text_from_html.txt）も同じフォルダに置く。これはX-13の出力ではなく、人が確認する原本は .html。
'             _trial.meta の REPORT_FORMAT・REPORT_MAIN・REPORT_ERROR・REPORT_TEXT・BATCH_LOG・ENGINE に形式とファイル名を記録する。
'             REPORT_FORMAT が無いフォルダ（HTML対応前に保存した試行）は、ASCII版として従来どおり読む。

Private Const TRIAL_MARKER As String = "TRIAL_R11"
Private Const FINAL_MARKER As String = "FINAL_R11"
Private Const TRIAL_META As String = "_trial.meta"
Private Const TRIAL_MANIFEST As String = "_manifest.tsv"
Private Const TRIAL_COMPLETE As String = "_trial.complete"
Private Const FINAL_META As String = "_final.meta"
Private Const FINAL_MANIFEST As String = "_final_manifest.tsv"
Private Const FINAL_COMPLETE As String = "_final.complete"
Private Const FAILED_MARK As String = "_failed.txt"

Private mBusy As Boolean

' 試行_一覧・最終_保存履歴の「状態」。
Private Function StateSaved() As String
    StateSaved = "保存済"
End Function

Private Function StateCurrent() As String
    StateCurrent = "採用中"
End Function

Private Function StatePrevious() As String
    StatePrevious = "過去の採用"
End Function

' ================= 表の見出し =================

Public Function TrialHeaders() As Variant
    TrialHeaders = Array("試行ID", "実行ID", "候補ID", "系列ID", "設定ID", "段階", "保存フォルダ", "ジョブ名", "保存日時UTC", "状態")
End Function

Public Function AdoptionHeaders() As Variant
    AdoptionHeaders = Array("保存ID", "試行ID", "系列ID", "設定ID", "実行ID", "決定方式", "保存フォルダ", "保存日時UTC", "状態")
End Function

Public Function DecisionHeaders() As Variant
    DecisionHeaders = Array("保存ID", "試行ID", "系列ID", "決定方式", "順位の統計量", "統計量の値", "比較群", "統計量の順位", _
        "条件付き順位", "変数検定を条件に使用", "有意水準", "同じ値の候補数", "決定日時UTC", "注記")
End Function

Public Function FinalMonthlyHeaders() As Variant
    FinalMonthlyHeaders = Array("保存ID", "系列ID", "年月", "区分", "原指数_原データ", "計算用の値_DAT", "季節調整値_D11", _
        "季節成分_D10", "季節・カレンダー結合成分_D16", "カレンダー成分_D18", "D18の出所", "試行ID")
End Function

Public Function FinalSettingsHeaders() As Variant
    FinalSettingsHeaders = Array("保存ID", "系列ID", "設定キー", "値", "出所", "系列別列名")
End Function

' 目的: 保存に使うシートを見出し付きで用意する（既存の見出しが違えば停止）。
Public Sub EnsureStoreSheets()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.TrialSheet()), TrialHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.DecisionSheet()), DecisionHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.AdoptionSheet()), AdoptionHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.FinalMonthlySheet()), FinalMonthlyHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.FinalSettingsSheet()), FinalSettingsHeaders()
End Sub

' ================= 試行の保存 =================

' 目的  : 検査に合格した1候補の作業フォルダを、試行として候補フォルダへ保存する。
' 入力  : job（modX13Runの候補）、jobFolder（作業フォルダ内のジョブフォルダ）、batchFolder（jobs.mtaのあるフォルダ）
' 出力  : 試行ID（例 R_20261017_183000_T00001）
' 副作用: 候補フォルダを新しく作り、試行_一覧へ1行追加する。
' エラー: コピー・照合・出力表の検査に失敗した場合は、_failed.txtを書いてから停止する（行は追加しない）。
Public Function SaveTrial(ByVal job As Object, ByVal jobFolder As String, ByVal batchFolder As String) As String
    Dim fso As Object, root As String, trial As String, folder As String, stage As String, file As Object
    Dim meta As Object, stem As String, ws As Worksheet, row As Long, owned As Boolean, specText As String
    Dim errorNumber As Long, errorText As String, n As Long, report As Object, indexName As String
    If mBusy Then Err.Raise vbObjectError + 3400, "SaveTrial", "別の保存処理を実行中です。"
    On Error GoTo Failed
    mBusy = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    stem = CStr(job("stem"))
    root = modPackagePaths.RequireFolder("X13_CANDIDATE_DIR", True)
    Set ws = ThisWorkbook.Worksheets(modSheetNames.TrialSheet())
    modX13IO.Headers ws, TrialHeaders()
    ' 試行IDは実行ID＋通し番号。保存に失敗したフォルダ（_failed.txt付き）の番号は使い直さない。
    n = Application.WorksheetFunction.CountIf(ws.Columns(2), job("run"))
    Do
        n = n + 1
        trial = job("run") & "_T" & Format$(n, "00000")
        folder = modMain.CombinePath(modMain.CombinePath(modMain.CombinePath(root, CStr(job("run"))), CStr(job("series"))), trial)
        stage = folder & ".pending"
    Loop While fso.FolderExists(folder) Or fso.FolderExists(stage) Or Application.WorksheetFunction.CountIf(ws.Columns(1), trial) > 0
    modPackagePaths.EnsureFolderTree fso.GetParentFolderName(folder)
    fso.CreateFolder stage
    owned = True

    ' 1) コピー：ジョブフォルダの全ファイル（SPC・DAT・X-13の出力・入力記録・祝日の複製）とバッチの記録
    '    バッチのログは名前の頭に _batch_ を付ける（ASCII版 jobs.log → _batch_jobs.log、HTML版 jobs_log.html → _batch_jobs_log.html）。
    '    HTML版の索引（jobs_mta.html）があれば _batch_jobs_mta.html として残す。
    If Len(CStr(job("reportFormat"))) = 0 Or Len(CStr(job("batchLog"))) = 0 Then Err.Raise vbObjectError + 3401, "SaveTrial", _
        "出力の検査で出力形式が記録されていません（" & stem & "）"
    For Each file In fso.GetFolder(jobFolder).Files
        modX13IO.CopyFileVerified file.path, modMain.CombinePath(stage, file.Name)
    Next file
    modX13IO.CopyFileVerified modMain.CombinePath(batchFolder, CStr(job("batchLog"))), modMain.CombinePath(stage, "_batch_" & job("batchLog"))
    indexName = modX13Results.BatchIndexFile(batchFolder)
    If Len(indexName) > 0 Then modX13IO.CopyFileVerified modMain.CombinePath(batchFolder, indexName), modMain.CombinePath(stage, "_batch_" & indexName)
    modX13IO.CopyFileVerified modMain.CombinePath(batchFolder, "jobs.mta"), modMain.CombinePath(stage, "_batch_jobs.mta")
    modX13IO.CopyFileVerified modMain.CombinePath(batchFolder, "runner.status.txt"), modMain.CombinePath(stage, "_runner_status.txt")
    ' 実行したSPCはそのまま残し、このフォルダだけで再実行できる版（ファイル参照からジョブ名の階層を外したもの）を別に作る。
    specText = modX13IO.ReadText(modMain.CombinePath(stage, stem & ".spc"), "UTF8")
    modX13IO.WriteNewText modMain.CombinePath(stage, "_portable.spc"), Replace(specText, "file=""" & stem & "/", "file=""")

    ' 2) 記録：何をどう計算したか（人が読める値と、再検査に使う値）
    Set meta = CreateObject("Scripting.Dictionary")
    meta.Add "VERSION", "r11"
    meta.Add "TRIAL", trial
    meta.Add "RUN", CStr(job("run"))
    meta.Add "CANDIDATE", CStr(job("candidate"))
    meta.Add "SERIES", CStr(job("series"))
    meta.Add "SETTING", CStr(job("setting"))
    meta.Add "STAGE", CStr(job("stage"))
    meta.Add "STEM", stem
    meta.Add "CALENDAR", CStr(job("calendar"))
    meta.Add "MODEL_MODE", CStr(job("modelMode"))
    meta.Add "ARIMA", CStr(job("arima"))
    meta.Add "OUTLIER_MODE", CStr(job("outlierMode"))
    meta.Add "SPECIFIED_OUTLIERS", CStr(job("specified"))
    meta.Add "FROZEN_OUTLIERS", CStr(job("frozen"))
    meta.Add "REFERENCE_TRIAL", CStr(job("reference"))
    meta.Add "FIRST", ReadCandidateValue(stage, "FIRST")
    meta.Add "ACTUAL_MONTHS", ReadCandidateValue(stage, "ACTUAL_MONTHS")
    meta.Add "FORECAST_MONTHS", ReadCandidateValue(stage, "FORECAST_MONTHS")
    meta.Add "DATA_CRC", modX13IO.FileCRC(modMain.CombinePath(stage, stem & ".dat"))
    meta.Add "SPEC_CRC", modX13IO.FileCRC(modMain.CombinePath(stage, stem & ".spc"))
    meta.Add "D18_ORIGIN", IIf(fso.FileExists(modMain.CombinePath(stage, "_d18_origin.txt")), "VBA_NEUTRAL", "X13")
    meta.Add "BATCH", fso.GetFileName(batchFolder)
    meta.Add "EXIT_CODE", CStr(job("exitCode"))
    ' 出力形式と、X-13が書いた主出力・エラー記録の実際のファイル名（保存したファイルから決め直す）。
    Set report = modX13Results.DetectReport(stage, stem)
    If report("Format") <> CStr(job("reportFormat")) Then Err.Raise vbObjectError + 3401, "SaveTrial", _
        "保存したファイルの出力形式（" & report("Format") & "）が検査時（" & job("reportFormat") & "）と違います"
    meta.Add "REPORT_FORMAT", CStr(report("Format"))
    meta.Add "REPORT_MAIN", CStr(report("Main"))
    meta.Add "REPORT_ERROR", CStr(report("Error"))
    meta.Add "REPORT_TEXT", IIf(report("Format") = modX13Results.REPORT_HTML, modX13Results.REPORT_TEXT_FILE, "")
    meta.Add "BATCH_LOG", "_batch_" & job("batchLog")
    meta.Add "ENGINE", CStr(job("engine"))
    meta.Add "ENGINE_BYTES", CStr(job("engineBytes"))
    meta.Add "SAVED_UTC", modMain.UtcNowText()
    modX13IO.WriteMetaFile modMain.CombinePath(stage, TRIAL_META), meta

    ' 3) 照合一覧 → 移動 → 再照合 → 出力表の検査 → 完了印
    WriteManifest stage, TRIAL_MANIFEST
    VerifyManifest stage, TRIAL_MANIFEST
    fso.MoveFolder stage, folder
    VerifyManifest folder, TRIAL_MANIFEST
    ValidateTrialFolder folder, meta
    modX13IO.WriteNewText modMain.CombinePath(folder, TRIAL_COMPLETE), TRIAL_MARKER & vbCrLf

    ' 4) 表へ記録（ここまで成功した試行だけが比較の対象になる）
    row = modX13IO.NextRow(ws)
    ws.Cells(row, 1).Resize(1, UBound(TrialHeaders()) + 1).NumberFormat = "@"
    ws.Cells(row, 1).Resize(1, UBound(TrialHeaders()) + 1).Value2 = Array(trial, job("run"), job("candidate"), job("series"), _
        job("setting"), job("stage"), modPackagePaths.RelativeToPackage(folder), stem, meta("SAVED_UTC"), StateSaved())
    mBusy = False
    SaveTrial = trial
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    mBusy = False
    If owned Then MarkFailure folder, stage, errorText
    Err.Raise errorNumber, "SaveTrial", errorText
End Function

' 候補の定義ファイル（_candidate.txt）から値を読む。
Private Function ReadCandidateValue(ByVal folder As String, ByVal key As String) As String
    Dim values As Object
    Set values = modX13IO.ReadMetaFile(modMain.CombinePath(folder, "_candidate.txt"))
    If Not values.Exists(key) Then Err.Raise vbObjectError + 3402, "ReadCandidateValue", "_candidate.txt に " & key & " がありません"
    ReadCandidateValue = CStr(values(key))
End Function

' 目的  : 保存フォルダの中身が記録と合っているか確かめる（試行・最終の両方で使う）。
' 検査  : 失敗印が無い、SPC・DATのCRCが記録と同じ、主出力・エラー記録が記録した形式で合格（ValidateReportFiles）、.udgがある、
'         D11が実績月数・D10/D16/D18が実績＋予測の月数で連続、D18を補った場合はその記録がある、
'         祝日変数を使った場合は、HOLが同じフォルダの祝日CSVの複製から作り直した値と一致する（deep=Trueのとき）。
' 補足  : 祝日の作り直しはSHA-256の計算を含み時間がかかるため、保存時と最終保存の検査だけで行う。
'         比較の取りまとめでは、照合一覧（CRC32）で保存後に変わっていないことを確かめる。
Private Sub ValidateTrialFolder(ByVal folder As String, ByVal meta As Object, Optional ByVal deep As Boolean = True)
    Dim stem As String, base As String, first As Long, actual As Long, forecast As Long, table As Variant, name As Variant
    If modMain.FileExists(modMain.CombinePath(folder, FAILED_MARK)) Then Err.Raise vbObjectError + 3403, , "保存に失敗したフォルダです：" & folder
    stem = CStr(meta("STEM"))
    base = modMain.CombinePath(folder, stem)
    first = modX13IO.MonthIndexFromText(CStr(meta("FIRST")))
    actual = CLng(meta("ACTUAL_MONTHS"))
    forecast = CLng(meta("FORECAST_MONTHS"))
    If actual <> modX13Options.ACTUAL_MONTHS_REQUIRED Then Err.Raise vbObjectError + 3404, , "保存された実績月数が" & actual & "です：" & folder
    If modX13IO.FileCRC(base & ".dat") <> CStr(meta("DATA_CRC")) Then Err.Raise vbObjectError + 3405, , "DATが記録と違います：" & folder
    If modX13IO.FileCRC(base & ".spc") <> CStr(meta("SPEC_CRC")) Then Err.Raise vbObjectError + 3406, , "SPCが記録と違います：" & folder
    ValidateReportFiles folder, stem, meta, deep
    If Not modMain.FileExists(base & ".udg") Then Err.Raise vbObjectError + 3408, , "UDGがありません：" & folder
    table = modX13IO.ReadMonthlyTable(base & ".d11", first, actual)
    For Each name In Array("d10", "d16", "d18")
        table = modX13IO.ReadMonthlyTable(base & "." & CStr(name), first, actual + forecast)
    Next name
    If CStr(meta("D18_ORIGIN")) = "VBA_NEUTRAL" And Not modMain.FileExists(modMain.CombinePath(folder, "_d18_origin.txt")) Then _
        Err.Raise vbObjectError + 3409, , "中立値のD18の記録（_d18_origin.txt）がありません：" & folder
    If deep And modX13Calendar.UsesHoliday(CStr(meta("CALENDAR"))) Then modHolidayIIP.VerifySnapshot folder
End Sub

' 目的  : 保存フォルダの主出力とエラー記録を、記録した出力形式で検査する（modX13Results.InspectReport）。
' 規則  : 記録した主出力・エラー記録の名前が、実際にあるファイルと同じであること。
'         HTML版は内部用テキスト（_report_text_from_html.txt）があること。deep=Trueなら、原本の .html から作り直した
'         テキストと一字一句同じであること（内部用テキストだけが原本と食い違っていないこと）。
' エラー: 合わなければ停止する。
Private Sub ValidateReportFiles(ByVal folder As String, ByVal stem As String, ByVal meta As Object, ByVal deep As Boolean)
    Dim expected As String, report As Object, textPath As String
    expected = TrialReportFormat(meta)
    Set report = modX13Results.InspectReport(folder, stem, expected, deep And expected = modX13Results.REPORT_HTML)
    If meta.Exists("REPORT_MAIN") Then
        If CStr(meta("REPORT_MAIN")) <> CStr(report("Main")) Or CStr(meta("REPORT_ERROR")) <> CStr(report("Error")) Then _
            Err.Raise vbObjectError + 3407, , "主出力・エラー記録のファイル名が記録と違います（" & report("Main") & "、" & report("Error") & "）：" & folder
    End If
    If expected = modX13Results.REPORT_HTML Then
        textPath = modMain.CombinePath(folder, modX13Results.REPORT_TEXT_FILE)
        If Not modMain.FileExists(textPath) Then Err.Raise vbObjectError + 3407, , _
            "HTML版の解析用テキスト（" & modX13Results.REPORT_TEXT_FILE & "）がありません：" & folder
        If deep Then
            If modX13IO.ReadText(textPath, "UTF8") <> CStr(report("Text")) Then Err.Raise vbObjectError + 3407, , _
                "HTML版の解析用テキストが、原本の " & report("Main") & " から作り直した内容と違います：" & folder
        End If
    End If
End Sub

' 目的  : 試行の記録（_trial.meta）から出力形式を返す。
'         REPORT_FORMAT が無い記録は、HTML対応前に保存したASCII版の試行として扱う。
' エラー: 記録の値が ASCII・HTML 以外なら停止する。
Public Function TrialReportFormat(ByVal meta As Object) As String
    TrialReportFormat = modX13Results.REPORT_ASCII
    If meta.Exists("REPORT_FORMAT") Then TrialReportFormat = CStr(meta("REPORT_FORMAT"))
    If TrialReportFormat <> modX13Results.REPORT_ASCII And TrialReportFormat <> modX13Results.REPORT_HTML Then _
        Err.Raise vbObjectError + 3407, "TrialReportFormat", "試行の記録の出力形式が不正です：" & TrialReportFormat
End Function

' 目的  : 試行IDから保存フォルダ（絶対パス）を返す。
' エラー: 試行_一覧に無い、または同じ試行IDが2行ある場合は停止する。
Public Function TrialFolder(ByVal trialId As String) As String
    Dim row As Long
    row = TrialRow(trialId)
    TrialFolder = modPackagePaths.ResolvePath(CStr(ThisWorkbook.Worksheets(modSheetNames.TrialSheet()).Cells(row, 7).Value2))
End Function

Private Function TrialRow(ByVal trialId As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(modSheetNames.TrialSheet())
    For r = 2 To modX13IO.NextRow(ws) - 1
        If CStr(ws.Cells(r, 1).Value2) = trialId Then
            If TrialRow > 0 Then Err.Raise vbObjectError + 3410, "TrialRow", "試行_一覧に同じ試行IDが2行あります：" & trialId
            TrialRow = r
        End If
    Next r
    If TrialRow = 0 Then Err.Raise vbObjectError + 3411, "TrialRow", "試行_一覧に試行IDがありません：" & trialId
End Function

' 目的  : 保存済みの試行を検査し、記録（meta）を返す。現在のシートの設定ではなく、保存したファイルを根拠にする。
' 入力  : deep=Trueなら祝日変数の作り直しまで行う（最終保存の前）。Falseは照合一覧と出力表の検査（取りまとめ）。
' 出力  : Dictionary（_trial.meta の内容＋ "FOLDER"＝絶対パス）
' エラー: 完了印が無い、照合一覧と合わない、記録の識別子が表と違う、出力表が欠けている場合は停止する。
Public Function ValidateTrial(ByVal trialId As String, Optional ByVal deep As Boolean = False) As Object
    Dim ws As Worksheet, row As Long, folder As String, meta As Object
    row = TrialRow(trialId)
    Set ws = ThisWorkbook.Worksheets(modSheetNames.TrialSheet())
    If CStr(ws.Cells(row, 10).Value2) <> StateSaved() Then Err.Raise vbObjectError + 3412, "ValidateTrial", "保存済の試行ではありません：" & trialId
    folder = modPackagePaths.ResolvePath(CStr(ws.Cells(row, 7).Value2))
    If modX13IO.ReadMarker(modMain.CombinePath(folder, TRIAL_COMPLETE)) <> TRIAL_MARKER Then _
        Err.Raise vbObjectError + 3413, "ValidateTrial", "試行の完了印がありません：" & folder
    VerifyManifest folder, TRIAL_MANIFEST
    Set meta = modX13IO.ReadMetaFile(modMain.CombinePath(folder, TRIAL_META))
    If meta("TRIAL") <> trialId Or meta("RUN") <> CStr(ws.Cells(row, 2).Value2) Or meta("SERIES") <> CStr(ws.Cells(row, 4).Value2) Or _
       meta("CANDIDATE") <> CStr(ws.Cells(row, 3).Value2) Then Err.Raise vbObjectError + 3414, "ValidateTrial", "試行の記録が試行_一覧と一致しません：" & trialId
    ValidateTrialFolder folder, meta, deep
    meta.Add "FOLDER", folder
    Set ValidateTrial = meta
End Function

' 目的: ある実行の保存済み試行を、保存した順に返す（要素は試行ID）。
Public Function TrialsOfRun(ByVal runId As String) As Collection
    Dim ws As Worksheet, r As Long, result As New Collection
    Set ws = modMain.WorksheetOrNothing(modSheetNames.TrialSheet())
    If Not ws Is Nothing Then
        For r = 2 To modX13IO.NextRow(ws) - 1
            If CStr(ws.Cells(r, 2).Value2) = runId And CStr(ws.Cells(r, 10).Value2) = StateSaved() Then result.Add CStr(ws.Cells(r, 1).Value2)
        Next r
    End If
    Set TrialsOfRun = result
End Function

' ================= 最終保存 =================

' 目的  : 決定した試行を最終結果として保存する。
' 入力  : trialId、decision（Dictionary：mode=自動/手動、metric、metricValue、group、rank、filteredRank、
'         requireCalendar、alpha、ties、note）
' 出力  : 保存ID（例 F_20261017_190000_1）
' 副作用: X13_FINAL_DIR\<系列ID>\<保存ID>\ を新しく作る（試行フォルダの全ファイル＋_final.meta）。
'         最終_決定・最終_保存履歴・最終_月次結果・最終_使用設定へ行を追加し、同じ系列の以前の採用を「過去の採用」にする。
' エラー: 途中で失敗した場合は、追加した行を消し、状態を戻し、フォルダに _failed.txt を書いてから停止する。
'         以前の最終結果のフォルダは消さない。
Public Function FinalizeTrial(ByVal trialId As String, ByVal decision As Object) As String
    Dim fso As Object, meta As Object, root As String, adoption As String, folder As String, stage As String, file As Object
    Dim finalMeta As Object, owned As Boolean, n As Long, item As Variant
    Dim monthly As Worksheet, settings As Worksheet, history As Worksheet, decisions As Worksheet
    Dim monthlyStart As Long, monthlyCount As Long, settingsStart As Long, settingsCount As Long
    Dim historyRow As Long, decisionRow As Long, previousRows As Collection
    Dim errorNumber As Long, errorText As String
    If mBusy Then Err.Raise vbObjectError + 3420, "FinalizeTrial", "別の保存処理を実行中です。"
    On Error GoTo Failed
    mBusy = True
    EnsureStoreSheets
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set meta = ValidateTrial(trialId, True)
    root = modPackagePaths.RequireFolder("X13_FINAL_DIR", True)
    Do
        n = n + 1
        adoption = "F_" & modMain.TimeStamp() & "_" & n
        folder = modMain.CombinePath(modMain.CombinePath(root, CStr(meta("SERIES"))), adoption)
        stage = folder & ".pending"
    Loop While fso.FolderExists(folder) Or fso.FolderExists(stage) Or _
               Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets(modSheetNames.AdoptionSheet()).Columns(1), adoption) > 0
    modPackagePaths.EnsureFolderTree fso.GetParentFolderName(folder)
    fso.CreateFolder stage
    owned = True

    For Each file In fso.GetFolder(CStr(meta("FOLDER"))).Files
        modX13IO.CopyFileVerified file.path, modMain.CombinePath(stage, file.Name)
    Next file
    VerifyManifest stage, TRIAL_MANIFEST
    Set finalMeta = CreateObject("Scripting.Dictionary")
    finalMeta.Add "ADOPTION", adoption
    finalMeta.Add "TRIAL", trialId
    finalMeta.Add "SERIES", CStr(meta("SERIES"))
    finalMeta.Add "SETTING", CStr(meta("SETTING"))
    finalMeta.Add "RUN", CStr(meta("RUN"))
    finalMeta.Add "DECISION_MODE", CStr(decision("mode"))
    finalMeta.Add "RANK_METRIC", CStr(decision("metric"))
    finalMeta.Add "METRIC_VALUE", CStr(decision("metricValue"))
    finalMeta.Add "COMPARISON_GROUP", CStr(decision("group"))
    finalMeta.Add "REQUIRE_CALENDAR_SIGNIFICANCE", CStr(decision("requireCalendar"))
    finalMeta.Add "ALPHA", CStr(decision("alpha"))
    finalMeta.Add "D18_ORIGIN", CStr(meta("D18_ORIGIN"))
    finalMeta.Add "REPORT_FORMAT", TrialReportFormat(meta)
    finalMeta.Add "SOURCE_TRIAL_FOLDER", modPackagePaths.RelativeToPackage(CStr(meta("FOLDER")))
    finalMeta.Add "DECIDED_UTC", modMain.UtcNowText()
    modX13IO.WriteMetaFile modMain.CombinePath(stage, FINAL_META), finalMeta
    WriteManifest stage, FINAL_MANIFEST
    VerifyManifest stage, FINAL_MANIFEST
    fso.MoveFolder stage, folder
    VerifyManifest folder, FINAL_MANIFEST
    ValidateTrialFolder folder, meta
    modX13IO.WriteNewText modMain.CombinePath(folder, FINAL_COMPLETE), FINAL_MARKER & vbCrLf

    ' 表への記録。失敗したらここで追加した行を消す。
    Set monthly = ThisWorkbook.Worksheets(modSheetNames.FinalMonthlySheet())
    modX13IO.AppendRows monthly, MonthlyRows(folder, meta, adoption), monthlyStart, monthlyCount, Array(5, 6, 7, 8, 9, 10)
    Set settings = ThisWorkbook.Worksheets(modSheetNames.FinalSettingsSheet())
    modX13IO.AppendRows settings, SettingsRows(folder, meta, adoption), settingsStart, settingsCount
    Set decisions = ThisWorkbook.Worksheets(modSheetNames.DecisionSheet())
    decisionRow = modX13IO.NextRow(decisions)
    decisions.Cells(decisionRow, 1).Resize(1, UBound(DecisionHeaders()) + 1).NumberFormat = "@"
    decisions.Cells(decisionRow, 1).Resize(1, UBound(DecisionHeaders()) + 1).Value2 = Array(adoption, trialId, meta("SERIES"), _
        decision("mode"), decision("metric"), decision("metricValue"), decision("group"), decision("rank"), decision("filteredRank"), _
        decision("requireCalendar"), decision("alpha"), decision("ties"), finalMeta("DECIDED_UTC"), decision("note"))
    Set history = ThisWorkbook.Worksheets(modSheetNames.AdoptionSheet())
    Set previousRows = CurrentRowsOfSeries(history, CStr(meta("SERIES")))
    historyRow = modX13IO.NextRow(history)
    history.Cells(historyRow, 1).Resize(1, UBound(AdoptionHeaders()) + 1).NumberFormat = "@"
    history.Cells(historyRow, 1).Resize(1, UBound(AdoptionHeaders()) + 1).Value2 = Array(adoption, trialId, meta("SERIES"), _
        meta("SETTING"), meta("RUN"), decision("mode"), modPackagePaths.RelativeToPackage(folder), finalMeta("DECIDED_UTC"), StateCurrent())
    For Each item In previousRows
        history.Cells(CLng(item), 9).Value2 = StatePrevious()
    Next item
    modLog.WriteLog "FinalizeTrial", "INFO", meta("SERIES") & " 保存ID=" & adoption & " 試行ID=" & trialId & " 決定=" & decision("mode")
    mBusy = False
    FinalizeTrial = adoption
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    ' 後始末：この処理で追加した行を消し、以前の採用の状態を戻す。後始末の失敗で元のエラーを隠さない。
    On Error Resume Next
    If monthlyCount > 0 Then monthly.Cells(monthlyStart, 1).Resize(monthlyCount, UBound(FinalMonthlyHeaders()) + 1).ClearContents
    If settingsCount > 0 Then settings.Cells(settingsStart, 1).Resize(settingsCount, UBound(FinalSettingsHeaders()) + 1).ClearContents
    If decisionRow > 0 Then decisions.Rows(decisionRow).ClearContents
    If historyRow > 0 Then history.Rows(historyRow).ClearContents
    If Not previousRows Is Nothing Then
        For Each item In previousRows
            history.Cells(CLng(item), 9).Value2 = StateCurrent()
        Next item
    End If
    If owned Then MarkFailure folder, stage, errorText
    mBusy = False
    On Error GoTo 0
    modLog.WriteLog "FinalizeTrial", "ERROR", trialId & " " & errorText
    Err.Raise errorNumber, "FinalizeTrial", errorText
End Function

Private Function CurrentRowsOfSeries(ByVal history As Worksheet, ByVal seriesId As String) As Collection
    Dim r As Long, result As New Collection
    For r = 2 To modX13IO.NextRow(history) - 1
        If CStr(history.Cells(r, 3).Value2) = seriesId And CStr(history.Cells(r, 9).Value2) = StateCurrent() Then result.Add r
    Next r
    Set CurrentRowsOfSeries = result
End Function

' 最終_月次結果の行。実績月は原値・計算用の値・D11を、予測月はD10・D16・D18だけを書く（予測の原値は無い）。
Private Function MonthlyRows(ByVal folder As String, ByVal meta As Object, ByVal adoption As String) As Variant
    Dim base As String, first As Long, actual As Long, forecast As Long, d10 As Variant, d11 As Variant, d16 As Variant, d18 As Variant
    Dim inputs As Object, data() As Variant, k As Long, origin As String
    base = modMain.CombinePath(folder, CStr(meta("STEM")))
    first = modX13IO.MonthIndexFromText(CStr(meta("FIRST")))
    actual = CLng(meta("ACTUAL_MONTHS"))
    forecast = CLng(meta("FORECAST_MONTHS"))
    d11 = modX13IO.ReadMonthlyTable(base & ".d11", first, actual)
    d10 = modX13IO.ReadMonthlyTable(base & ".d10", first, actual + forecast)
    d16 = modX13IO.ReadMonthlyTable(base & ".d16", first, actual + forecast)
    d18 = modX13IO.ReadMonthlyTable(base & ".d18", first, actual + forecast)
    Set inputs = InputValues(modMain.CombinePath(folder, "_input_values.tsv"), first, actual)
    origin = IIf(CStr(meta("D18_ORIGIN")) = "VBA_NEUTRAL", "VBAの中立値（カレンダー調整なしのためX-13はD18を出さない）", "X-13の出力")
    ReDim data(1 To actual + forecast, 1 To UBound(FinalMonthlyHeaders()) + 1)
    For k = 1 To actual + forecast
        data(k, 1) = adoption
        data(k, 2) = meta("SERIES")
        data(k, 3) = modX13IO.MonthText(first + k - 1)
        data(k, 4) = IIf(k <= actual, "実績", "予測")
        If k <= actual Then
            data(k, 5) = inputs(k)(0)
            data(k, 6) = inputs(k)(1)
            data(k, 7) = d11(k, 2)
        Else
            data(k, 5) = ""
            data(k, 6) = ""
            data(k, 7) = ""
        End If
        data(k, 8) = d10(k, 2)
        data(k, 9) = d16(k, 2)
        data(k, 10) = d18(k, 2)
        data(k, 11) = origin
        data(k, 12) = meta("TRIAL")
    Next k
    MonthlyRows = data
End Function

' _input_values.tsv を読み、月の位置 → Array(原値, 計算用の値) を返す。年月の並びが実績期間と一致しなければ停止する。
Private Function InputValues(ByVal path As String, ByVal first As Long, ByVal actual As Long) As Object
    Dim lines As Variant, fields As Variant, i As Long, result As Object
    Set result = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(modX13IO.ReadText(path, "UTF8"), vbCr, ""), vbLf)
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            fields = Split(lines(i), vbTab)
            If CStr(fields(0)) <> modX13IO.MonthText(first + result.Count) Then Err.Raise vbObjectError + 3421, "InputValues", _
                "_input_values.tsv の年月の並びが実績期間と一致しません：" & fields(0)
            result.Add result.Count + 1, Array(modX13IO.NumberValue(CStr(fields(1))), modX13IO.NumberValue(CStr(fields(2))))
        End If
    Next i
    If result.Count <> actual Then Err.Raise vbObjectError + 3422, "InputValues", "_input_values.tsv の行数が" & result.Count & "です（期待" & actual & "）"
    Set InputValues = result
End Function

' 最終_使用設定の行（_used_settings.tsv の各行）。
Private Function SettingsRows(ByVal folder As String, ByVal meta As Object, ByVal adoption As String) As Variant
    Dim lines As Variant, fields As Variant, i As Long, rows As New Collection
    lines = Split(Replace(modX13IO.ReadText(modMain.CombinePath(folder, "_used_settings.tsv"), "UTF8"), vbCr, ""), vbLf)
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            fields = Split(lines(i) & vbTab & vbTab & vbTab, vbTab)
            rows.Add Array(adoption, meta("SERIES"), fields(0), fields(1), fields(2), fields(3))
        End If
    Next i
    SettingsRows = modX13IO.RowsToArray(rows, UBound(FinalSettingsHeaders()) + 1)
End Function

' 目的  : 最終保存のフォルダを検査し、記録（_trial.meta＋_final.metaの値）を返す。
' エラー: 完了印・照合一覧・出力表・祝日変数のいずれかが合わなければ停止する。
Public Function ValidateFinal(ByVal adoptionId As String) As Object
    Dim ws As Worksheet, r As Long, row As Long, folder As String, meta As Object, finalMeta As Object, key As Variant
    Set ws = ThisWorkbook.Worksheets(modSheetNames.AdoptionSheet())
    For r = 2 To modX13IO.NextRow(ws) - 1
        If CStr(ws.Cells(r, 1).Value2) = adoptionId Then
            If row > 0 Then Err.Raise vbObjectError + 3423, "ValidateFinal", "最終_保存履歴に同じ保存IDが2行あります：" & adoptionId
            row = r
        End If
    Next r
    If row = 0 Then Err.Raise vbObjectError + 3424, "ValidateFinal", "最終_保存履歴に保存IDがありません：" & adoptionId
    folder = modPackagePaths.ResolvePath(CStr(ws.Cells(row, 7).Value2))
    If modX13IO.ReadMarker(modMain.CombinePath(folder, FINAL_COMPLETE)) <> FINAL_MARKER Then _
        Err.Raise vbObjectError + 3425, "ValidateFinal", "最終保存の完了印がありません：" & folder
    VerifyManifest folder, FINAL_MANIFEST
    Set meta = modX13IO.ReadMetaFile(modMain.CombinePath(folder, TRIAL_META))
    Set finalMeta = modX13IO.ReadMetaFile(modMain.CombinePath(folder, FINAL_META))
    If finalMeta("ADOPTION") <> adoptionId Or finalMeta("TRIAL") <> CStr(ws.Cells(row, 2).Value2) Or _
       meta("SERIES") <> CStr(ws.Cells(row, 3).Value2) Then Err.Raise vbObjectError + 3426, "ValidateFinal", "最終保存の記録が最終_保存履歴と一致しません：" & adoptionId
    ValidateTrialFolder folder, meta
    For Each key In finalMeta.Keys
        If Not meta.Exists("FINAL_" & CStr(key)) Then meta.Add "FINAL_" & CStr(key), finalMeta(key)
    Next key
    meta.Add "FOLDER", folder
    meta.Add "STATE", CStr(ws.Cells(row, 9).Value2)
    Set ValidateFinal = meta
End Function

' 目的  : 最終_月次結果の行が、保存フォルダのD11・D10・D16・D18と一致するか確かめる（保存・再読込み後の照合）。
' 出力  : 照合した行数（実績＋予測の月数）
' エラー: 行数・年月・値のいずれかが違えば停止する。
Public Function VerifyFinalAgainstSheet(ByVal adoptionId As String) As Long
    Dim meta As Object, expected As Variant, ws As Worksheet, r As Long, k As Long, c As Long
    Set meta = ValidateFinal(adoptionId)
    expected = MonthlyRows(CStr(meta("FOLDER")), meta, adoptionId)
    Set ws = ThisWorkbook.Worksheets(modSheetNames.FinalMonthlySheet())
    For r = 2 To modX13IO.NextRow(ws) - 1
        If CStr(ws.Cells(r, 1).Value2) = adoptionId Then
            k = k + 1
            If k > UBound(expected, 1) Then Err.Raise vbObjectError + 3427, "VerifyFinalAgainstSheet", "最終_月次結果の行が多すぎます：" & adoptionId
            For c = 1 To UBound(expected, 2)
                If Not SameCell(ws.Cells(r, c).Value2, expected(k, c)) Then Err.Raise vbObjectError + 3428, "VerifyFinalAgainstSheet", _
                    "最終_月次結果が保存ファイルと違います（" & ws.Cells(r, 3).Value2 & "、" & ws.Cells(1, c).Value2 & "）"
            Next c
        End If
    Next r
    If k <> UBound(expected, 1) Then Err.Raise vbObjectError + 3429, "VerifyFinalAgainstSheet", "最終_月次結果の行数が" & k & "です（期待" & UBound(expected, 1) & "）"
    VerifyFinalAgainstSheet = k
End Function

' セルの値と期待値が同じか。数値は相対1E-12の差まで同じとみなす（保存・再読込みでの最下位桁の違いを誤検出しないため）。
Private Function SameCell(ByVal actual As Variant, ByVal expected As Variant) As Boolean
    Dim magnitude As Double
    If IsError(actual) Then Exit Function
    If IsNumeric(actual) And IsNumeric(expected) And VarType(actual) <> vbString And VarType(expected) <> vbString Then
        magnitude = Abs(CDbl(expected))
        If magnitude < 1 Then magnitude = 1
        SameCell = (Abs(CDbl(actual) - CDbl(expected)) <= 0.000000000001 * magnitude)
    Else
        SameCell = (CStr(actual) = CStr(expected))
    End If
End Function

' 目的: 系列ごとの現在の採用（状態が採用中）を返す。キー=系列ID、値=保存ID。
Public Function CurrentAdoptions() As Object
    Dim ws As Worksheet, r As Long, result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set ws = modMain.WorksheetOrNothing(modSheetNames.AdoptionSheet())
    If Not ws Is Nothing Then
        For r = 2 To modX13IO.NextRow(ws) - 1
            If CStr(ws.Cells(r, 9).Value2) = StateCurrent() Then
                If result.Exists(CStr(ws.Cells(r, 3).Value2)) Then Err.Raise vbObjectError + 3430, "CurrentAdoptions", _
                    "同じ系列に採用中の結果が2つあります：" & ws.Cells(r, 3).Value2
                result.Add CStr(ws.Cells(r, 3).Value2), CStr(ws.Cells(r, 1).Value2)
            End If
        Next r
    End If
    Set CurrentAdoptions = result
End Function

' ================= 照合一覧 =================

' フォルダ内の全ファイル（照合一覧と完了印を除く）の名前・バイト数・CRC32を書く。
Private Sub WriteManifest(ByVal folder As String, ByVal name As String)
    Dim fso As Object, file As Object, text As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    text = "file" & vbTab & "bytes" & vbTab & "crc32" & vbCrLf
    For Each file In fso.GetFolder(folder).Files
        If Not IsBookkeeping(file.Name, name) Then text = text & file.Name & vbTab & CStr(file.Size) & vbTab & modX13IO.FileCRC(file.path) & vbCrLf
    Next file
    modX13IO.WriteNewText modMain.CombinePath(folder, name), text
End Sub

' 照合一覧と完了印は、照合の対象外（自分自身を含められないため）。最終保存では試行の照合一覧・完了印は対象に含める。
Private Function IsBookkeeping(ByVal fileName As String, ByVal manifestName As String) As Boolean
    If fileName = manifestName Then IsBookkeeping = True
    If manifestName = TRIAL_MANIFEST And fileName = TRIAL_COMPLETE Then IsBookkeeping = True
    If manifestName = FINAL_MANIFEST And fileName = FINAL_COMPLETE Then IsBookkeeping = True
End Function

' 照合一覧どおりのファイルがあり、一覧に無いファイルが無いことを確かめる。
Private Sub VerifyManifest(ByVal folder As String, ByVal name As String)
    Dim fso As Object, lines As Variant, fields As Variant, i As Long, seen As Object, file As Object, path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    lines = Split(Replace(modX13IO.ReadText(modMain.CombinePath(folder, name), "UTF8"), vbCr, ""), vbLf)
    If CStr(lines(0)) <> "file" & vbTab & "bytes" & vbTab & "crc32" Then Err.Raise vbObjectError + 3431, "VerifyManifest", "照合一覧の見出しが不正です：" & folder
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            fields = Split(lines(i), vbTab)
            If UBound(fields) <> 2 Then Err.Raise vbObjectError + 3432, "VerifyManifest", "照合一覧の行が不正です：" & folder
            If seen.Exists(fields(0)) Then Err.Raise vbObjectError + 3433, "VerifyManifest", "照合一覧に同じファイルが2回あります：" & fields(0)
            seen.Add fields(0), True
            path = modMain.CombinePath(folder, CStr(fields(0)))
            If Not fso.FileExists(path) Then Err.Raise vbObjectError + 3434, "VerifyManifest", "照合一覧のファイルがありません：" & path
            If CStr(fso.GetFile(path).Size) <> CStr(fields(1)) Or modX13IO.FileCRC(path) <> CStr(fields(2)) Then _
                Err.Raise vbObjectError + 3435, "VerifyManifest", "ファイルの内容が照合一覧と違います：" & path
        End If
    Next i
    For Each file In fso.GetFolder(folder).Files
        If Not IsBookkeeping(file.Name, name) And Not seen.Exists(file.Name) Then _
            Err.Raise vbObjectError + 3436, "VerifyManifest", "照合一覧に無いファイルがあります：" & file.Name
    Next file
End Sub

Private Sub MarkFailure(ByVal folder As String, ByVal stage As String, ByVal description As String)
    Dim target As String
    ' 後始末：失敗の記録だけを残す。記録できなくても元のエラーを優先する。
    On Error Resume Next
    If modMain.FolderExists(folder) Then target = folder
    If Len(target) = 0 And modMain.FolderExists(stage) Then target = stage
    If Len(target) > 0 Then modX13IO.WriteNewText modMain.CombinePath(target, FAILED_MARK), description & vbCrLf
    On Error GoTo 0
End Sub
