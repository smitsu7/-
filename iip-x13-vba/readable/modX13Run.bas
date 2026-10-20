Attribute VB_Name = "modX13Run"
Option Explicit

' 役割      : 季節調整の実行を最初から最後までつなぐ。
'               1) 系列ごとに原値を読み、設定を解決し、計算用の値（ゼロ置換後）を作る
'               2) 比較する候補（カレンダー × ARIMAの決め方 × 外れ値の扱い）を作る
'               3) 候補をまとめてX-13を1回起動し、各候補の出力を検査して、合格した候補を試行として保存する
'               4) 自動選択の探索で現れたARIMA次数を、外れ値の月と種類を固定して再推定する（任意）
'               5) 取りまとめ・順位・自動決定・系列別Excelを modX13Summary へ渡す
' 主な入口  : RunSelectedFromUI / RunOneFromUI / RunAllFromUI（画面）、RunSeries（試験・自動化）
' 依存先    : modPublicImport、modSeriesSettings、modZeroReplace、modX13Options、modX13Calendar、
'             modX13Interventions、modHolidayIIP、modX13Spec、modX13Native、modX13Results、
'             modX13Store、modX13Summary、modPackagePaths、modX13IO、modMain、modSheetNames、modUI、modLog
' 成功の判定: X-13の終了コードだけでは成功としない（入力エラーでも0を返した実例がある）。
'             各候補について、エラー記録にERRORが無いこと、主出力・.udgがあること、UDGの観測数が144か月であること、
'             D11が実績144か月、D10・D16・D18が実績＋予測の月数で、開始月から連続していることを確かめる。
'             カレンダー変数を使わない候補だけ、X-13が作らないD18を中立値で補い、その旨を記録する。
' 出力形式  : ControlのX13_EXECUTABLE_PATHに、ASCII版（x13as_ascii.exe）とHTML版（x13as_html.exe）のどちらを指定してもよい。
'             形式は実行ファイルの名前ではなく、実行後にX-13が書いたファイルで決める（modX13Results.BatchLogFile・InspectReport）。
'               ASCII版 … 主出力 <stem>.out、エラー記録 <stem>.err、バッチのログ jobs.log
'               HTML版  … 主出力 <stem>.html、エラー記録 <stem>_err.html、バッチのログ jobs_log.html
'             HTML版は、主出力から読んだ最終次数・情報量規準・予測誤差をUDGと照合し、自動選択の候補では
'             探索履歴が空でないことも確かめる。HTMLから作った解析用テキストは _report_text_from_html.txt に残す。
' 書く場所  : 試行_候補定義、試行_ジョブ、計算用_ゼロ置換記録、X13_WORK_DIR\<実行ID>_bNNN（作業領域）、
'             X13_CANDIDATE_DIR（modX13Store経由）

Private mBusy As Boolean

' 段階の名前（試行_候補定義の「段階」列）。
Private Function StageInitial() As String
    StageInitial = "初回"
End Function

Private Function StageRefit() As String
    StageRefit = "再推定"
End Function

' 試行_ジョブの「状態」。
Private Function StateReady() As String
    StateReady = "準備済"
End Function

Private Function StateSucceeded() As String
    StateSucceeded = "成功"
End Function

Private Function StateFailed() As String
    StateFailed = "失敗"
End Function

Private Function StatePreRunError() As String
    StatePreRunError = "実行前エラー"
End Function

' ================= 画面の入口 =================

' 目的: 設定_対象系列で「実行」をYESにした系列を計算する。
Public Sub RunSelectedFromUI()
    Dim ids As Collection
    On Error GoTo Failed
    Set ids = modSeriesSettings.SelectedSeriesIds()
    RunWithConfirmation ids
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "季節調整の実行"
End Sub

' 目的: 系列IDを1つ入力して計算する。
Public Sub RunOneFromUI()
    Dim id As String, ids As New Collection
    On Error GoTo Failed
    id = modUI.AskText("計算する系列IDを入力してください（例 industry_1000000000）。", "1系列の実行")
    If Len(id) = 0 Then Exit Sub
    If Not modPublicImport.IsIdent(id) Then Err.Raise vbObjectError + 3300, , "系列IDは半角英数字と下線です：" & id
    ids.Add id
    RunWithConfirmation ids
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "1系列の実行"
End Sub

' 目的: 取り込んだ全系列（原データ_系列一覧で状態がIMPORTED_の系列）を計算する（設定_対象系列のYES/NOは見ない）。
Public Sub RunAllFromUI()
    On Error GoTo Failed
    RunWithConfirmation AllReadySeriesIds()
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "全系列の実行"
End Sub

Private Sub RunWithConfirmation(ByVal ids As Collection)
    Dim runId As String
    If Not modUI.Confirm(ids.Count & "系列を、設定ID「" & SettingIdFromControl() & "」で計算します。" & vbLf & _
        "実績 " & modX13Options.PeriodLabel() & "、予測 " & modX13Options.ForecastMonths() & "か月。" & vbLf & _
        "X-13を外部プロセスとして起動します。続けますか？", "季節調整の実行") Then Exit Sub
    runId = RunSeries(ids)
    ThisWorkbook.Worksheets(modSheetNames.ComparisonSheet()).Activate
    modUI.ShowMessage "計算が終わりました。実行ID：" & runId & vbLf & _
        "候補の比較は「" & modSheetNames.ComparisonSheet() & "」、失敗した候補と理由は「" & _
        modSheetNames.ExecutionRecordSheet() & "」にあります。", vbInformation, "季節調整の実行"
End Sub

' 目的  : 「全系列の実行」の対象＝実際に取り込んだ系列の系列IDを返す（名前はr11初版のまま）。
'         判定は modPublicImport.ImportedSeriesIds に一本化している（X13_SERIES_SCOPE=ALL_INPUT と同じ対象）。
'         取込候補（IIP_取込候補）でREADYでも、取り込んでいない系列は含まない。
' エラー: 取込済みの系列が無ければ停止する。
Public Function AllReadySeriesIds() As Collection
    Dim result As Collection
    Set result = modPublicImport.ImportedSeriesIds()
    If result.Count = 0 Then Err.Raise vbObjectError + 3303, "AllReadySeriesIds", _
        "取り込んだ系列がありません。先に「" & modSheetNames.ImportPreviewSheet() & "」で系列をYESにして取り込んでください。"
    Set AllReadySeriesIds = result
End Function

Private Function SettingIdFromControl() As String
    SettingIdFromControl = modMain.GetControlValue("X13_SETTING_ID", "standard")
End Function

' ================= 実行の本体 =================

' 目的  : 指定した系列を計算し、候補を保存し、取りまとめまで行う。
' 入力  : ids（系列IDのCollection）、settingId（空欄ならControlのX13_SETTING_ID）
' 出力  : 実行ID（例 R_20261017_183000）
' 副作用: 試行_候補定義・試行_ジョブ・計算用_ゼロ置換記録へ追記、作業フォルダと候補フォルダへ書く。
'         取りまとめのシートを作り直し、系列別Excelを保存する。最終決定が「自動」なら最終保存も行う。
' エラー: 期間（144か月）・実行許可・X-13の場所・作業フォルダの不備は、何も計算する前に停止する。
'         系列ごとの入力不正（期間不足、0とlogの組合せ等）はその系列だけを「実行前エラー」として記録し、他の系列は続ける。
'         候補ごとのX-13の失敗は「失敗」として理由を残し、比較・最終決定に入れない。
Public Function RunSeries(ByVal ids As Collection, Optional ByVal settingId As String = "") As String
    Dim runId As String, first As Long, actual As Long, forecast As Long, raw As Object, id As Variant
    Dim contexts As Object, ctx As Object, reason As String, initialJobs As Collection, refitJobs As Collection
    Dim holiday As Object, holidayReason As String, needsHoliday As Boolean, job As Variant
    Dim errorNumber As Long, errorText As String, oldStatus As Variant

    If mBusy Then Err.Raise vbObjectError + 3304, "RunSeries", "別の季節調整を実行中です。"
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    mBusy = True
    If ids.Count = 0 Then Err.Raise vbObjectError + 3305, "RunSeries", "計算する系列がありません。"
    If Len(settingId) = 0 Then settingId = SettingIdFromControl()
    If Not modPublicImport.IsIdent(settingId) Then Err.Raise vbObjectError + 3306, "RunSeries", "設定IDは半角英数字と下線で指定してください：" & settingId

    ' ---- 実行前の共通検査（ここで止まれば何も書かない） ----
    first = modX13Options.FirstMonth()
    actual = modX13Options.ActualMonths()          ' 144か月でなければ停止
    forecast = modX13Options.ForecastMonths()
    RequireExecutionReady
    EnsureRunSheets
    runId = NewRunId()
    modLog.WriteLog "RunSeries", "INFO", runId & " 系列=" & ids.Count & " 設定ID=" & settingId & " 期間=" & modX13Options.PeriodLabel() & " 予測=" & forecast

    ' ---- 1) 系列ごとの準備 ----
    Set raw = modPublicImport.LoadRawValues(ids, first, actual)
    Set contexts = CreateObject("Scripting.Dictionary")
    For Each id In ids
        If Not contexts.Exists(CStr(id)) Then
            Application.StatusBar = "計算の準備：" & CStr(id)
            Set ctx = Nothing
            reason = PrepareSeries(CStr(id), settingId, raw(CStr(id)), first, actual, forecast, runId, ctx)
            If Len(reason) > 0 Then
                RecordSeriesError runId, CStr(id), settingId, reason
            Else
                contexts.Add CStr(id), ctx
            End If
        End If
    Next id

    ' ---- 2) 初回の候補 ----
    Set initialJobs = New Collection
    For Each id In contexts.Keys
        AddInitialJobs contexts(id), runId, initialJobs
    Next id
    For Each job In initialJobs
        If modX13Calendar.UsesHoliday(CStr(job("calendar"))) Then needsHoliday = True
    Next job
    If needsHoliday Then holidayReason = TryPrepareHoliday(first, actual, forecast, holiday)

    ' ---- 3) X-13の実行と検査・保存 ----
    ExecuteJobs initialJobs, contexts, holiday, holidayReason, runId

    ' ---- 4) 探索で現れた次数の再推定 ----
    If modMain.GetControlBoolean("X13_REFIT_SEARCH_MODELS", True) Then
        Set refitJobs = New Collection
        For Each id In contexts.Keys
            Set ctx = contexts(id)
            If ctx("modelMode") = "automdl" Then AddRefitJobs ctx, initialJobs, runId, refitJobs
        Next id
        If refitJobs.Count > 0 Then ExecuteJobs refitJobs, contexts, holiday, holidayReason, runId
    End If

    ' ---- 5) 取りまとめ ----
    Application.StatusBar = "候補を取りまとめています"
    modX13Summary.RefreshSummary runId
    If modMain.GetControlValue("X13_FINAL_MODE", "手動") = "自動" Then modX13Summary.AutoFinalize runId
    modX13Summary.ExportSummaries runId
    modLog.WriteLog "RunSeries", "INFO", runId & " 完了"
    mBusy = False
    Application.StatusBar = oldStatus
    RunSeries = runId
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    mBusy = False
    Application.StatusBar = oldStatus
    modLog.WriteLog "RunSeries", "ERROR", runId & " " & errorText
    Err.Raise errorNumber, "RunSeries", errorText
End Function

' 目的  : 実行の許可、X-13の実行ファイル、作業フォルダ（ASCIIのみ）を確かめる。
' エラー: どれかが不備なら停止する。
Private Sub RequireExecutionReady()
    Dim workRoot As String
    If Not modMain.GetControlBoolean("ALLOW_X13_EXECUTION", False) Then Err.Raise vbObjectError + 3307, "RequireExecutionReady", _
        "ControlのALLOW_X13_EXECUTIONがYESではありません。X-13の外部実行を確認したうえでYESにしてください。"
    modPackagePaths.RequireFile "X13_EXECUTABLE_PATH"
    workRoot = modPackagePaths.RequireFolder("X13_WORK_DIR", True)
    modPackagePaths.RequireAsciiPath workRoot, "X-13の作業フォルダ（X13_WORK_DIR）"
    modPackagePaths.RequireFolder "X13_CANDIDATE_DIR", True
End Sub

' 実行ID。秒まで同じ実行IDがあれば末尾に番号を付ける。
Private Function NewRunId() As String
    Dim base As String, n As Long, ws As Worksheet
    base = "R_" & modMain.TimeStamp()
    Set ws = ThisWorkbook.Worksheets(modSheetNames.JobSheet())
    NewRunId = base
    Do While Application.WorksheetFunction.CountIf(ws.Columns(1), NewRunId) > 0 Or _
             modMain.FolderExists(modMain.CombinePath(modPackagePaths.SettingPath("X13_CANDIDATE_DIR"), NewRunId))
        n = n + 1
        NewRunId = base & "_" & n
    Loop
End Function

' 目的  : 1系列の計算の準備（設定・計算用の値・候補の材料）。
' 出力  : 空文字なら成功（ctxに系列の文脈）。失敗なら理由の文字列（ctxはNothing）。
' 副作用: 計算用_ゼロ置換記録へ置換の記録（0件でも件数行）を書く。
Private Function PrepareSeries(ByVal seriesId As String, ByVal settingId As String, ByVal entry As Object, _
                               ByVal first As Long, ByVal actual As Long, ByVal forecast As Long, _
                               ByVal runId As String, ByRef ctx As Object) As String
    Dim cfg As Object, prepared As Object
    On Error GoTo Failed
    If Len(CStr(entry("missing"))) > 0 Then
        PrepareSeries = CStr(entry("missing"))
        Exit Function
    End If
    Set cfg = modSeriesSettings.ResolveAll(seriesId, settingId)
    Set prepared = modZeroReplace.PrepareComputationValues(cfg, entry("values"), first, modX13Options.DatDecimals())
    modZeroReplace.RecordReplacements runId, seriesId, settingId, prepared
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "id", seriesId
    ctx.Add "setting", settingId
    ctx.Add "cfg", cfg
    ctx.Add "prepared", prepared
    ctx.Add "sources", entry("sources")
    ctx.Add "first", first
    ctx.Add "actual", actual
    ctx.Add "forecast", forecast
    ctx.Add "modelMode", modX13Options.ModelMode(modSeriesSettings.SettingText(cfg, "X13_MODEL_MODE"))
    ctx.Add "arima", ""
    If ctx("modelMode") = "fixed" Then ctx("arima") = modX13Options.ArimaModelText(modSeriesSettings.SettingText(cfg, "X13_ARIMA_MODEL"))
    ctx.Add "outlierMode", modX13Options.OutlierMode(modSeriesSettings.SettingText(cfg, "X13_OUTLIER_MODE"))
    ' 指定外れ値は「自動＋指定」「指定のみ」の系列だけで使う。他の系列の行は読まない。
    ctx.Add "specified", ""
    If ctx("outlierMode") <> "自動" Then ctx("specified") = modX13Interventions.InterventionVariables(seriesId, settingId, first, first + actual - 1)
    ctx.Add "calendars", modX13Calendar.CandidateCalendars(cfg)
    Exit Function
Failed:
    Set ctx = Nothing
    PrepareSeries = Err.Description
End Function

' 祝日変数を準備する。失敗した場合は理由を返し、祝日を使う候補だけを実行前エラーにする。
Private Function TryPrepareHoliday(ByVal first As Long, ByVal actual As Long, ByVal forecast As Long, ByRef holiday As Object) As String
    On Error GoTo Failed
    Set holiday = modHolidayIIP.PrepareForRun(first, actual, forecast)
    Exit Function
Failed:
    Set holiday = Nothing
    TryPrepareHoliday = "祝日変数を準備できません：" & Err.Description
    modLog.WriteLog "TryPrepareHoliday", "ERROR", TryPrepareHoliday
End Function

' ================= 候補（ジョブ）の作成 =================

' 目的: 初回の候補を、系列のカレンダー候補の数だけ作る。
Private Sub AddInitialJobs(ByVal ctx As Object, ByVal runId As String, ByVal jobs As Collection)
    Dim calendar As Variant, job As Object
    For Each calendar In ctx("calendars")
        Set job = NewJob(runId, ctx, StageInitial(), CStr(calendar), CStr(ctx("modelMode")), CStr(ctx("arima")), _
            CStr(ctx("outlierMode")), CStr(ctx("specified")), "", "", _
            IIf(ctx("modelMode") = "automdl", "自動選択（automdl）の初回", "次数指定の計算"))
        jobs.Add job
    Next calendar
End Sub

' 目的  : 探索で現れたARIMA次数を、各カレンダー条件で再推定する候補を作る。
' 規則  : 基準試行 = その系列の初回の候補のうち、最初に成功した試行。
'         基準試行で自動検出された外れ値の月と種類を固定し（係数は固定しない）、自動検出はしない（指定のみ）。
'         指定外れ値（設定_外れ値指定）がある系列は、それも同じく加える。
'         次数は、成功した初回の試行の探索履歴（推定できた次数）と最終次数から重複なく集める。
'         探索途中のBIC2は再推定の候補を集めるためだけに使い、順位には使わない。
Private Sub AddRefitJobs(ByVal ctx As Object, ByVal initialJobs As Collection, ByVal runId As String, ByVal jobs As Collection)
    Dim job As Variant, reference As Object, folder As String, orders As Object, item As Variant
    Dim model As String, frozen As String, calendar As Variant, regressors As String, evidence As Object, created As Long
    Set orders = CreateObject("Scripting.Dictionary")
    For Each job In initialJobs
        If job("series") = ctx("id") And job("state") = StateSucceeded() Then
            If reference Is Nothing Then Set reference = job
            folder = modX13Store.TrialFolder(CStr(job("trial")))
            ' 主出力の解析用テキスト（ASCII版は .out、HTML版は保存済みの内部用テキスト）から探索履歴を読む。
            For Each item In modX13Results.ParseSearchHistory(modX13Results.ReportText(folder, CStr(job("stem"))))
                If item("State") = "推定済" Then AddModel orders, CStr(item("Model"))
            Next item
            Set evidence = modX13Results.ReadOutputEvidence(folder, CStr(job("stem")))
            If evidence.Exists("X13_ACTUAL_MODEL") Then AddModel orders, CStr(evidence("X13_ACTUAL_MODEL"))
        End If
    Next job
    If reference Is Nothing Then
        modLog.WriteLog "AddRefitJobs", "SKIP", ctx("id") & "：初回に成功した試行が無いため再推定しません。"
        Exit Sub
    End If
    folder = modX13Store.TrialFolder(CStr(reference("trial")))
    frozen = modX13Interventions.FrozenOutliers(modX13Results.ReadUdgValues(modMain.CombinePath(folder, reference("stem") & ".udg")))
    regressors = modX13Interventions.MergeRegressors(frozen, CStr(ctx("specified")))
    For Each item In orders.Keys
        For Each calendar In ctx("calendars")
            jobs.Add NewJob(runId, ctx, StageRefit(), CStr(calendar), "fixed", CStr(item), "指定のみ", CStr(ctx("specified")), frozen, _
                CStr(reference("trial")), "探索次数の再推定（外れ値の月と種類を基準試行 " & reference("trial") & " に固定）")
            created = created + 1
        Next calendar
    Next item
    modLog.WriteLog "AddRefitJobs", "INFO", ctx("id") & "：次数" & orders.Count & "×カレンダー" & ctx("calendars").Count & "＝" & created & "候補。固定した外れ値：" & IIf(Len(frozen) > 0, frozen, "なし")
End Sub

' 次数を検査して重複なく加える。季節差分が2以上等、次数指定で使えない次数は記録して除く。
Private Sub AddModel(ByVal orders As Object, ByVal rawModel As String)
    Dim model As String
    On Error GoTo Rejected
    model = modX13Options.ArimaModelText(rawModel)
    If Not orders.Exists(model) Then orders.Add model, True
    Exit Sub
Rejected:
    modLog.WriteLog "AddModel", "SKIP", "再推定に使えない次数を除きました：" & rawModel & "（" & Err.Description & "）"
End Sub

' 目的  : 候補を1つ作り、試行_候補定義へ記録する。
' 出力  : Dictionary（modX13Spec.WriteJobFiles が読むキーを含む）
Private Function NewJob(ByVal runId As String, ByVal ctx As Object, ByVal stage As String, ByVal calendar As String, _
                        ByVal modelMode As String, ByVal arima As String, ByVal outlierMode As String, _
                        ByVal specified As String, ByVal frozen As String, ByVal reference As String, ByVal origin As String) As Object
    Dim job As Object, ws As Worksheet, row As Long
    Set job = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(modSheetNames.CandidateSheet())
    row = modX13IO.NextRow(ws)
    job.Add "run", runId
    job.Add "candidate", "C" & Format$(CountRunRows(ws, runId) + 1, "00000")
    job.Add "series", ctx("id")
    job.Add "setting", ctx("setting")
    job.Add "stage", stage
    job.Add "calendar", calendar
    job.Add "modelMode", modelMode
    job.Add "arima", arima
    job.Add "outlierMode", outlierMode
    job.Add "specified", specified
    job.Add "frozen", frozen
    job.Add "regressors", modX13Interventions.MergeRegressors(frozen, specified)
    job.Add "reference", reference
    job.Add "origin", origin
    job.Add "state", ""
    job.Add "reason", ""
    job.Add "trial", ""
    job.Add "stem", ""
    job.Add "batch", ""
    job.Add "exitCode", ""
    job.Add "jobRow", 0
    ' 出力形式の記録（RunChunk・CheckJobOutputsが書き、modX13Store.SaveTrialが_trial.metaへ写す）。
    job.Add "engine", ""          ' 起動した実行ファイルの名前（記録だけ。形式の判定には使わない）
    job.Add "engineBytes", ""     ' その実行ファイルのバイト数
    job.Add "batchLog", ""        ' バッチのログの名前（jobs.log／jobs_log.html）
    job.Add "batchFormat", ""     ' バッチのログから決めた形式（ASCII／HTML）
    job.Add "reportFormat", ""    ' この候補の出力ファイルから決めた形式（ASCII／HTML）
    ws.Cells(row, 1).Resize(1, 14).NumberFormat = "@"
    ws.Cells(row, 1).Resize(1, 14).Value2 = Array(runId, job("candidate"), job("series"), job("setting"), stage, calendar, _
        modX13Calendar.CalendarLabel(calendar), IIf(modelMode = "automdl", "自動選択", "次数指定"), arima, outlierMode, _
        specified, frozen, reference, origin)
    Set NewJob = job
End Function

Private Function CountRunRows(ByVal ws As Worksheet, ByVal runId As String) As Long
    CountRunRows = Application.WorksheetFunction.CountIf(ws.Columns(1), runId)
End Function

' ================= X-13の実行 =================

' 目的  : 候補を X13_COMPARE_CHUNK 件ずつのバッチに分けて実行する。系列数・候補数の固定上限はない。
Private Sub ExecuteJobs(ByVal jobs As Collection, ByVal contexts As Object, ByVal holiday As Object, _
                        ByVal holidayReason As String, ByVal runId As String)
    Dim size As Long, chunk As Collection, i As Long, batchNo As Long
    size = modX13Options.IntegerSetting("X13_COMPARE_CHUNK", 1, 500)
    batchNo = CountBatches(runId)
    Set chunk = New Collection
    For i = 1 To jobs.Count
        chunk.Add jobs(i)
        If chunk.Count >= size Or i = jobs.Count Then
            batchNo = batchNo + 1
            Application.StatusBar = "X-13を実行しています：バッチ" & batchNo & "（" & i & "／" & jobs.Count & "候補）"
            RunChunk chunk, contexts, holiday, holidayReason, runId, batchNo
            Set chunk = New Collection
        End If
    Next i
End Sub

Private Function CountBatches(ByVal runId As String) As Long
    Dim fso As Object, folder As Object, root As String
    root = modPackagePaths.RequireFolder("X13_WORK_DIR", True)
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each folder In fso.GetFolder(root).SubFolders
        If Left$(folder.Name, Len(runId) + 2) = runId & "_b" Then CountBatches = CountBatches + 1
    Next folder
End Function

' 目的  : 1つのバッチを準備・実行・検査する。
' 副作用: 作業フォルダ <X13_WORK_DIR>\<実行ID>_bNNN を新しく作り、各候補のジョブフォルダ・jobs.mtaを書く。
'         合格した候補を modX13Store.SaveTrial で保存し、試行_ジョブへ状態を書く。
' エラー: X-13を起動できない（実行ファイル・構造体・許可の不備）場合は、バッチの全候補を失敗として記録してから停止する。
Private Sub RunChunk(ByVal jobs As Collection, ByVal contexts As Object, ByVal holiday As Object, _
                     ByVal holidayReason As String, ByVal runId As String, ByVal batchNo As Long)
    Dim fso As Object, workRoot As String, batchFolder As String, batchId As String, executable As String
    Dim item As Variant, job As Object, n As Long, stem As String, jobFolder As String, reason As String
    Dim mta As String, prepared As Long, exitCode As Long, status As String, errorText As String, batchLog As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    workRoot = modPackagePaths.RequireFolder("X13_WORK_DIR", True)
    modPackagePaths.RequireAsciiPath workRoot, "X-13の作業フォルダ（X13_WORK_DIR）"
    executable = modPackagePaths.RequireFile("X13_EXECUTABLE_PATH")
    batchId = runId & "_b" & Format$(batchNo, "000")
    batchFolder = modMain.CombinePath(workRoot, batchId)
    If fso.FolderExists(batchFolder) Then Err.Raise vbObjectError + 3308, "RunChunk", "作業フォルダが既にあります：" & batchFolder
    fso.CreateFolder batchFolder

    For Each item In jobs
        Set job = item
        n = n + 1
        stem = "j" & Format$(n, "000000")
        jobFolder = modMain.CombinePath(batchFolder, stem)
        fso.CreateFolder jobFolder
        job("stem") = stem
        job("batch") = batchId
        reason = ""
        If modX13Calendar.UsesHoliday(CStr(job("calendar"))) And Len(holidayReason) > 0 Then reason = holidayReason
        If Len(reason) = 0 Then reason = TryWriteJob(job, contexts(CStr(job("series"))), holiday, jobFolder, stem)
        If Len(reason) > 0 Then
            job("state") = StatePreRunError()
            job("reason") = reason
        Else
            job("state") = StateReady()
            mta = mta & Chr$(34) & stem & "/" & stem & Chr$(34) & vbCrLf
            prepared = prepared + 1
        End If
        WriteJobRow job, batchFolder
    Next item
    If prepared = 0 Then Exit Sub

    modX13IO.WriteNewText modMain.CombinePath(batchFolder, "jobs.mta"), mta
    modX13IO.RequireAsciiFile modMain.CombinePath(batchFolder, "jobs.mta")
    exitCode = RunEngine(executable, batchFolder, errorText)
    If Len(errorText) > 0 Then
        FailPrepared jobs, "X-13を起動できませんでした：" & errorText, ""
        Err.Raise vbObjectError + 3309, "RunChunk", "X-13を起動できませんでした：" & errorText
    End If
    status = modX13Native.ReadStatus(batchFolder)
    If status <> "COMPLETED" Then
        FailPrepared jobs, "X-13の実行が完了していません（状態 " & status & "、終了コード " & exitCode & "）", CStr(exitCode)
        Exit Sub
    End If
    ' バッチのログ（ASCII版 jobs.log／HTML版 jobs_log.html）から、このバッチで動いたX-13の出力形式を決める。
    batchLog = TryBatchLog(batchFolder, reason)
    If Len(batchLog) = 0 Then
        FailPrepared jobs, reason & "（終了コード " & exitCode & "）", CStr(exitCode)
        Exit Sub
    End If
    For Each item In jobs
        Set job = item
        If job("state") = StateReady() Then
            job("exitCode") = CStr(exitCode)
            job("engine") = fso.GetFileName(executable)
            job("engineBytes") = CStr(fso.GetFile(executable).Size)
            job("batchLog") = batchLog
            job("batchFormat") = modX13Results.BatchLogFormat(batchLog)
            jobFolder = modMain.CombinePath(batchFolder, CStr(job("stem")))
            reason = CheckJobOutputs(job, contexts(CStr(job("series"))), jobFolder)
            If Len(reason) = 0 Then reason = TrySaveTrial(job, jobFolder, batchFolder)
            If Len(reason) > 0 Then
                job("state") = StateFailed()
                job("reason") = reason
            Else
                job("state") = StateSucceeded()
            End If
            WriteJobRow job, batchFolder
        End If
    Next item
    modLog.WriteLog "RunChunk", "INFO", batchId & " 候補=" & jobs.Count & " 実行=" & prepared & " 終了コード=" & exitCode
End Sub

' バッチのログの名前を返す。無い・2つの形式がある場合は空文字を返し、reasonに理由を入れる（呼出側が全候補を失敗にする）。
Private Function TryBatchLog(ByVal batchFolder As String, ByRef reason As String) As String
    On Error GoTo Failed
    TryBatchLog = modX13Results.BatchLogFile(batchFolder)
    Exit Function
Failed:
    TryBatchLog = ""
    reason = Err.Description
End Function

' X-13を起動する。起動処理自体のエラーは理由の文字列で返す（呼出側が全候補を失敗にしてから停止するため）。
Private Function RunEngine(ByVal executable As String, ByVal batchFolder As String, ByRef errorText As String) As Long
    On Error GoTo Failed
    RunEngine = modX13Native.RunBatch(executable, batchFolder, modX13Options.IntegerSetting("X13_BATCH_TIMEOUT_SECONDS", 1, 3600))
    Exit Function
Failed:
    errorText = Err.Description
    RunEngine = -1
End Function

Private Function TryWriteJob(ByVal job As Object, ByVal ctx As Object, ByVal holiday As Object, _
                             ByVal jobFolder As String, ByVal stem As String) As String
    On Error GoTo Failed
    modX13Spec.WriteJobFiles job, ctx, holiday, jobFolder, stem
    Exit Function
Failed:
    TryWriteJob = "SPC・DATを作れません：" & Err.Description
End Function

Private Function TrySaveTrial(ByVal job As Object, ByVal jobFolder As String, ByVal batchFolder As String) As String
    On Error GoTo Failed
    job("trial") = modX13Store.SaveTrial(job, jobFolder, batchFolder)
    Exit Function
Failed:
    TrySaveTrial = "試行を保存できません：" & Err.Description
End Function

Private Sub FailPrepared(ByVal jobs As Collection, ByVal reason As String, ByVal exitCode As String)
    Dim item As Variant, job As Object, batchFolder As String
    For Each item In jobs
        Set job = item
        If job("state") = StateReady() Then
            job("state") = StateFailed()
            job("reason") = reason
            job("exitCode") = exitCode
            batchFolder = modMain.CombinePath(modPackagePaths.SettingPath("X13_WORK_DIR"), CStr(job("batch")))
            WriteJobRow job, batchFolder
        End If
    Next item
End Sub

' 目的  : 1候補のX-13出力を検査する。終了コードは見ない（0でも失敗した実例があるため）。
' 出力  : 空文字なら合格。不合格なら最初に見つかった理由。
' 検査  : 1) 主出力とエラー記録（modX13Results.InspectReport）：形式がバッチのログと同じ、エラー記録にERRORが無い、
'            主出力が空でない（HTML版は </html> で終わる）。ASCII版 <stem>.out・.err、HTML版 <stem>.html・_err.html
'         2) <stem>.udg があり、nobs が実績月数、予測ありなら nfcst が予測月数
'         3) D11 が実績月数、D10・D16 が実績＋予測の月数で、開始月から連続
'         4) D18：カレンダー変数・祝日変数を使った候補はX-13の出力が必要。使わない候補だけ中立値を補う
'            （modX13Results.EnsureD18。補った場合は _d18_origin.txt に記録）。補った後に月数を確かめる。
'         5) HTML版だけ：主出力から読んだ最終次数・情報量規準・予測誤差がUDGと合い、自動選択の候補は探索履歴が空でない
'            （modX13Results.HtmlEvidenceProblem）。合格したら、HTMLから作った解析用テキストを
'            _report_text_from_html.txt としてジョブフォルダに書く（原本の <stem>.html はそのまま残す）。
Private Function CheckJobOutputs(ByVal job As Object, ByVal ctx As Object, ByVal jobFolder As String) As String
    Dim base As String, stem As String, values As Object, first As Long, actual As Long, forecast As Long
    Dim specText As String, table As Variant, name As Variant, report As Object, problem As String
    On Error GoTo Failed
    stem = CStr(job("stem"))
    base = modMain.CombinePath(jobFolder, stem)
    first = ctx("first"): actual = ctx("actual"): forecast = ctx("forecast")
    Set report = modX13Results.InspectReport(jobFolder, stem, CStr(job("batchFormat")), True)
    job("reportFormat") = report("Format")
    If Not modMain.FileExists(base & ".udg") Then Err.Raise vbObjectError + 3314, , "X-13の診断要約（.udg）がありません（入力エラーで停止した可能性）"
    Set values = modX13Results.ReadUdgValues(base & ".udg")
    If modX13Results.ReadValue(values, "nobs") <> CStr(actual) Then Err.Raise vbObjectError + 3315, , _
        "UDGの観測数（nobs=" & modX13Results.ReadValue(values, "nobs") & "）が実績" & actual & "か月と違います"
    If forecast > 0 And modX13Results.ReadValue(values, "nfcst") <> CStr(forecast) Then Err.Raise vbObjectError + 3316, , _
        "UDGの予測月数（nfcst=" & modX13Results.ReadValue(values, "nfcst") & "）が設定の" & forecast & "か月と違います"
    table = modX13IO.ReadMonthlyTable(base & ".d11", first, actual)
    For Each name In Array("d10", "d16")
        table = modX13IO.ReadMonthlyTable(base & "." & CStr(name), first, actual + forecast)
    Next name
    specText = modX13IO.ReadText(base & ".spc", "UTF8")
    modX13Results.EnsureD18 base, specText, first, actual + forecast
    table = modX13IO.ReadMonthlyTable(base & ".d18", first, actual + forecast)
    If report("Format") = modX13Results.REPORT_HTML Then
        problem = modX13Results.HtmlEvidenceProblem(CStr(report("Text")), specText, values, CStr(job("modelMode")) = "automdl")
        If Len(problem) > 0 Then Err.Raise vbObjectError + 3317, , problem
        modX13IO.WriteNewText modMain.CombinePath(jobFolder, modX13Results.REPORT_TEXT_FILE), CStr(report("Text"))
    End If
    Exit Function
Failed:
    CheckJobOutputs = Err.Description
End Function

' ================= 記録 =================

' 目的: 試行_ジョブへ候補の状態を書く（初回は追記、2回目以降は同じ行を更新）。
Private Sub WriteJobRow(ByVal job As Object, ByVal batchFolder As String)
    Dim ws As Worksheet, row As Long
    Set ws = ThisWorkbook.Worksheets(modSheetNames.JobSheet())
    row = CLng(job("jobRow"))
    If row = 0 Then
        row = modX13IO.NextRow(ws)
        job("jobRow") = row
    End If
    ws.Cells(row, 1).Resize(1, UBound(JobHeaders()) + 1).NumberFormat = "@"
    ws.Cells(row, 1).Resize(1, UBound(JobHeaders()) + 1).Value2 = Array(job("run"), job("candidate"), job("series"), job("setting"), _
        job("stage"), job("batch"), job("stem"), modPackagePaths.RelativeToPackage(modMain.CombinePath(batchFolder, CStr(job("stem")))), _
        job("state"), job("reason"), job("exitCode"), job("trial"), modMain.UtcNowText())
End Sub

' 系列単位の実行前エラー（期間不足、ゼロとlogの組合せ、設定の誤り等）を試行_ジョブへ記録する。
Private Sub RecordSeriesError(ByVal runId As String, ByVal seriesId As String, ByVal settingId As String, ByVal reason As String)
    Dim ws As Worksheet, row As Long
    Set ws = ThisWorkbook.Worksheets(modSheetNames.JobSheet())
    row = modX13IO.NextRow(ws)
    ws.Cells(row, 1).Resize(1, UBound(JobHeaders()) + 1).NumberFormat = "@"
    ws.Cells(row, 1).Resize(1, UBound(JobHeaders()) + 1).Value2 = Array(runId, "（系列）", seriesId, settingId, "", "", "", "", _
        StatePreRunError(), reason, "", "", modMain.UtcNowText())
    modLog.WriteLog "RunSeries", "SKIP", seriesId & "：" & reason
End Sub

Public Function CandidateHeaders() As Variant
    CandidateHeaders = Array("実行ID", "候補ID", "系列ID", "設定ID", "段階", "カレンダー変数", "カレンダー", "ARIMAの決め方", _
        "ARIMA次数", "外れ値の扱い", "指定外れ値", "固定した自動検出外れ値", "基準試行ID", "由来")
End Function

Public Function JobHeaders() As Variant
    JobHeaders = Array("実行ID", "候補ID", "系列ID", "設定ID", "段階", "バッチ", "ジョブ名", "作業フォルダ", "状態", "理由", _
        "X-13終了コード", "試行ID", "記録日時UTC")
End Function

' 目的: 実行に使うシートを見出し付きで用意する（既存の見出しが違えば停止）。
Public Sub EnsureRunSheets()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.CandidateSheet()), CandidateHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.JobSheet()), JobHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.ZeroLogSheet()), modZeroReplace.ZeroLogHeaders()
    modX13Store.EnsureStoreSheets
End Sub
