Attribute VB_Name = "modSetup"
Option Explicit

' 役割      : 配布ブックの初期シートを作る（ホーム、Control、取込設定、設定表、結果の表の見出し）。
'             ビルド手順（05_source\build）が新しいブックへモジュールを取り込んだ後に1回呼ぶ。
'             既にあるシートの値は変えない（見出しが違うシートがあれば停止する）。
' 主な入口  : BuildWorkbook / BuildHome / AbsolutePathSettings
' 依存先    : modX13Options、modPublicImport、modSeriesSettings、modX13Interventions、modX13Run、modX13Summary、
'             modX13Provisional、modUI、modMain、modSheetNames、modX13IO、modLog
' 補足      : X-13の実行・ネット取得・取込は行わない（利用者がボタンで始める）。

' 目的  : 全ての初期シートを作り、ホームを先頭に置く。
' 副作用: シートを追加し、空の既定シート（Sheet1等）を削除する。
Public Sub BuildWorkbook()
    Dim oldScreen As Boolean, ws As Worksheet, keep As Object, name As Variant, i As Long
    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    modMain.EnsureWorksheet modSheetNames.HomeSheet()
    modX13Options.SetupControl
    modPublicImport.SetupImportSheets
    modSeriesSettings.RefreshSelectionSheet
    modSeriesSettings.SetupSeriesSettingsSheet
    modX13Interventions.SetupInterventionSheet
    modX13Run.EnsureRunSheets
    modX13Summary.EnsureSummarySheets
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.ProvisionalSheet()), modX13Provisional.ProvisionalHeaders()
    modLog.WriteLog "BuildWorkbook", "INFO", "初期シートを作成しました"
    BuildHome

    ' ブック作成時の空の既定シートを削除する（このツールのシート名でないもの、かつ空のものだけ）。
    Set keep = CreateObject("Scripting.Dictionary")
    For Each name In SheetOrder()
        keep(CStr(name)) = True
    Next name
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not keep.Exists(ws.Name) And Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
    Next i
    Application.DisplayAlerts = True
    ArrangeSheets
    ThisWorkbook.Worksheets(modSheetNames.HomeSheet()).Activate
    Application.ScreenUpdating = oldScreen
End Sub

' 表示の順序（作業の流れの順）。
Private Function SheetOrder() As Variant
    SheetOrder = Array(modSheetNames.HomeSheet(), modSheetNames.ControlSheet(), modSheetNames.ImportConfigSheet(), _
        modSheetNames.ImportPreviewSheet(), modSheetNames.SeriesListSheet(), modSheetNames.RawIndexSheet(), modSheetNames.ProvenanceSheet(), _
        modSheetNames.HolidayCalendarSheet(), modSheetNames.HolidayCalculationSheet(), modSheetNames.ZeroLogSheet(), _
        modSheetNames.SeriesSelectionSheet(), modSheetNames.SeriesSettingsSheet(), modSheetNames.InterventionSheet(), _
        modSheetNames.CandidateSheet(), modSheetNames.JobSheet(), modSheetNames.TrialSheet(), modSheetNames.ComparisonSheet(), _
        modSheetNames.RegressionSheet(), modSheetNames.SearchHistorySheet(), modSheetNames.DetectedOutlierSheet(), _
        modSheetNames.ExecutionRecordSheet(), modSheetNames.DecisionSheet(), modSheetNames.AdoptionSheet(), _
        modSheetNames.FinalMonthlySheet(), modSheetNames.FinalSettingsSheet(), modSheetNames.ProvisionalSheet(), modSheetNames.LogSheet())
End Function

Private Sub ArrangeSheets()
    Dim name As Variant, ws As Worksheet, previous As Worksheet
    For Each name In SheetOrder()
        Set ws = modMain.WorksheetOrNothing(CStr(name))
        If Not ws Is Nothing Then
            If previous Is Nothing Then
                ws.Move Before:=ThisWorkbook.Worksheets(1)
            Else
                ws.Move After:=previous
            End If
            Set previous = ws
        End If
    Next name
End Sub

' 目的  : ホーム（季節調整）シートに、作業の順序・ボタン・注意を書く。
' 副作用: ホームのセルとボタンを作り直す。
Public Sub BuildHome()
    Dim ws As Worksheet, lines As Variant, i As Long, row As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.HomeSheet())
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 90
    ws.Range("B1").Value2 = "鉱工業指数 X-13 季節調整ツール r11"
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B2").Value2 = "実績期間はControlの開始年月～終了年月です（初期値 2014年1月～2025年12月の144か月。96か月・120か月等も指定可）。期間内の全月に原指数がある系列だけを計算し、足りない月は埋めません。予測は実績の外側に別に数えます（初期値12か月）。季節調整はX-11です。"
    ws.Range("B3").Value2 = "このブックは IIP_X13_Package フォルダの直下に置いて使います。エンジン・原指数・祝日・結果の場所は、このブックのフォルダからの相対パスです（Controlシート）。"

    row = 5
    HomeStep ws, row, "1. 原指数の取込", "公表Excel（生産・出荷・在庫・在庫率・能力・稼働率）を直接読みます。プレビューで系列ごとの月数を確かめてから取り込みます。"
    modUI.AddButton ws, "btnHome_bundled", "同梱の公表Excelを指定", "modPublicImport.UseBundledPublicExcel", ws.Cells(row + 1, 2), 180
    modUI.AddButton ws, "btnHome_choose", "取込元ファイルの選択", "modPublicImport.ChooseFiles", ws.Cells(row + 1, 3), 180
    modUI.AddButton ws, "btnHome_preview", "プレビュー作成", "modPublicImport.PreviewFromUI", ws.Cells(row + 2, 2), 180
    modUI.AddButton ws, "btnHome_commit", "選択系列の取込", "modPublicImport.CommitFromUI", ws.Cells(row + 2, 3), 180

    row = 9
    HomeStep ws, row, "2. 祝日変数", "内閣府「国民の祝日」CSVから、経産省IIP方式の1列（A=月～金の祝日数、B=実績期間の同じ暦月の平均、A−B）を作ります。"
    modUI.AddButton ws, "btnHome_holidayDownload", "公式URLからCSVを取得", "modHolidayIIP.DownloadOfficialCsv", ws.Cells(row + 1, 2), 180
    modUI.AddButton ws, "btnHome_holidayLocal", "手元のCSVを選択", "modHolidayIIP.ChooseLocalCsv", ws.Cells(row + 1, 3), 180
    modUI.AddButton ws, "btnHome_holidayBuild", "祝日変数の作成と確認", "modHolidayIIP.BuildFromUI", ws.Cells(row + 2, 2), 180

    row = 13
    HomeStep ws, row, "3. 設定", "Controlで全体の設定、設定_系列別で系列ごとの設定、設定_外れ値指定で指定の外れ値を書きます。X-13の実行前に ALLOW_X13_EXECUTION を YES にします。" & _
        "Controlで初期値と違う値は淡い黄色になり、Controlの「初期値に戻す」ボタンで戻せます。"

    row = 16
    HomeStep ws, row, "4. 季節調整の実行", "設定_対象系列で「実行」をYESにした系列、入力した1系列、または取り込んだ全系列（原データ_系列一覧）を計算します。取込候補のまま取り込んでいない系列は対象外です。失敗した候補は理由を残し、比較に入れません。"
    modUI.AddButton ws, "btnHome_runSelected", "選択系列を実行", "modX13Run.RunSelectedFromUI", ws.Cells(row + 1, 2), 180
    modUI.AddButton ws, "btnHome_runOne", "1系列を実行", "modX13Run.RunOneFromUI", ws.Cells(row + 1, 3), 180
    modUI.AddButton ws, "btnHome_runAll", "取込済みの全系列を実行", "modX13Run.RunAllFromUI", ws.Cells(row + 2, 2), 180

    row = 20
    HomeStep ws, row, "5. 比較と最終決定", "試行_候補比較で、同じ比較群の中の順位を見て決めます。手動は行を選んでボタンを押します。自動はControlの条件で各系列の1位を保存します。"
    modUI.AddButton ws, "btnHome_finalize", "選択行を最終決定", "modX13Summary.FinalizeSelectedFromUI", ws.Cells(row + 1, 2), 180
    modUI.AddButton ws, "btnHome_autoFinalize", "条件で自動決定", "modX13Summary.AutoFinalizeFromUI", ws.Cells(row + 1, 3), 180
    modUI.AddButton ws, "btnHome_export", "系列別Excelを保存", "modX13Summary.ExportFromUI", ws.Cells(row + 2, 2), 180
    modUI.AddButton ws, "btnHome_copyOutliers", "検出外れ値を指定表へ転記", "modX13Summary.CopySelectedOutliersFromUI", ws.Cells(row + 2, 3), 180

    lines = Array( _
        "注意", _
        "・対数変換（log・auto）では、原指数の0を計算用データだけ0.000001に置き換えます。その月は対数で大きな負の値になり、モデルの選択や最終結果に影響することがあります。置換した月は計算用_ゼロ置換記録で確認できます。", _
        "・負の値・空欄・文字は0として扱わず、計算前に止めます。原データのシートと取込元のExcelは変更しません。", _
        "・事後診断（check・spectrum・history・slidingspans）は今回は実行しません。X-13の自動選択（automdl）の内部の検定と、変数のt値・p値・結合検定は出力します。", _
        "・予測誤差AAPEはX-13の標本内の値（直近3年・1年目～3年目）で、独立した将来データでの評価ではありません。", _
        "・X-13の作業フォルダ（X13_WORK_DIR）は半角英数字だけのパスが必要です。パッケージが日本語を含む場所にある場合は、C:\IIP_X13_work 等を絶対パスで指定してください。", _
        "・従来の手順（次数指定で外れ値を検出→検出した月と種類を設定_外れ値指定へ記入→別の設定IDでモデル選択）は、設定の切替と転記の補助ボタンで利用者が1段ずつ行います。")
    For i = 0 To UBound(lines)
        ws.Cells(24 + i, 2).Value2 = lines(i)
        ws.Range(ws.Cells(24 + i, 2), ws.Cells(24 + i, 4)).Merge
        ws.Cells(24 + i, 2).WrapText = True
        ws.Rows(24 + i).RowHeight = IIf(i = 0, 18, 32)
    Next i
    ws.Cells(24, 2).Font.Bold = True
End Sub

Private Sub HomeStep(ByVal ws As Worksheet, ByVal row As Long, ByVal title As String, ByVal note As String)
    ws.Cells(row, 2).Value2 = title
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 12
    ws.Cells(row, 4).Value2 = note
    ws.Cells(row, 4).WrapText = True
    ws.Rows(row + 1).RowHeight = 26
    ws.Rows(row + 2).RowHeight = 26
End Sub

' 目的  : Controlの場所の設定のうち、絶対パス（C:\ 等）が書かれている項目を返す。
'         配布ブックに作成時の絶対パスを残していないことの確認に使う。
' 出力  : 項目名を「;」で結んだ文字列。無ければ空文字。
Public Function AbsolutePathSettings() As String
    Dim key As Variant
    For Each key In Array("X13_EXECUTABLE_PATH", "X13_WORK_DIR", "X13_CANDIDATE_DIR", "X13_FINAL_DIR", "X13_SUMMARY_DIR", _
                          "INPUT_PUBLIC_EXCEL_DIR", "INPUT_SERIES_DIR", "HOLIDAY_CSV_PATH", "HOLIDAY_DOWNLOAD_DIR", "HOLIDAY_GENERATED_DIR")
        If modPackagePaths.IsAbsolutePath(modMain.GetControlValue(CStr(key), "")) Then _
            AbsolutePathSettings = AbsolutePathSettings & IIf(Len(AbsolutePathSettings) > 0, ";", "") & CStr(key)
    Next key
End Function
