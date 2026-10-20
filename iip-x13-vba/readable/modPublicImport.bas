Attribute VB_Name = "modPublicImport"
Option Explicit

' 役割      : 原指数を取り込む。経産省の公表Excelを直接読む経路と、系列別の縦表CSV（パッケージの
'             raw.csvと同じ列）を読む経路の2つ。取込はプレビューと確定に分け、確定時にファイルを読み直す。
' 主な入口  : SetupImportSheets / ChooseFiles / UseBundledPublicExcel / PreviewImport / CommitImport（画面）
'             SetImportFiles / SetImportOption（試験・自動化用）/ LoadRawValues（計算時の読出し）
'             ImportedSeriesIds（取込済みの全系列。「全系列の実行」の対象）
' 依存先    : modIIPTextTable、modX13IO、modHash、modMain、modSheetNames、modX13Options、
'             modPackagePaths、modSeriesSettings（対象系列表の更新）、modUI、modLog
' 系列ID    : 種類の接頭辞＋公表Excelの品目番号。同じ品目番号でも種類が違えば別の系列になる。
'               生産 industry_ / 出荷 shipment_ / 在庫 inventory_ / 在庫率 inventory_ratio_ /
'               能力 capacity_ / 稼働率 operating_ratio_
' 取込候補と取込済みの区別:
'             IIP_取込候補（プレビュー）＝元ファイルにある全系列。状態は「取り込めるか」を表す。
'               設定期間（初期値2014-01～2025-12）の全月に数値がある系列だけをREADY_<月数>とする。
'               足りない系列はINSUFFICIENT_PERIOD、期間内に文字・空欄がある系列はINVALID_VALUE。
'             原データ_系列一覧＝実際に取り込んだ系列だけ。状態はIMPORTED_<月数>、取込IDはその系列を最後に取り込んだ回。
'               取込の確定で書くのはYESにした系列の行だけ。今回選ばなかった既存の系列の行
'               （取込ID・原値・元セル・raw.csv）は変えない。「全系列」はこの一覧のIMPORTED_の系列を指す。
'             足りない月を推測・複製で埋めない。接続原指数を別ファイルで用意した場合は、
'             ファイル一覧へ追加すると、重複月の値が一致することを確かめたうえで月を合わせる。
' 書く場所  : 原データ_原指数、原データ_系列一覧、原データ_来歴、02_input\series\<系列ID>\raw.csv
'             （raw.csvが既にあれば値を照合し、違えば停止する。上書きしない）

Private Const MAX_FILES As Long = 32
Private Const FILE_FIRST_ROW As Long = 7
Private Const SETTING_FIRST_ROW As Long = 7
Private Const PREVIEW_HEADER_ROW As Long = 6
Private Const DEFAULT_PUBLIC_URL As String = "https://www.meti.go.jp/statistics/tyo/iip/index.html"
Private Const MAX_OBSERVATIONS As Long = 900000
' 原データ_系列一覧の「状態」の接頭辞。取り込めるか（READY_）と、取り込んだか（IMPORTED_）を文字で区別する。
Private Const STATUS_READY As String = "READY_"
Private Const STATUS_IMPORTED As String = "IMPORTED_"
' 原データ_系列一覧の列（SeriesListHeadersの順。1始まり）。
Private Const LIST_COL_STATUS As Long = 8
Private Const LIST_COL_IMPORT_ID As Long = 13

Private mBusy As Boolean
' 1回の読込で固定する対象期間（月番号）。
Private mFirstMonth As Long
Private mLastMonth As Long

' ================= 画面 =================

' 目的  : 取込設定シートを作る（無い場合）。既にある値は変更しない。
' 副作用: IIP_取込設定 シートとボタンを作る。
Public Sub SetupImportSheets()
    Dim ws As Worksheet, items As Variant, i As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.ImportConfigSheet())
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        ws.Cells.NumberFormat = "@"
        ws.Range("A1").Value2 = "原指数の取込設定"
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value2 = "F列に取込元ファイルを並べ、「プレビュー作成」で系列ごとの月数を確認してから取り込みます。パスはこのブックのフォルダからの相対パスで書けます。"
        ws.Range("A6:D6").Value2 = Array("設定項目", "値", "説明", "key")
        items = Array( _
            Array("表の形式", "公表Excel", "公表Excel：経産省の原指数Excel（系列が行・年月が列、または接続指数の系列が列・年月が行）。縦表CSV：series_id,date,value の列を持つCSV。", "format"), _
            Array("CSVの文字コード", "AUTO", "縦表CSVだけで使用。AUTOはUTF-8（BOM有無）とCP932を判別。", "csv_encoding"), _
            Array("読むシート", "すべて", "公表Excelで読むシート。「すべて」は生産・出荷・在庫・在庫率・能力・稼働率のうちファイルにあるもの。カンマ区切りで限定可。", "sheets"), _
            Array("出典URL", "", "空欄は経産省IIPの公表ページを記録。", "source_url"), _
            Array("資料の版", "", "空欄はファイル名を記録。", "vintage"))
        For i = 0 To UBound(items)
            ws.Cells(SETTING_FIRST_ROW + i, 1).Resize(1, 4).Value2 = items(i)
        Next i
        ws.Range("F6").Value2 = "取込元ファイル（相対パスまたは絶対パス、最大32件）"
        ws.Range("A6:D6,F6").Font.Bold = True
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("B").ColumnWidth = 14
        ws.Columns("C").ColumnWidth = 70
        ws.Columns("C").WrapText = True
        ws.Columns("D").Hidden = True
        ws.Columns("E").ColumnWidth = 3
        ws.Columns("F").ColumnWidth = 70
        modX13Options.AddChoice ws.Cells(SETTING_FIRST_ROW, 2), "公表Excel,縦表CSV"
        modX13Options.AddChoice ws.Cells(SETTING_FIRST_ROW + 1, 2), "AUTO,UTF8,CP932"
    End If
    modUI.AddButton ws, "btnImport_choose", "ファイル選択", "modPublicImport.ChooseFiles", ws.Range("A4")
    modUI.AddButton ws, "btnImport_bundled", "同梱の公表Excelを指定", "modPublicImport.UseBundledPublicExcel", ws.Range("C4")
    modUI.AddButton ws, "btnImport_preview", "プレビュー作成", "modPublicImport.PreviewFromUI", ws.Range("F4")
End Sub

' 目的: 取込元ファイルをダイアログで選び、F列へ書く（パッケージ内は相対パス）。
Public Sub ChooseFiles()
    Dim dialog As Object, ws As Worksheet, i As Long
    On Error GoTo Failed
    SetupImportSheets
    Set ws = ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet())
    Set dialog = Application.FileDialog(3)
    dialog.Title = "原指数の公表Excel（.xlsx）または縦表CSVを選択してください"
    dialog.AllowMultiSelect = True
    dialog.Filters.Clear
    dialog.Filters.Add "Excel・CSV", "*.xlsx;*.csv"
    If dialog.Show <> -1 Then Exit Sub
    If dialog.SelectedItems.Count > MAX_FILES Then Err.Raise vbObjectError + 2420, , "一度に選ぶファイルは" & MAX_FILES & "個以内としてください。"
    ws.Range(ws.Cells(FILE_FIRST_ROW, 6), ws.Cells(FILE_FIRST_ROW + MAX_FILES - 1, 6)).ClearContents
    For i = 1 To dialog.SelectedItems.Count
        ws.Cells(FILE_FIRST_ROW + i - 1, 6).Value2 = modPackagePaths.RelativeToPackage(CStr(dialog.SelectedItems(i)))
    Next i
    ws.Activate
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "原指数のファイル選択"
End Sub

' 目的: 02_input\public_excel にある.xlsxをすべて取込元に指定し、形式を公表Excelにする。
Public Sub UseBundledPublicExcel()
    Dim folder As String, fso As Object, file As Object, paths As New Collection
    On Error GoTo Failed
    SetupImportSheets
    folder = modPackagePaths.RequireFolder("INPUT_PUBLIC_EXCEL_DIR", False)
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each file In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "xlsx" And Left$(file.Name, 2) <> "~$" Then paths.Add modPackagePaths.RelativeToPackage(file.path)
    Next file
    If paths.Count = 0 Then Err.Raise vbObjectError + 2421, , "公表Excelが見つかりません：" & folder
    SetImportFiles paths
    SetImportOption "format", "公表Excel"
    ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet()).Activate
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "同梱の公表Excel"
End Sub

' 目的: 画面のボタンからプレビューを作る。
Public Sub PreviewFromUI()
    On Error GoTo Failed
    PreviewImport
    ThisWorkbook.Worksheets(modSheetNames.ImportPreviewSheet()).Activate
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description & vbLf & "取込済みの原データは変更していません。", vbExclamation, "原指数のプレビュー"
End Sub

' 目的: 画面のボタンから、プレビューでYESにした系列を取り込む。
Public Sub CommitFromUI()
    Dim count As Long
    On Error GoTo Failed
    count = CommitImport()
    ThisWorkbook.Worksheets(modSheetNames.SeriesListSheet()).Activate
    modUI.ShowMessage "今回YESにした" & count & "系列の原指数を取り込みました（取込済みは累計" & ImportedSeriesCount() & "系列）。" & vbLf & _
        "NOの系列と、期間不足等で取り込めない候補は取り込んでいません（「" & modSheetNames.ImportPreviewSheet() & "」に残っています）。" & vbLf & _
        "「" & modSheetNames.SeriesSelectionSheet() & "」で計算する系列をYESにしてください。", vbInformation, "原指数の取込"
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description & vbLf & "原データは更新前の状態を保持しています。", vbExclamation, "原指数の取込"
End Sub

' ================= 試験・自動化用の設定 =================

' 目的: 取込元ファイルの一覧を書き換える（相対パス・絶対パス）。
Public Sub SetImportFiles(ByVal paths As Collection)
    Dim ws As Worksheet, i As Long
    If paths.Count > MAX_FILES Then Err.Raise vbObjectError + 2422, "SetImportFiles", "ファイルは" & MAX_FILES & "件以内です。"
    SetupImportSheets
    Set ws = ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet())
    ws.Range(ws.Cells(FILE_FIRST_ROW, 6), ws.Cells(FILE_FIRST_ROW + MAX_FILES - 1, 6)).ClearContents
    For i = 1 To paths.Count
        ws.Cells(FILE_FIRST_ROW + i - 1, 6).Value2 = CStr(paths(i))
    Next i
End Sub

' 目的: 取込設定の値を書き換える。keyは format / csv_encoding / sheets / source_url / vintage。
Public Sub SetImportOption(ByVal key As String, ByVal value As String)
    SetupImportSheets
    ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet()).Cells(OptionRow(key), 2).Value2 = value
End Sub

Private Function OptionRow(ByVal key As String) As Long
    Dim found As Range
    Set found = ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet()).Columns(4).Find( _
        What:=key, LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=True)
    If found Is Nothing Then Err.Raise vbObjectError + 2423, "OptionRow", "取込設定の項目がありません：" & key
    OptionRow = found.Row
End Function

Private Function OptionValue(ByVal key As String) As String
    Dim cell As Range
    Set cell = ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet()).Cells(OptionRow(key), 2)
    If cell.HasFormula Or IsError(cell.Value2) Then Err.Raise vbObjectError + 2424, "OptionValue", "取込設定には値を直接記入してください：" & key
    OptionValue = Trim$(CStr(cell.Value2))
End Function

Private Function FormatCode() As String
    Select Case OptionValue("format")
        Case "公表Excel": FormatCode = "PUBLIC"
        Case "縦表CSV": FormatCode = "LONG_CSV"
        Case Else: Err.Raise vbObjectError + 2425, "FormatCode", "表の形式は「公表Excel」または「縦表CSV」です。"
    End Select
End Function

' 目的  : F列のファイルを絶対パスで返す。
' エラー: 空、重複、存在しない、拡張子違いは停止する。
Private Function SourcePaths() As Collection
    Dim ws As Worksheet, i As Long, text As String, path As String, used As Object, result As New Collection, fso As Object
    Set ws = ThisWorkbook.Worksheets(modSheetNames.ImportConfigSheet())
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 0 To MAX_FILES - 1
        If ws.Cells(FILE_FIRST_ROW + i, 6).HasFormula Or IsError(ws.Cells(FILE_FIRST_ROW + i, 6).Value2) Then _
            Err.Raise vbObjectError + 2426, "SourcePaths", "ファイルのパスは直接記入してください（F" & FILE_FIRST_ROW + i & "）。"
        text = Trim$(CStr(ws.Cells(FILE_FIRST_ROW + i, 6).Value2))
        If Len(text) > 0 Then
            path = modPackagePaths.ResolvePath(text)
            If Not fso.FileExists(path) Then Err.Raise vbObjectError + 2427, "SourcePaths", "ファイルが見つかりません：" & path
            If used.Exists(path) Then Err.Raise vbObjectError + 2428, "SourcePaths", "同じファイルが重複しています：" & path
            Select Case LCase$(fso.GetExtensionName(path))
                Case "xlsx", "csv"
                Case Else: Err.Raise vbObjectError + 2429, "SourcePaths", "対応するファイルは.xlsxと.csvです：" & path
            End Select
            used.Add path, True
            result.Add path
        End If
    Next i
    If result.Count = 0 Then Err.Raise vbObjectError + 2430, "SourcePaths", "取込元ファイルを指定してください（取込設定のF列）。"
    Set SourcePaths = result
End Function

' ================= プレビュー =================

' 目的  : 取込元を読み、系列ごとの月数と状態をプレビューシートへ書く。原データは変更しない。
' 副作用: IIP_取込候補 シートを作り直す。READYの系列は取込欄をYESにする。
' エラー: ファイル・設定の不正は停止する。
Public Sub PreviewImport()
    Dim observations As Object, series As Object, files As Collection, signature As String
    Dim oldStatus As Variant, errorNumber As Long, errorText As String
    If mBusy Then Err.Raise vbObjectError + 2431, "PreviewImport", "別の取込処理を実行中です。"
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    mBusy = True
    SetupImportSheets
    LoadSources observations, series, files, signature
    WritePreview series, signature
    mBusy = False
    Application.StatusBar = oldStatus
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    mBusy = False
    Application.StatusBar = oldStatus
    Err.Raise errorNumber, "PreviewImport", errorText
End Sub

' 目的  : 全ファイルを読み、観測値と系列情報を集める。
' 出力  : observations（キー "系列ID|月番号" → Array(値, 速報区分, 元ファイル相対, 元シート, 元セル)）
'         series（キー 系列ID → Dictionary：name, kind, code, base, months, issue, files）
'         files（Collection：Array(絶対パス, CRC32, バイト数, 読んだシート)）
'         signature（期間・設定・ファイルの内容を表す文字列。確定時の照合に使う）
Private Sub LoadSources(ByRef observations As Object, ByRef series As Object, ByRef files As Collection, ByRef signature As String)
    Dim paths As Collection, path As Variant, crc As String, mode As String, fso As Object, sheetsRead As String
    Set observations = CreateObject("Scripting.Dictionary")
    Set series = CreateObject("Scripting.Dictionary")
    Set files = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    mFirstMonth = modX13Options.FirstMonth()
    mLastMonth = modX13Options.LastMonth()
    ' 実績期間の終了が開始より前なら、読み始める前に停止する（期間の誤りで系列の状態を判定しないため）。
    modX13Options.ActualMonths
    mode = FormatCode()
    Set paths = SourcePaths()
    signature = "period=" & mFirstMonth & "-" & mLastMonth & "|format=" & mode & "|sheets=" & OptionValue("sheets") & _
        "|encoding=" & OptionValue("csv_encoding") & "|url=" & OptionValue("source_url") & "|vintage=" & OptionValue("vintage") & "|"
    For Each path In paths
        Application.StatusBar = "原指数を読み込んでいます：" & fso.GetFileName(CStr(path))
        crc = modX13IO.FileCRC(CStr(path))
        signature = signature & CStr(path) & ":" & fso.GetFile(CStr(path)).Size & ":" & crc & "|"
        If mode = "PUBLIC" Then
            If LCase$(fso.GetExtensionName(CStr(path))) <> "xlsx" Then Err.Raise vbObjectError + 2432, "LoadSources", _
                "公表Excelの形式では.xlsxを指定してください：" & CStr(path)
            sheetsRead = ReadPublicWorkbook(CStr(path), observations, series)
        Else
            If LCase$(fso.GetExtensionName(CStr(path))) <> "csv" Then Err.Raise vbObjectError + 2433, "LoadSources", _
                "縦表CSVの形式では.csvを指定してください：" & CStr(path)
            ReadLongCsv CStr(path), observations, series
            sheetsRead = "CSV"
        End If
        If crc <> modX13IO.FileCRC(CStr(path)) Then Err.Raise vbObjectError + 2434, "LoadSources", "読込中に元ファイルが変わりました：" & CStr(path)
        files.Add Array(CStr(path), crc, fso.GetFile(CStr(path)).Size, sheetsRead)
    Next path
    Application.StatusBar = False
    If series.Count = 0 Then Err.Raise vbObjectError + 2435, "LoadSources", "原指数の系列が見つかりません。ファイルと形式を確認してください。"
End Sub

' ---- 公表Excel ----

' 目的  : 公表Excelを読み取り専用で開き、認識できるシートをすべて読む。
' 出力  : 読んだシート名（読点区切り）
' 副作用: 開いている間だけ、マクロ・イベント・警告を止めて開き、終わったら元に戻す。リンクは更新しない。
' エラー: 認識できるシートが1つも無い場合は停止する。
Private Function ReadPublicWorkbook(ByVal path As String, ByVal observations As Object, ByVal series As Object) As String
    Dim book As Workbook, ws As Worksheet, openBook As Workbook, table As Variant, prefix As String, kind As String
    Dim oldSecurity As Long, oldEvents As Boolean, oldAlerts As Boolean, readList As String, wanted As String
    Dim errorNumber As Long, errorText As String
    For Each openBook In Application.Workbooks
        If StrComp(openBook.FullName, path, vbTextCompare) = 0 Then Err.Raise vbObjectError + 2436, "ReadPublicWorkbook", _
            "取込元のExcelを閉じてから実行してください：" & path
    Next openBook
    wanted = OptionValue("sheets")
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    ' 3 = msoAutomationSecurityForceDisable。開く間だけ取込元のマクロを無効にする（元の値へ戻す）。
    Application.AutomationSecurity = 3
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set book = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    For Each ws In book.Worksheets
        If KindOfSheet(ws.Name, prefix, kind) Then
            If wanted = "すべて" Or InStr("," & Replace(wanted, " ", "") & ",", "," & ws.Name & ",") > 0 Then
                table = SheetValues(ws)
                NormalizePublicSheet table, ws.Name, prefix, kind, path, observations, series
                readList = readList & IIf(Len(readList) > 0, "、", "") & ws.Name
            End If
        End If
    Next ws
    book.Close SaveChanges:=False
    Set book = Nothing
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    If Len(readList) = 0 Then Err.Raise vbObjectError + 2437, "ReadPublicWorkbook", _
        "生産・出荷・在庫・在庫率・能力・稼働率のシートが見つかりません：" & path
    ReadPublicWorkbook = readList
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    ' 後始末：開いたブックを閉じ、変更したアプリケーション設定を戻す。閉じる失敗で元のエラーを隠さない。
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise errorNumber, "ReadPublicWorkbook", errorText
End Function

' 目的  : シート名から系列の種類と接頭辞を決める。対象外のシート（利用上の注意等）はFalse。
Public Function KindOfSheet(ByVal sheetName As String, ByRef prefix As String, ByRef kind As String) As Boolean
    KindOfSheet = True
    Select Case Trim$(sheetName)
        Case "生産", "生産付加価値額": prefix = "industry_": kind = "生産"
        Case "出荷": prefix = "shipment_": kind = "出荷"
        Case "在庫": prefix = "inventory_": kind = "在庫"
        Case "在庫率": prefix = "inventory_ratio_": kind = "在庫率"
        Case "能力": prefix = "capacity_": kind = "能力"
        Case "稼働率": prefix = "operating_ratio_": kind = "稼働率"
        Case Else: KindOfSheet = False
    End Select
End Function

' 目的  : シートの使用範囲の値を2次元配列で返す。数式のセルはエラー値に置き換える
'         （再計算の結果に依存する値を原指数として扱わないため）。
Private Function SheetValues(ByVal ws As Worksheet) As Variant
    Dim area As Range, result As Variant, formulaCells As Range, cell As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, _
        ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1))
    If area.Rows.Count < 4 Or area.Columns.Count < 4 Then Err.Raise vbObjectError + 2438, "SheetValues", "表に必要な行・列がありません：" & ws.Name
    If CDbl(area.Rows.Count) * area.Columns.Count > 2000000# Then Err.Raise vbObjectError + 2439, "SheetValues", "表が大きすぎます：" & ws.Name
    result = area.Value2
    ' SpecialCellsは対象が無いとエラーになるため、その確認の1行だけエラーを無視する。
    On Error Resume Next
    Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells
            result(cell.Row, cell.Column) = CVErr(xlErrValue)
        Next cell
    End If
    SheetValues = result
End Function

' 目的  : 公表Excelの1シートを読む。2つの配置に対応する。
'         ① 系列が行・年月が列（月次の原指数）：「品目番号」の行が見出し、その右に201801等の年月
'         ② 系列が列・年月が行（接続指数）：「品目番号」の行に系列コード、次の行に名称、下に年月
' エラー: 題名が原指数でない、基準年が読めない、見出し・年月が認識できない場合は停止する。
Private Sub NormalizePublicSheet(ByRef table As Variant, ByVal sheetName As String, ByVal prefix As String, ByVal kind As String, _
                                 ByVal path As String, ByVal observations As Object, ByVal series As Object)
    Dim title As String, re As Object, base As String, headerRow As Long, idCol As Long, firstCol As Long, firstRow As Long
    Dim r As Long, c As Long, code As String, name As String, monthNumber As Long, preliminary As Boolean, text As String
    Dim fileSeen As Object
    Set fileSeen = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(table, 2)
        If Not IsError(table(1, c)) Then title = title & CStr(table(1, c))
    Next c
    If InStr(title, "原指数") = 0 Or InStr(title, "季節調整済") > 0 Then Err.Raise vbObjectError + 2440, "NormalizePublicSheet", _
        "題名が原指数を示していません：" & path & "（" & sheetName & "）"
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([12][0-9]{3})\s*[=＝]\s*100"
    If Not re.Test(title) Then Err.Raise vbObjectError + 2441, "NormalizePublicSheet", "題名から基準年を読み取れません：" & title
    base = re.Execute(title)(0).SubMatches(0)
    For r = 2 To Application.WorksheetFunction.Min(10, UBound(table, 1))
        For c = 1 To Application.WorksheetFunction.Min(5, UBound(table, 2))
            If Not IsError(table(r, c)) Then
                text = Trim$(CStr(table(r, c)))
                If text = "品目番号" Or text = "品目コード" Or text = "業種コード" Then headerRow = r: idCol = c: Exit For
            End If
        Next c
        If headerRow > 0 Then Exit For
    Next r
    If headerRow = 0 Then Err.Raise vbObjectError + 2442, "NormalizePublicSheet", "「品目番号」の見出しが見つかりません：" & path & "（" & sheetName & "）"
    For c = idCol + 2 To UBound(table, 2)
        If PublicMonth(table(headerRow, c), preliminary) >= 0 Then firstCol = c: Exit For
    Next c
    If firstCol > 0 Then
        ' ① 系列が行・年月が列
        For r = headerRow + 1 To UBound(table, 1)
            code = CellText(table, r, idCol)
            If Len(code) > 0 Then
                If Not IsDigits(code) Then Err.Raise vbObjectError + 2443, "NormalizePublicSheet", "品目番号を認識できません：" & CellAddress(r, idCol)
                name = CellText(table, r, idCol + 1)
                For c = firstCol To UBound(table, 2)
                    monthNumber = PublicMonth(table(headerRow, c), preliminary)
                    If monthNumber < 0 Then Err.Raise vbObjectError + 2444, "NormalizePublicSheet", "年月を認識できません：" & CellAddress(headerRow, c)
                    AddObservation prefix & code, name, kind, code, base, monthNumber, table(r, c), preliminary, _
                        path, sheetName, CellAddress(r, c), observations, series, fileSeen
                Next c
            End If
        Next r
    Else
        ' ② 系列が列・年月が行（接続係数の行は年月ではないため読まない）
        For r = headerRow + 2 To Application.WorksheetFunction.Min(headerRow + 10, UBound(table, 1))
            If PublicMonth(table(r, idCol), preliminary) >= 0 Then firstRow = r: Exit For
        Next r
        If firstRow = 0 Then Err.Raise vbObjectError + 2445, "NormalizePublicSheet", "年月の行を認識できません：" & path & "（" & sheetName & "）"
        For c = idCol + 1 To UBound(table, 2)
            code = CellText(table, headerRow, c)
            If Len(code) > 0 Then
                If Not IsDigits(code) Then Err.Raise vbObjectError + 2446, "NormalizePublicSheet", "系列コードを認識できません：" & CellAddress(headerRow, c)
                name = CellText(table, headerRow + 1, c)
                For r = firstRow To UBound(table, 1)
                    monthNumber = PublicMonth(table(r, idCol), preliminary)
                    If monthNumber < 0 Then Err.Raise vbObjectError + 2447, "NormalizePublicSheet", "年月を認識できません：" & CellAddress(r, idCol)
                    AddObservation prefix & code, name, kind, code, base, monthNumber, table(r, c), preliminary, _
                        path, sheetName, CellAddress(r, c), observations, series, fileSeen
                Next r
            End If
        Next c
    End If
End Sub

' 公表Excelの年月欄。「201801」「p 202608」（pは速報）を月番号にする。それ以外は-1。
Private Function PublicMonth(ByVal raw As Variant, ByRef preliminary As Boolean) As Long
    Dim text As String
    preliminary = False
    PublicMonth = -1
    If IsError(raw) Or IsEmpty(raw) Then Exit Function
    text = Trim$(CStr(raw))
    If LCase$(Left$(text, 2)) = "p " Then
        preliminary = True
        text = Trim$(Mid$(text, 3))
    End If
    PublicMonth = CompactMonth(text)
End Function

' "YYYYMM"（6桁）を月番号に。形式違い・範囲外は-1。
Private Function CompactMonth(ByVal text As String) As Long
    Dim re As Object
    CompactMonth = -1
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^([12][0-9]{3})(0[1-9]|1[0-2])$"
    If re.Test(text) Then CompactMonth = modX13IO.MonthIndex(CLng(Left$(text, 4)), CLng(Right$(text, 2)))
End Function

' ---- 縦表CSV ----

' 目的  : 縦表CSVを読む。必須列 series_id, date, value。任意列 series_name, series_kind, public_code, base_year。
'         日付は YYYY-MM または YYYYMM。値は小数点がピリオドの数値。
' エラー: 必須列の欠落、系列IDの形式違い、日付の形式違い、同じ系列・年月の重複は停止する。
'         期間内の値が空欄・文字の場合は停止せず、その系列をINVALID_VALUEとして示す。
Private Sub ReadLongCsv(ByVal path As String, ByVal observations As Object, ByVal series As Object)
    Dim table As Variant, header As Object, c As Long, r As Long, id As String, monthNumber As Long, text As String
    Dim name As String, kind As String, code As String, base As String, fileSeen As Object, required As Variant
    table = modIIPTextTable.ReadTextTable(path, OptionValue("csv_encoding"), "comma")
    Set header = CreateObject("Scripting.Dictionary")
    header.CompareMode = vbTextCompare
    For c = 1 To UBound(table, 2)
        text = Trim$(CStr(table(1, c)))
        If Len(text) > 0 Then
            If header.Exists(text) Then Err.Raise vbObjectError + 2448, "ReadLongCsv", "CSVの見出しが重複しています：" & text
            header.Add text, c
        End If
    Next c
    For Each required In Array("series_id", "date", "value")
        If Not header.Exists(CStr(required)) Then Err.Raise vbObjectError + 2449, "ReadLongCsv", "CSVに必須の列「" & required & "」がありません：" & path
    Next required
    Set fileSeen = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(table, 1)
        id = FieldText(table, r, header, "series_id")
        If Len(id) = 0 And Len(FieldText(table, r, header, "date")) = 0 And Len(FieldText(table, r, header, "value")) = 0 Then GoTo NextRow
        If Not IsIdent(id) Then Err.Raise vbObjectError + 2450, "ReadLongCsv", "系列IDは半角英数字と下線で指定してください（" & r & "行目）：" & id
        text = Replace(FieldText(table, r, header, "date"), "-", "")
        monthNumber = CompactMonth(text)
        If monthNumber < 0 Then Err.Raise vbObjectError + 2451, "ReadLongCsv", "日付はYYYY-MMで指定してください（" & r & "行目）：" & FieldText(table, r, header, "date")
        name = FieldText(table, r, header, "series_name")
        If Len(name) = 0 Then name = id
        kind = FieldText(table, r, header, "series_kind")
        If Len(kind) = 0 Then kind = KindFromId(id)
        code = FieldText(table, r, header, "public_code")
        If Len(code) = 0 Then code = CodeFromId(id)
        base = FieldText(table, r, header, "base_year")
        AddObservation id, name, kind, code, base, monthNumber, FieldText(table, r, header, "value"), False, _
            path, "CSV", r & "行目", observations, series, fileSeen
NextRow:
    Next r
End Sub

Private Function FieldText(ByRef table As Variant, ByVal r As Long, ByVal header As Object, ByVal name As String) As String
    If Not header.Exists(name) Then Exit Function
    If header(name) > UBound(table, 2) Then Exit Function
    FieldText = Trim$(CStr(table(r, header(name))))
End Function

' 系列IDの接頭辞から種類を推定する（CSVにseries_kindが無い場合）。不明は空文字。
Private Function KindFromId(ByVal id As String) As String
    Dim prefixes As Variant, kinds As Variant, i As Long
    ' 長い接頭辞から調べる（inventory_ratio_ を inventory_ より先に）。
    prefixes = Array("inventory_ratio_", "operating_ratio_", "industry_", "shipment_", "inventory_", "capacity_")
    kinds = Array("在庫率", "稼働率", "生産", "出荷", "在庫", "能力")
    For i = 0 To UBound(prefixes)
        If Left$(id, Len(prefixes(i))) = prefixes(i) Then KindFromId = kinds(i): Exit Function
    Next i
End Function

Private Function CodeFromId(ByVal id As String) As String
    Dim pos As Long
    pos = InStrRev(id, "_")
    If pos > 0 Then CodeFromId = Mid$(id, pos + 1)
End Function

' ---- 観測値の登録 ----

' 目的  : 1つの観測値を登録する。設定期間外の月は読み飛ばす（期間外の値は使わないため）。
' 規則  : 数値はそのまま保持する（丸めない）。末尾の「p」は公表Excelの速報表示として外し、区分を記録する。
'         空欄・文字は値として登録せず、その系列の検査結果に最初の1件を記録する。
'         別のファイルに同じ系列・年月があれば、値が一致する場合だけ受け入れる。
' エラー: 同じファイル内の重複、ファイル間の値の不一致、系列名・基準年の不一致は停止する。
Private Sub AddObservation(ByVal id As String, ByVal name As String, ByVal kind As String, ByVal code As String, ByVal base As String, _
                           ByVal monthNumber As Long, ByVal raw As Variant, ByVal preliminary As Boolean, _
                           ByVal path As String, ByVal sheetName As String, ByVal cellText As String, _
                           ByVal observations As Object, ByVal series As Object, ByVal fileSeen As Object)
    Dim info As Object, key As String, value As Double, inputValue As Variant, status As String, existing As Variant
    If monthNumber < mFirstMonth Or monthNumber > mLastMonth Then Exit Sub
    If Not IsIdent(id) Then Err.Raise vbObjectError + 2452, "AddObservation", "系列IDの形式が不正です：" & id
    If Len(name) = 0 Then Err.Raise vbObjectError + 2453, "AddObservation", "系列名が空欄です：" & id
    If Not series.Exists(id) Then
        Set info = CreateObject("Scripting.Dictionary")
        info.Add "name", name
        info.Add "kind", kind
        info.Add "code", code
        info.Add "base", base
        info.Add "months", CreateObject("Scripting.Dictionary")
        info.Add "issue", ""
        info.Add "files", CreateObject("Scripting.Dictionary")
        series.Add id, info
    End If
    Set info = series(id)
    If CStr(info("name")) <> name Then Err.Raise vbObjectError + 2454, "AddObservation", _
        "同じ系列IDの名称がファイル間で違います：" & id & "（" & info("name") & "／" & name & "）"
    If Len(base) > 0 And Len(CStr(info("base"))) > 0 And CStr(info("base")) <> base Then Err.Raise vbObjectError + 2455, "AddObservation", _
        "同じ系列IDの基準年がファイル間で違います：" & id
    If Len(CStr(info("base"))) = 0 Then info("base") = base
    info("files")(modPackagePaths.RelativeToPackage(path)) = True
    key = id & "|" & monthNumber
    If fileSeen.Exists(key) Then Err.Raise vbObjectError + 2456, "AddObservation", _
        "同じファイル内に系列・年月の重複があります：" & id & " " & modX13IO.MonthText(monthNumber)
    fileSeen.Add key, True
    status = IIf(preliminary, "速報", "")
    inputValue = raw
    If VarType(raw) = vbString Then
        If LCase$(Right$(Trim$(CStr(raw)), 1)) = "p" And sheetName <> "CSV" Then
            inputValue = Trim$(Left$(Trim$(CStr(raw)), Len(Trim$(CStr(raw))) - 1))
            status = "速報"
        End If
    End If
    If Not TryNumber(inputValue, value) Then
        If Len(CStr(info("issue"))) = 0 Then info("issue") = "数値でない値：" & modX13IO.MonthText(monthNumber) & "（" & sheetName & "!" & cellText & "）"
        Exit Sub
    End If
    If observations.Exists(key) Then
        existing = observations(key)
        If CDbl(existing(0)) <> value Then Err.Raise vbObjectError + 2457, "AddObservation", _
            "重複する月の原指数がファイル間で違います：" & id & " " & modX13IO.MonthText(monthNumber) & _
            "（" & existing(2) & "=" & modX13IO.NumberText(CDbl(existing(0))) & "、" & modPackagePaths.RelativeToPackage(path) & "=" & modX13IO.NumberText(value) & "）"
        Exit Sub
    End If
    If observations.Count >= MAX_OBSERVATIONS Then Err.Raise vbObjectError + 2458, "AddObservation", "観測値の数が上限を超えています。"
    observations.Add key, Array(value, status, modPackagePaths.RelativeToPackage(path), sheetName, cellText)
    info("months")(CStr(monthNumber)) = True
End Sub

' 目的  : セルの値を数値にする。Excelの数値はそのまま、文字はピリオド小数の厳密な形式だけ受け付ける。
'         空欄・文字・真偽値・エラー値はFalse（0にしない）。
Public Function TryNumber(ByVal raw As Variant, ByRef value As Double) As Boolean
    If IsError(raw) Or IsEmpty(raw) Then Exit Function
    Select Case VarType(raw)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            value = CDbl(raw)
            TryNumber = True
        Case vbString
            If modX13IO.IsNumberText(CStr(raw)) Then
                value = modX13IO.NumberValue(CStr(raw))
                TryNumber = True
            End If
    End Select
End Function

' 系列の状態と検査結果。
Private Sub SeriesStatus(ByVal info As Object, ByRef status As String, ByRef detail As String)
    Dim m As Long, missing As Long, sample As String, expected As Long
    expected = mLastMonth - mFirstMonth + 1
    For m = mFirstMonth To mLastMonth
        If Not info("months").Exists(CStr(m)) Then
            missing = missing + 1
            If missing <= 3 Then sample = sample & IIf(Len(sample) > 0, "、", "") & modX13IO.MonthText(m)
        End If
    Next m
    If Len(CStr(info("issue"))) > 0 Then
        status = "INVALID_VALUE"
        detail = CStr(info("issue"))
    ElseIf missing > 0 Then
        status = "INSUFFICIENT_PERIOD"
        detail = "期間内" & expected - missing & "か月（不足" & missing & "か月：" & sample & IIf(missing > 3, "ほか", "") & "）。不足月は推測で埋めません"
    Else
        status = "READY_" & expected
        detail = expected & "か月の連続した原指数"
    End If
End Sub

' 目的: プレビューシートを作り直す。READYの系列は取込欄をYES、それ以外はNO（選べない）。
Private Sub WritePreview(ByVal series As Object, ByVal signature As String)
    Dim ws As Worksheet, data() As Variant, id As Variant, info As Object, i As Long, status As String, detail As String
    Dim ready As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.ImportPreviewSheet())
    If ws.ProtectContents Then Err.Raise vbObjectError + 2459, "WritePreview", "取込候補のシートが保護されています。"
    ws.Cells.Clear
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Value2 = "原指数の取込候補"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value2 = "設定期間 " & modX13Options.PeriodLabel() & " の全月がそろう系列（状態READY）を取り込めます。不要な系列はNOにしてください。" & _
        "この表は取込の候補です（READYは「取り込める」の意味で、取込済みではありません）。取り込んだ系列は「" & modSheetNames.SeriesListSheet() & "」に載ります。"
    ws.Cells(PREVIEW_HEADER_ROW, 1).Resize(1, 10).Value2 = Array("取込", "系列ID", "系列名", "種類", "元コード", "基準年", _
        "期間内の有効月数", "状態", "検査結果", "元ファイル")
    ws.Rows(PREVIEW_HEADER_ROW).Font.Bold = True
    ReDim data(1 To series.Count, 1 To 10)
    For Each id In series.Keys
        i = i + 1
        Set info = series(id)
        SeriesStatus info, status, detail
        If Left$(status, 6) = "READY_" Then ready = ready + 1
        data(i, 1) = IIf(Left$(status, 6) = "READY_", "YES", "NO")
        data(i, 2) = CStr(id)
        data(i, 3) = info("name")
        data(i, 4) = info("kind")
        data(i, 5) = info("code")
        data(i, 6) = info("base")
        data(i, 7) = info("months").Count
        data(i, 8) = status
        data(i, 9) = detail
        data(i, 10) = Join(info("files").Keys, "；")
    Next id
    ws.Cells(PREVIEW_HEADER_ROW + 1, 1).Resize(series.Count, 10).Value2 = data
    ws.Cells(PREVIEW_HEADER_ROW + 1, 7).Resize(series.Count, 1).NumberFormat = "0"
    ws.Range("A3").Value2 = "候補 " & series.Count & "系列。うち全月がそろう系列 " & ready & "、不足または不正 " & series.Count - ready & "。"
    modX13Options.AddChoice ws.Cells(PREVIEW_HEADER_ROW + 1, 1).Resize(series.Count, 1), "YES,NO"
    ws.Range("K1").Value2 = "照合用の署名（編集しない）"
    ws.Range("K2").Value2 = signature
    ws.Range("K3").Value2 = modMain.UtcNowText()
    ws.Columns("K").Hidden = True
    ws.Columns("A").ColumnWidth = 7
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D:H").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 22
    ws.Columns("I").ColumnWidth = 60
    ws.Columns("J").ColumnWidth = 50
    modUI.AddButton ws, "btnImport_commit", "選択系列の取込", "modPublicImport.CommitFromUI", ws.Range("A4")
End Sub

' 目的  : 試験・自動化用。プレビューの取込欄を、指定した系列だけYESにする（READYでない系列は指定不可）。
' 入力  : ids = 系列IDのCollection。空のCollectionなら、READYの系列をすべてYESに戻す。
Public Sub SelectPreviewSeries(ByVal ids As Collection)
    Dim ws As Worksheet, r As Long, wanted As Object, id As Variant
    Set ws = ThisWorkbook.Worksheets(modSheetNames.ImportPreviewSheet())
    Set wanted = CreateObject("Scripting.Dictionary")
    For Each id In ids
        wanted(CStr(id)) = True
    Next id
    For r = PREVIEW_HEADER_ROW + 1 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If ids.Count = 0 Then
            ws.Cells(r, 1).Value2 = IIf(Left$(CStr(ws.Cells(r, 8).Value2), 6) = "READY_", "YES", "NO")
        Else
            ws.Cells(r, 1).Value2 = IIf(wanted.Exists(CStr(ws.Cells(r, 2).Value2)), "YES", "NO")
        End If
    Next r
End Sub

' ================= 確定 =================

' 目的  : プレビューでYESにした系列を取り込む。ファイルと設定を読み直し、プレビュー時と同じか確かめる。
' 出力  : 取り込んだ系列数
' 副作用: 書くのはYESにした系列（selected）についてだけ。今回NOの系列・プレビューに無い既存の系列は変えない。
'         1) 02_input\series\<系列ID>\raw.csv を新規作成（既にあれば値を照合するだけ）
'         2) 原データ_原指数 の該当系列の行を入れ替える（他の系列の行は残す）
'         3) 原データ_系列一覧 の該当系列の行を、状態IMPORTED_・今回の取込IDで書く（他の系列の行は残す）
'         4) 原データ_来歴 にファイルごとの行を追加（今回取り込んだ系列数を併記）
'         5) 設定_対象系列 に新しい系列を追加（既存のYES/NOは保つ）
' エラー: プレビュー後の変更、READYでない系列のYES、系列一覧の系列IDの重複、既存raw.csvとの値の違いは、
'         何も書く前に停止する。raw.csvを書いた後に失敗した場合は、この取込で作ったraw.csvを消してから停止する。
Public Function CommitImport() As Long
    Dim observations As Object, series As Object, files As Collection, signature As String, ws As Worksheet
    Dim selected As Object, r As Long, id As String, flag As String, status As String, detail As String
    Dim importId As String, acquired As String, rawRows As Collection, created As Collection, csvPlan As Object
    Dim errorNumber As Long, errorText As String, oldStatus As Variant, item As Variant, listRows As Object
    If mBusy Then Err.Raise vbObjectError + 2460, "CommitImport", "別の取込処理を実行中です。"
    oldStatus = Application.StatusBar
    Set created = New Collection
    On Error GoTo Failed
    mBusy = True
    Set ws = modMain.WorksheetOrNothing(modSheetNames.ImportPreviewSheet())
    If ws Is Nothing Then Err.Raise vbObjectError + 2461, "CommitImport", "先にプレビューを作成してください。"
    LoadSources observations, series, files, signature
    If signature <> CStr(ws.Range("K2").Value2) Then Err.Raise vbObjectError + 2462, "CommitImport", _
        "元ファイル・取込設定・期間がプレビュー後に変わっています。プレビューを作り直してください。"
    Set selected = CreateObject("Scripting.Dictionary")
    For r = PREVIEW_HEADER_ROW + 1 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        id = CStr(ws.Cells(r, 2).Value2)
        flag = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
        If Not series.Exists(id) Then Err.Raise vbObjectError + 2463, "CommitImport", "取込候補の系列IDが元ファイルにありません。プレビューを作り直してください：" & id
        If flag <> "YES" And flag <> "NO" Then Err.Raise vbObjectError + 2464, "CommitImport", "取込欄はYESまたはNOです（" & id & "）。"
        If flag = "YES" Then
            SeriesStatus series(id), status, detail
            If Left$(status, 6) <> "READY_" Then Err.Raise vbObjectError + 2465, "CommitImport", _
                "期間内の月がそろわない系列は取り込めません：" & id & "（" & detail & "）"
            selected.Add id, True
        End If
    Next r
    If selected.Count = 0 Then Err.Raise vbObjectError + 2466, "CommitImport", "取り込む系列をYESにしてください。"
    Set listRows = SeriesListRows()                                               ' 系列一覧の重複検査（書込み前）
    importId = NewImportId()
    acquired = modMain.UtcNowText()
    Set rawRows = BuildRawRows(observations, selected)
    Set csvPlan = PlanRawCsv(observations, series, selected, files, acquired)     ' 既存raw.csvとの照合（書込み前）
    For Each item In csvPlan.Keys
        If Len(CStr(csvPlan(item)(1))) > 0 Then
            modPackagePaths.EnsureFolderTree CreateObject("Scripting.FileSystemObject").GetParentFolderName(CStr(csvPlan(item)(0)))
            modX13IO.WriteNewText CStr(csvPlan(item)(0)), CStr(csvPlan(item)(1))
            created.Add CStr(csvPlan(item)(0))
        End If
    Next item
    ReplaceRawIndexRows rawRows, selected
    UpdateSeriesList series, selected, listRows, csvPlan, importId, acquired
    AppendProvenance files, importId, acquired, selected.Count
    modSeriesSettings.RefreshSelectionSheet
    modLog.WriteLog "CommitImport", "INFO", importId & " 系列=" & selected.Count & " 新規raw.csv=" & created.Count
    mBusy = False
    Application.StatusBar = oldStatus
    CommitImport = selected.Count
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    ' 後始末：この取込で作ったraw.csvだけを消す。消せない場合もログに残して元のエラーを返す。
    On Error Resume Next
    For Each item In created
        Kill CStr(item)
        If Err.Number <> 0 Then modLog.WriteLog "CommitImport", "ERROR", "作成途中のraw.csvを消せません：" & CStr(item): Err.Clear
    Next item
    mBusy = False
    Application.StatusBar = oldStatus
    On Error GoTo 0
    modLog.WriteLog "CommitImport", "ERROR", errorText
    Err.Raise errorNumber, "CommitImport", errorText
End Function

' 原データ_原指数へ書く行（選択系列の期間内の全月）。要素は Array(系列ID, 年月, 原指数, 速報区分, 元ファイル, 元シート, 元セル)。
Private Function BuildRawRows(ByVal observations As Object, ByVal selected As Object) As Collection
    Dim rows As New Collection, id As Variant, m As Long, obs As Variant
    For Each id In selected.Keys
        For m = mFirstMonth To mLastMonth
            obs = observations(CStr(id) & "|" & m)
            rows.Add Array(CStr(id), modX13IO.MonthText(m), CDbl(obs(0)), obs(1), obs(2), obs(3), obs(4))
        Next m
    Next id
    Set BuildRawRows = rows
End Function

' 目的  : 系列別raw.csvの計画を作る。既にあるファイルは値を照合し、無いファイルは本文を用意する。
' 出力  : Dictionary（系列ID → Array(絶対パス, 新規本文（既存なら空文字）, 状態の説明)）
' エラー: 既存ファイルと値・月数が違えば停止する（上書きしない）。
Private Function PlanRawCsv(ByVal observations As Object, ByVal series As Object, ByVal selected As Object, _
                            ByVal files As Collection, ByVal acquired As String) As Object
    Dim plan As Object, folder As String, id As Variant, path As String, body As String, m As Long, obs As Variant
    Dim info As Object, fileNames As String, item As Variant
    Set plan = CreateObject("Scripting.Dictionary")
    folder = modPackagePaths.RequireFolder("INPUT_SERIES_DIR", True)
    For Each item In files
        fileNames = fileNames & IIf(Len(fileNames) > 0, " + ", "") & CreateObject("Scripting.FileSystemObject").GetFileName(CStr(item(0)))
    Next item
    For Each id In selected.Keys
        path = modMain.CombinePath(modMain.CombinePath(folder, CStr(id)), "raw.csv")
        Set info = series(id)
        If modMain.FileExists(path) Then
            CompareExistingRawCsv path, CStr(id), observations
            plan.Add CStr(id), Array(path, "", "既存のraw.csvと一致")
        Else
            body = "series_id,date,value,series_name,base_year,adjustment,unit,source_url,acquired_at,vintage,source_excel,series_kind" & vbCrLf
            For m = mFirstMonth To mLastMonth
                obs = observations(CStr(id) & "|" & m)
                body = body & CsvField(CStr(id)) & "," & modX13IO.MonthText(m) & "," & modX13IO.NumberText(CDbl(obs(0))) & "," & _
                    CsvField(CStr(info("name"))) & "," & CsvField(CStr(info("base"))) & ",RAW," & _
                    CsvField("指数（" & info("base") & "年＝100）") & "," & CsvField(SourceUrl()) & "," & acquired & "," & _
                    CsvField(VintageText(fileNames)) & "," & CsvField(CStr(obs(2))) & "," & CsvField(CStr(info("kind"))) & vbCrLf
            Next m
            plan.Add CStr(id), Array(path, body, "この取込で作成")
        End If
    Next id
    Set PlanRawCsv = plan
End Function

' 既存raw.csvの series_id・date・value が今回の値と同じか確かめる。違えば停止する。
' 規則: 期間内の各年月がちょうど1行ずつあること（重複も欠落も拒否）。行数だけで判断しない
'       （同じ月の重複と別の月の欠落が相殺して行数が合う場合を見逃さないため）。期間外の行は照合しない。
Private Sub CompareExistingRawCsv(ByVal path As String, ByVal id As String, ByVal observations As Object)
    Dim table As Variant, header As Object, c As Long, r As Long, count As Long, m As Long, value As Double, obs As Variant
    Dim seen As Object, missing As Long, sample As String
    table = modIIPTextTable.ReadTextTable(path, "AUTO", "comma")
    Set header = CreateObject("Scripting.Dictionary")
    header.CompareMode = vbTextCompare
    For c = 1 To UBound(table, 2)
        header(Trim$(CStr(table(1, c)))) = c
    Next c
    If Not header.Exists("series_id") Or Not header.Exists("date") Or Not header.Exists("value") Then _
        Err.Raise vbObjectError + 2467, "CompareExistingRawCsv", "既存のraw.csvに series_id・date・value の列がありません：" & path
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(table, 1)
        If Len(FieldText(table, r, header, "series_id")) > 0 Then
            If FieldText(table, r, header, "series_id") <> id Then Err.Raise vbObjectError + 2468, "CompareExistingRawCsv", _
                "既存のraw.csvに別の系列IDがあります（" & r & "行目）：" & path
            m = CompactMonth(Replace(FieldText(table, r, header, "date"), "-", ""))
            If m < 0 Then Err.Raise vbObjectError + 2475, "CompareExistingRawCsv", _
                "既存のraw.csvの日付を読めません（" & r & "行目）：" & FieldText(table, r, header, "date") & "　" & path
            If m >= mFirstMonth And m <= mLastMonth Then
                If seen.Exists(m) Then Err.Raise vbObjectError + 2476, "CompareExistingRawCsv", _
                    "既存のraw.csvに同じ年月の行が2つあります（" & modX13IO.MonthText(m) & "、" & seen(m) & "行目と" & r & "行目）：" & path
                seen.Add m, r
                count = count + 1
                If Not TryNumber(FieldText(table, r, header, "value"), value) Then Err.Raise vbObjectError + 2469, "CompareExistingRawCsv", _
                    "既存のraw.csvの値が数値ではありません（" & r & "行目）：" & path
                obs = observations(id & "|" & m)
                If value <> CDbl(obs(0)) Then Err.Raise vbObjectError + 2470, "CompareExistingRawCsv", _
                    "既存のraw.csvの値が今回の取込と違います（" & modX13IO.MonthText(m) & "：既存 " & modX13IO.NumberText(value) & _
                    "、今回 " & modX13IO.NumberText(CDbl(obs(0))) & "）。上書きしないため停止しました。古いファイルを別の場所へ移してから取り込んでください：" & path
            End If
        End If
    Next r
    For m = mFirstMonth To mLastMonth
        If Not seen.Exists(m) Then
            missing = missing + 1
            If missing <= 3 Then sample = sample & IIf(Len(sample) > 0, "、", "") & modX13IO.MonthText(m)
        End If
    Next m
    If missing > 0 Then Err.Raise vbObjectError + 2471, "CompareExistingRawCsv", _
        "既存のraw.csvに期間内の" & missing & "か月がありません（" & sample & IIf(missing > 3, "ほか", "") & "）：" & path
    If count <> mLastMonth - mFirstMonth + 1 Then Err.Raise vbObjectError + 2477, "CompareExistingRawCsv", _
        "既存のraw.csvの期間内の月数が" & count & "です（期待" & mLastMonth - mFirstMonth + 1 & "）：" & path
End Sub

Private Function CsvField(ByVal text As String) As String
    If InStr(text, ",") > 0 Or InStr(text, """") > 0 Or InStr(text, vbLf) > 0 Then
        CsvField = """" & Replace(text, """", """""") & """"
    Else
        CsvField = text
    End If
End Function

Private Function SourceUrl() As String
    SourceUrl = OptionValue("source_url")
    If Len(SourceUrl) = 0 Then SourceUrl = DEFAULT_PUBLIC_URL
End Function

Private Function VintageText(ByVal fileNames As String) As String
    VintageText = OptionValue("vintage")
    If Len(VintageText) = 0 Then VintageText = fileNames
End Function

' 目的: 原データ_原指数 の選択系列の行を入れ替える。他の系列の行は順序を保って残す。
Private Sub ReplaceRawIndexRows(ByVal rawRows As Collection, ByVal selected As Object)
    Dim ws As Worksheet, headers As Variant, old As Variant, lastRow As Long, keep As New Collection, r As Long, c As Long
    Dim data() As Variant, i As Long, item As Variant, total As Long
    headers = RawIndexHeaders()
    Set ws = modMain.EnsureWorksheet(modSheetNames.RawIndexSheet())
    modX13IO.Headers ws, headers
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        old = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, UBound(headers) + 1)).Value2
        For r = 1 To UBound(old, 1)
            If Not selected.Exists(CStr(old(r, 1))) Then
                keep.Add r
            End If
        Next r
    End If
    total = keep.Count + rawRows.Count
    If total + 1 > ws.Rows.Count Then Err.Raise vbObjectError + 2472, "ReplaceRawIndexRows", "原データの行数がExcelの上限を超えます。"
    ReDim data(1 To total, 1 To UBound(headers) + 1)
    For Each item In keep
        i = i + 1
        For c = 1 To UBound(headers) + 1
            data(i, c) = old(CLng(item), c)
        Next c
    Next item
    For Each item In rawRows
        i = i + 1
        For c = 1 To UBound(headers) + 1
            data(i, c) = item(c - 1)
        Next c
    Next item
    If lastRow >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, UBound(headers) + 1)).ClearContents
    ws.Cells(2, 1).Resize(total, UBound(headers) + 1).NumberFormat = "@"
    ws.Cells(2, 3).Resize(total, 1).NumberFormat = "General"
    ws.Cells(2, 1).Resize(total, UBound(headers) + 1).Value2 = data
End Sub

Public Function RawIndexHeaders() As Variant
    RawIndexHeaders = Array("系列ID", "年月", "原指数", "速報区分", "元ファイル", "元シート", "元セル")
End Function

Public Function SeriesListHeaders() As Variant
    SeriesListHeaders = Array("系列ID", "系列名", "種類", "元コード", "基準年", "期間", "期間内の有効月数", "状態", "検査結果", _
        "原値CSV", "原値CSVの扱い", "元ファイル", "取込ID", "取込日時UTC")
End Function

' 目的  : 原データ_系列一覧の既存の行を、系列ID → 行番号で返す。シートが無ければ空。
' エラー: 見出しが違う、同じ系列IDの行が2つある場合は停止する（どちらを更新すべきか決められないため）。
'         取込の確定では、何かを書く前にこれを呼ぶ。
Private Function SeriesListRows() As Object
    Dim ws As Worksheet, rowOf As Object, r As Long, id As String
    Set rowOf = CreateObject("Scripting.Dictionary")
    Set ws = modMain.WorksheetOrNothing(modSheetNames.SeriesListSheet())
    If Not ws Is Nothing Then
        modX13IO.Headers ws, SeriesListHeaders()
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            id = Trim$(CStr(ws.Cells(r, 1).Value2))
            If Len(id) > 0 Then
                If rowOf.Exists(id) Then Err.Raise vbObjectError + 2478, "SeriesListRows", _
                    "原データ_系列一覧で系列IDが重複しています（" & rowOf(id) & "行目と" & r & "行目）：" & id
                rowOf.Add id, r
            End If
        Next r
    End If
    Set SeriesListRows = rowOf
End Function

' 目的  : 原データ_系列一覧 に、今回取り込んだ系列（selected）の行だけを書く。
'         既にある系列は同じ行を上書きし、新しい系列は末尾へ加える。selectedに無い系列の行には触れない
'         （前回までに取り込んだ系列の取込ID・取込日時・raw.csvの扱いは、その系列を再び取り込むまで変わらない）。
' 入力  : series（プレビューと同じ読込結果）、selected（今回YESの系列ID）、listRows（SeriesListRowsの結果）、
'         csvPlan（PlanRawCsvの結果）、importId、acquired
' 規則  : 状態は IMPORTED_<月数>。検査結果の欄に「取込済み」と取込時の検査内容を書く。
Private Sub UpdateSeriesList(ByVal series As Object, ByVal selected As Object, ByVal listRows As Object, _
                             ByVal csvPlan As Object, ByVal importId As String, ByVal acquired As String)
    Dim ws As Worksheet, headers As Variant, lastRow As Long, id As Variant, info As Object
    Dim status As String, detail As String, target As Long, values As Variant
    headers = SeriesListHeaders()
    Set ws = modMain.EnsureWorksheet(modSheetNames.SeriesListSheet())
    modX13IO.Headers ws, headers
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For Each id In selected.Keys
        Set info = series(id)
        SeriesStatus info, status, detail
        ' CommitImportで確かめ済みだが、READYでない系列を取込済みと書かないよう、ここでも確かめる。
        If Left$(status, Len(STATUS_READY)) <> STATUS_READY Then Err.Raise vbObjectError + 2479, "UpdateSeriesList", _
            "期間内の月がそろわない系列を取込済みにできません：" & CStr(id)
        If listRows.Exists(CStr(id)) Then
            target = listRows(CStr(id))
        Else
            lastRow = lastRow + 1
            target = lastRow
        End If
        values = Array(CStr(id), info("name"), info("kind"), info("code"), info("base"), PeriodText(), info("months").Count, _
            STATUS_IMPORTED & Mid$(status, Len(STATUS_READY) + 1), "取込済み：" & detail, _
            modPackagePaths.RelativeToPackage(CStr(csvPlan(CStr(id))(0))), csvPlan(CStr(id))(2), _
            Join(info("files").Keys, "；"), importId, acquired)
        ws.Cells(target, 1).Resize(1, UBound(headers) + 1).NumberFormat = "@"
        ws.Cells(target, 7).NumberFormat = "0"
        ws.Cells(target, 1).Resize(1, UBound(headers) + 1).Value2 = values
    Next id
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("I").ColumnWidth = 50
End Sub

' 目的: 取込ID（I_年月日_時分秒）。同じ秒の取込IDが系列一覧・来歴にあれば末尾に番号を付ける
'       （続けて取り込んだ回を区別し、前回の取込IDと混ざらないようにするため）。
Private Function NewImportId() As String
    Dim base As String, n As Long
    base = "I_" & modMain.TimeStamp()
    NewImportId = base
    Do While ImportIdUsed(NewImportId)
        n = n + 1
        NewImportId = base & "_" & n
    Loop
End Function

Private Function ImportIdUsed(ByVal importId As String) As Boolean
    Dim ws As Worksheet
    Set ws = modMain.WorksheetOrNothing(modSheetNames.SeriesListSheet())
    If Not ws Is Nothing Then
        If Application.WorksheetFunction.CountIf(ws.Columns(LIST_COL_IMPORT_ID), importId) > 0 Then ImportIdUsed = True
    End If
    Set ws = modMain.WorksheetOrNothing(modSheetNames.ProvenanceSheet())
    If Not ws Is Nothing Then
        If Application.WorksheetFunction.CountIf(ws.Columns(1), importId) > 0 Then ImportIdUsed = True
    End If
End Function

' 目的  : 取込済みの全系列の系列IDを、原データ_系列一覧の行の順に返す。
'         「全系列の実行」（modX13Run.AllReadySeriesIds、X13_SERIES_SCOPE=ALL_INPUT）はこれを使う。
' 規則  : 状態がIMPORTED_で始まり、取込IDがある行だけ。プレビューだけで取り込んでいない候補や、
'         期間不足の候補は含まない（それらは系列一覧に載らない）。
' 出力  : 系列IDのCollection。シートが無い・取込済みが0件なら空のCollection（呼出側が理由を示して停止する）。
' エラー: 取込済みの系列IDが重複していれば停止する。
Public Function ImportedSeriesIds() As Collection
    Dim ws As Worksheet, result As New Collection, seen As Object, r As Long, id As String
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = modMain.WorksheetOrNothing(modSheetNames.SeriesListSheet())
    If Not ws Is Nothing Then
        modX13IO.Headers ws, SeriesListHeaders()
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            id = Trim$(CStr(ws.Cells(r, 1).Value2))
            If Len(id) > 0 And Left$(CStr(ws.Cells(r, LIST_COL_STATUS).Value2), Len(STATUS_IMPORTED)) = STATUS_IMPORTED And _
               Len(Trim$(CStr(ws.Cells(r, LIST_COL_IMPORT_ID).Value2))) > 0 Then
                If seen.Exists(id) Then Err.Raise vbObjectError + 2480, "ImportedSeriesIds", "原データ_系列一覧で系列IDが重複しています：" & id
                seen.Add id, True
                result.Add id
            End If
        Next r
    End If
    Set ImportedSeriesIds = result
End Function

' 目的: 取込済みの系列数（画面の表示用）。
Public Function ImportedSeriesCount() As Long
    ImportedSeriesCount = ImportedSeriesIds().Count
End Function

Private Function PeriodText() As String
    PeriodText = modX13IO.MonthText(mFirstMonth) & "～" & modX13IO.MonthText(mLastMonth)
End Function

' 目的: 原データ_来歴 に、取込元のファイルごとの行を追加する（SHA-256を含む）。前回までの行は残す。
'       ファイルの全系列を取り込んだとは限らないため、今回取り込んだ系列数を併記する
'       （どの系列をどの取込IDで取り込んだかは原データ_系列一覧の取込ID、月ごとの元セルは原データ_原指数）。
Private Sub AppendProvenance(ByVal files As Collection, ByVal importId As String, ByVal acquired As String, ByVal seriesCount As Long)
    Dim ws As Worksheet, rows As New Collection, item As Variant, startRow As Long, count As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.ProvenanceSheet())
    modX13IO.Headers ws, ProvenanceHeaders()
    For Each item In files
        Application.StatusBar = "SHA-256を計算しています：" & CStr(item(0))
        rows.Add Array(importId, acquired, modPackagePaths.RelativeToPackage(CStr(item(0))), modHash.FileSHA256(CStr(item(0))), _
            item(1), CStr(item(2)), item(3), OptionValue("format"), SourceUrl(), _
            VintageText(CreateObject("Scripting.FileSystemObject").GetFileName(CStr(item(0)))), PeriodText(), CStr(seriesCount))
    Next item
    modX13IO.AppendRows ws, modX13IO.RowsToArray(rows, UBound(ProvenanceHeaders()) + 1), startRow, count
End Sub

Public Function ProvenanceHeaders() As Variant
    ProvenanceHeaders = Array("取込ID", "取込日時UTC", "元ファイル", "SHA-256", "CRC32", "バイト数", "読んだシート", "形式", "出典URL", _
        "資料の版", "対象期間", "今回取り込んだ系列数")
End Function

' ================= 計算時の読出し =================

' 目的  : 原データ_原指数から、指定した系列の期間内の値を読み出す（X-13へ渡す前の原値）。
' 入力  : ids（系列IDのCollection）、firstMonth、months（実績月数）
' 出力  : Dictionary（系列ID → Dictionary："values"（Variant配列1始まり、セルの値そのまま）、
'         "sources"（Collection：Array(年月, 原指数, 速報区分, 元ファイル, 元シート, 元セル, 原データ_原指数の行番号)）、
'         "missing"（不足の説明）)）
'         values と sources は、どちらも実績期間の1か月目から月の順に並べる（i番目＝firstMonth+i-1の月）。
'         原データ_原指数の行の並び順には依存しない（行を並べ替えても、年月・原値・元セルの対応は変わらない）。
'         不足・重複がある系列は "missing" に理由を入れる（呼出側が実行前エラーとして記録し、計算しない）。
Public Function LoadRawValues(ByVal ids As Collection, ByVal firstMonth As Long, ByVal months As Long) As Object
    Dim ws As Worksheet, data As Variant, lastRow As Long, r As Long, result As Object, id As Variant, entry As Object
    Dim m As Long, k As Long, values() As Variant, missingCount As Long, sample As String, ordered As Collection
    Set result = CreateObject("Scripting.Dictionary")
    For Each id In ids
        Set entry = CreateObject("Scripting.Dictionary")
        ' raw と origin は「月の位置k（0始まり）」をキーにする。行の順ではなく年月で値と元セルを結ぶため。
        entry.Add "raw", CreateObject("Scripting.Dictionary")
        entry.Add "origin", CreateObject("Scripting.Dictionary")
        entry.Add "missing", ""
        If Not result.Exists(CStr(id)) Then result.Add CStr(id), entry
    Next id
    Set ws = modMain.WorksheetOrNothing(modSheetNames.RawIndexSheet())
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 7)).Value2
            For r = 1 To UBound(data, 1)
                If result.Exists(CStr(data(r, 1))) Then
                    If Not CStr(data(r, 2)) Like "####-##" Then Err.Raise vbObjectError + 2474, "LoadRawValues", _
                        "原データ_原指数の年月はYYYY-MMです（" & r + 1 & "行目）：" & CStr(data(r, 2))
                    m = modX13IO.MonthIndexFromText(CStr(data(r, 2)))
                    k = m - firstMonth
                    If k >= 0 And k < months Then
                        Set entry = result(CStr(data(r, 1)))
                        If entry("raw").Exists(k) Then
                            entry("missing") = "原データ_原指数に同じ年月の行が2つあります：" & CStr(data(r, 2))
                        Else
                            entry("raw").Add k, data(r, 3)
                            entry("origin").Add k, Array(CStr(data(r, 2)), data(r, 3), data(r, 4), data(r, 5), data(r, 6), data(r, 7), r + 1)
                        End If
                    End If
                End If
            Next r
        End If
    End If
    For Each id In ids
        Set entry = result(CStr(id))
        ReDim values(1 To months)
        Set ordered = New Collection
        missingCount = 0
        sample = ""
        For k = 0 To months - 1
            If entry("raw").Exists(k) Then
                values(k + 1) = entry("raw")(k)
                ordered.Add entry("origin")(k)
            Else
                missingCount = missingCount + 1
                If missingCount <= 3 Then sample = sample & IIf(Len(sample) > 0, "、", "") & modX13IO.MonthText(firstMonth + k)
            End If
        Next k
        If missingCount > 0 And Len(entry("missing")) = 0 Then entry("missing") = "原データ_原指数に期間内の" & missingCount & "か月がありません（" & sample & IIf(missingCount > 3, "ほか", "") & _
            "）。期間不足の系列は計算しません。"
        entry.Add "values", values
        entry.Add "sources", ordered
    Next id
    Set LoadRawValues = result
End Function

' ================= 小さな補助 =================

Private Function CellText(ByRef table As Variant, ByVal r As Long, ByVal c As Long) As String
    If r < 1 Or c < 1 Or r > UBound(table, 1) Or c > UBound(table, 2) Then Exit Function
    If IsError(table(r, c)) Then Err.Raise vbObjectError + 2473, "CellText", "数式またはエラー値のセルがあります：" & CellAddress(r, c)
    CellText = Trim$(CStr(table(r, c)))
End Function

Private Function CellAddress(ByVal r As Long, ByVal c As Long) As String
    Dim letters As String, index As Long
    index = c
    Do While index > 0
        index = index - 1
        letters = Chr$(65 + index Mod 26) & letters
        index = index \ 26
    Loop
    CellAddress = letters & r
End Function

Private Function IsDigits(ByVal text As String) As Boolean
    Dim i As Long, code As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next i
    IsDigits = True
End Function

' 系列IDに使える文字（半角英数字と下線、128文字以内）か。
Public Function IsIdent(ByVal text As String) As Boolean
    Dim i As Long, code As Long
    If Len(text) = 0 Or Len(text) > 128 Then Exit Function
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 95) Then Exit Function
    Next i
    IsIdent = True
End Function
