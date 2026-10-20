Attribute VB_Name = "modSeriesSettings"
Option Explicit

' 役割      : 系列ID × 設定ID ごとに、全設定項目の値と「どこから来た値か」を一度に決める。
'             決めた結果（cfg）はSPC作成・ゼロ置換・保存の各処理へそのまま渡し、
'             「使用設定」ファイルとして試行フォルダと最終フォルダに残す。
'             あわせて、計算する系列を選ぶ表（設定_対象系列）を管理する。
' 主な入口  : ResolveAll / SettingText / SourceOf / UsedSettingsText / SetupSeriesSettingsSheet
'             RefreshSelectionSheet / SelectedSeriesIds
' 依存先    : modX13Options（定義表）、modMain、modSheetNames、modX13IO、modPublicImport（取込済みの全系列）
' 優先順位  : 1) 設定_系列別 の (series_id=その系列, setting_id=実行の設定ID) 行の非空欄
'             2) 設定_系列別 の (series_id="*", setting_id=実行の設定ID) 行の非空欄
'             3) Control の値   4) modX13Options.Definitions の初期値
'             系列別列名の無い項目（全系列共通）は 3) と 4) だけを見る。
' 表の形    : 1行目=見出し。A列 series_id、B列 setting_id、C列 説明、D列以降=系列別列名。
'             見出しに定義表に無い列名があれば停止する（打ち間違いを黙って無視しないため）。

' cfgの中で「値の出所」を入れるキーの接頭辞（画面には出ない内部の名前）。
Private Const SOURCE_PREFIX As String = "source|"

' 設定_系列別の説明欄の見出し。日本語の文字列はConstにしない（import/asciiで関数呼出しに置き換えるため）。
Private Function NoteColumnName() As String
    NoteColumnName = "説明"
End Function

' 目的  : 1系列・1設定IDの全設定を解決する。
' 入力  : seriesId、settingId
' 出力  : Dictionary。キー=設定キー、値=文字列。各キーの出所は SourceOf で読む。
' 副作用: なし（シートは読むだけ）
' エラー: 表の見出し不正、同じ系列ID×設定IDの重複行、セルのエラー値は停止。
Public Function ResolveAll(ByVal seriesId As String, ByVal settingId As String) As Object
    Dim cfg As Object, table As Variant, headerIndex As Object
    Dim seriesRow As Long, commonRow As Long
    Dim item As Variant, key As String, column As String, value As String, source As String

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbBinaryCompare
    ReadSettingsTable table, headerIndex
    seriesRow = FindRow(table, seriesId, settingId)
    commonRow = FindRow(table, "*", settingId)
    For Each item In modX13Options.Definitions()
        key = item(modX13Options.DEF_KEY)
        column = item(modX13Options.DEF_SERIES_COLUMN)
        value = ""
        source = ""
        If Len(column) > 0 And headerIndex.Exists(column) Then
            value = CellText(table, seriesRow, headerIndex(column))
            source = "設定_系列別（" & seriesId & "／" & settingId & "）"
            If Len(value) = 0 Then
                value = CellText(table, commonRow, headerIndex(column))
                source = "設定_系列別（*／" & settingId & "）"
            End If
        End If
        If Len(value) = 0 Then
            value = modMain.GetControlValue(key, "")
            source = "Control"
        End If
        If Len(value) = 0 Then
            value = item(modX13Options.DEF_DEFAULT)
            source = "初期値"
        End If
        cfg.Add key, value
        cfg.Add SOURCE_PREFIX & key, source
    Next item
    cfg.Add "SERIES_ID", seriesId
    cfg.Add "SETTING_ID", settingId
    Set ResolveAll = cfg
End Function

' 目的: 解決済み設定の値を読む。未定義キーは停止。
Public Function SettingText(ByVal cfg As Object, ByVal key As String) As String
    If Not cfg.Exists(key) Then Err.Raise vbObjectError + 2600, "SettingText", "解決済み設定にキーがありません：" & key
    SettingText = CStr(cfg(key))
End Function

' 目的: 解決済み設定の値がどこから来たか（設定_系列別／Control／初期値）。
Public Function SourceOf(ByVal cfg As Object, ByVal key As String) As String
    SourceOf = CStr(cfg(SOURCE_PREFIX & key))
End Function

' 目的: 「使用設定」ファイルの本文（タブ区切り：キー、値、出所、系列別列名）。
Public Function UsedSettingsText(ByVal cfg As Object) As String
    Dim item As Variant, body As String
    body = "key" & vbTab & "value" & vbTab & "source" & vbTab & "series_column" & vbCrLf
    body = body & "SERIES_ID" & vbTab & cfg("SERIES_ID") & vbTab & "実行対象" & vbTab & vbCrLf
    body = body & "SETTING_ID" & vbTab & cfg("SETTING_ID") & vbTab & "Control X13_SETTING_ID" & vbTab & vbCrLf
    For Each item In modX13Options.Definitions()
        body = body & item(modX13Options.DEF_KEY) & vbTab & OneLine(SettingText(cfg, CStr(item(modX13Options.DEF_KEY)))) & _
            vbTab & SourceOf(cfg, CStr(item(modX13Options.DEF_KEY))) & vbTab & item(modX13Options.DEF_SERIES_COLUMN) & vbCrLf
    Next item
    UsedSettingsText = body
End Function

Private Function OneLine(ByVal valueText As String) As String
    OneLine = Replace(Replace(Replace(valueText, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

' 目的  : 設定_系列別シートを作る（無い場合）。既にある場合は見出しを検査し、足りない列を右端へ追加する。
' 副作用: シートを作成・列を追加する。既存の値は変更しない。
Public Sub SetupSeriesSettingsSheet()
    Dim ws As Worksheet, item As Variant, column As Long, headerIndex As Object, table As Variant
    Set ws = modMain.EnsureWorksheet(modSheetNames.SeriesSettingsSheet())
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        ws.Cells(1, 1).Resize(1, 3).Value2 = Array("series_id", "setting_id", NoteColumnName())
        ws.Cells(2, 1).Resize(1, 3).Value2 = Array("*", "standard", "全系列共通の行。空欄の項目はControlの値を使う")
        ws.Columns("A:B").NumberFormat = "@"
    End If
    ReadSettingsTable table, headerIndex
    column = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For Each item In modX13Options.Definitions()
        If Len(item(modX13Options.DEF_SERIES_COLUMN)) > 0 Then
            If Not headerIndex.Exists(CStr(item(modX13Options.DEF_SERIES_COLUMN))) Then
                column = column + 1
                ws.Cells(1, column).Value2 = item(modX13Options.DEF_SERIES_COLUMN)
                headerIndex.Add CStr(item(modX13Options.DEF_SERIES_COLUMN)), column
            End If
            If Len(item(modX13Options.DEF_CHOICES)) > 0 Then _
                modX13Options.AddChoice ws.Range(ws.Cells(2, headerIndex(CStr(item(modX13Options.DEF_SERIES_COLUMN)))), _
                    ws.Cells(1000, headerIndex(CStr(item(modX13Options.DEF_SERIES_COLUMN))))), CStr(item(modX13Options.DEF_CHOICES))
        End If
    Next item
    ws.Rows(1).Font.Bold = True
    ws.Columns.ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 40
End Sub

' 目的  : 設定_系列別シートを配列で読み、見出し→列番号の辞書を作る。シートが無い場合は空の表。
' エラー: 必須見出しの欠落、未定義の列名、見出しの重複は停止。
Private Sub ReadSettingsTable(ByRef table As Variant, ByRef headerIndex As Object)
    Dim ws As Worksheet, lastRow As Long, lastColumn As Long, c As Long, name As String, known As Object, item As Variant

    Set headerIndex = CreateObject("Scripting.Dictionary")
    headerIndex.CompareMode = vbBinaryCompare
    table = Empty
    Set ws = modMain.WorksheetOrNothing(modSheetNames.SeriesSettingsSheet())
    If ws Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then Exit Sub
    Set known = CreateObject("Scripting.Dictionary")
    For Each item In modX13Options.Definitions()
        If Len(item(modX13Options.DEF_SERIES_COLUMN)) > 0 Then known(CStr(item(modX13Options.DEF_SERIES_COLUMN))) = True
    Next item
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Then lastRow = 1
    table = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Value2
    If Not IsArray(table) Then
        ' 1セルだけの範囲は配列でなく値が返るため、1×1の配列に包む。
        Dim oneCell(1 To 1, 1 To 1) As Variant
        oneCell(1, 1) = table
        table = oneCell
    End If
    For c = 1 To UBound(table, 2)
        If IsError(table(1, c)) Then Err.Raise vbObjectError + 2601, "ReadSettingsTable", "設定_系列別の見出しがエラー値です"
        name = Trim$(CStr(table(1, c)))
        If Len(name) > 0 Then
            If headerIndex.Exists(name) Then Err.Raise vbObjectError + 2602, "ReadSettingsTable", "設定_系列別の見出しが重複しています：" & name
            If c = 1 And name <> "series_id" Then Err.Raise vbObjectError + 2603, "ReadSettingsTable", "設定_系列別のA列見出しはseries_idです"
            If c = 2 And name <> "setting_id" Then Err.Raise vbObjectError + 2604, "ReadSettingsTable", "設定_系列別のB列見出しはsetting_idです"
            If c > 2 And name <> NoteColumnName() And Not known.Exists(name) Then _
                Err.Raise vbObjectError + 2605, "ReadSettingsTable", "設定_系列別に未定義の列名があります：" & name
            headerIndex.Add name, c
        End If
    Next c
End Sub

' 目的  : (series_id, setting_id) の行を探す。無ければ0。重複は停止。
Private Function FindRow(ByVal table As Variant, ByVal seriesId As String, ByVal settingId As String) As Long
    Dim r As Long
    If Not IsArray(table) Then Exit Function
    For r = 2 To UBound(table, 1)
        If Trim$(CStr(table(r, 1))) = seriesId And Trim$(CStr(table(r, 2))) = settingId Then
            If FindRow > 0 Then Err.Raise vbObjectError + 2606, "FindRow", "設定_系列別に同じ系列ID・設定IDの行が2つあります：" & seriesId & "／" & settingId
            FindRow = r
        End If
    Next r
End Function

' ================= 対象系列 =================

' 目的  : 設定_対象系列 を 原データ_系列一覧（取り込んだ系列だけ）から作り直す。既存の「実行」欄（YES/NO）は
'         系列IDで引き継ぐ。新しい系列はNOから始める。状態（IMPORTED_<月数>）も表示する。
'         取り込んでいない候補（期間不足等を含む）は載らない。それらはIIP_取込候補で確認する。
' 副作用: 設定_対象系列 シートを作り直す。
' エラー: 既存の表で系列IDが重複している場合は停止する（どちらの指定か決められないため）。
Public Sub RefreshSelectionSheet()
    Dim source As Worksheet, target As Worksheet, oldChoice As Object, r As Long, lastRow As Long
    Dim data() As Variant, srcData As Variant, count As Long, id As String
    Set target = modMain.EnsureWorksheet(modSheetNames.SeriesSelectionSheet())
    Set oldChoice = CreateObject("Scripting.Dictionary")
    lastRow = target.Cells(target.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        id = Trim$(CStr(target.Cells(r, 1).Value2))
        If Len(id) > 0 Then
            If oldChoice.Exists(id) Then Err.Raise vbObjectError + 2610, "RefreshSelectionSheet", "設定_対象系列で系列IDが重複しています：" & id
            oldChoice.Add id, UCase$(Trim$(CStr(target.Cells(r, 2).Value2)))
        End If
    Next r
    target.Cells.Clear
    target.Cells(1, 1).Resize(1, 5).Value2 = SelectionHeaders()
    target.Rows(1).Font.Bold = True
    Set source = modMain.WorksheetOrNothing(modSheetNames.SeriesListSheet())
    If Not source Is Nothing Then
        lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            srcData = source.Range(source.Cells(2, 1), source.Cells(lastRow, 8)).Value2
            ReDim data(1 To UBound(srcData, 1), 1 To 5)
            For r = 1 To UBound(srcData, 1)
                count = count + 1
                id = CStr(srcData(r, 1))
                data(count, 1) = id
                data(count, 2) = "NO"
                If oldChoice.Exists(id) Then data(count, 2) = oldChoice(id)
                data(count, 3) = srcData(r, 2)
                data(count, 4) = srcData(r, 3)
                data(count, 5) = srcData(r, 8)
            Next r
            target.Cells(2, 1).Resize(count, 5).NumberFormat = "@"
            target.Cells(2, 1).Resize(count, 5).Value2 = data
            modX13Options.AddChoice target.Cells(2, 2).Resize(count, 1), "YES,NO"
        End If
    End If
    target.Cells(1, 7).Value2 = "「実行」をYESにした系列を、「選択系列を実行」で計算します。ここには取り込んだ系列（原データ_系列一覧）だけが載ります。" & _
        "期間不足等で取り込めなかった候補はIIP_取込候補にあります。原データが計算に使えない系列は計算せず、試行_ジョブに理由を残します。"
    target.Columns("A").ColumnWidth = 28
    target.Columns("B").ColumnWidth = 8
    target.Columns("C").ColumnWidth = 30
    target.Columns("D").ColumnWidth = 10
    target.Columns("E").ColumnWidth = 22
End Sub

Public Function SelectionHeaders() As Variant
    SelectionHeaders = Array("系列ID", "実行", "系列名", "種類", "状態")
End Function

' 目的  : 今回計算する系列IDを返す。
' 入力  : Controlの X13_SERIES_SCOPE。SELECTED=設定_対象系列で実行がYES、
'         ALL_INPUT=取り込んだ全系列（原データ_系列一覧で状態がIMPORTED_。取込候補だけの系列は含まない）
' 出力  : 系列IDのCollection（表の順）
' エラー: 実行欄がYES/NO以外、系列IDの重複、該当0件は停止する。
Public Function SelectedSeriesIds() As Collection
    Dim ws As Worksheet, result As New Collection, seen As Object, r As Long, id As String, choice As String, scope As String
    scope = UCase$(modMain.GetControlValue("X13_SERIES_SCOPE", "SELECTED"))
    Set seen = CreateObject("Scripting.Dictionary")
    Select Case scope
        Case "SELECTED"
            Set ws = modMain.WorksheetOrNothing(modSheetNames.SeriesSelectionSheet())
            If ws Is Nothing Then Err.Raise vbObjectError + 2611, "SelectedSeriesIds", "設定_対象系列がありません。先に原指数を取り込んでください。"
            modX13IO.Headers ws, SelectionHeaders()
            For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                id = Trim$(CStr(ws.Cells(r, 1).Value2))
                If Len(id) > 0 Then
                    If seen.Exists(id) Then Err.Raise vbObjectError + 2612, "SelectedSeriesIds", "設定_対象系列で系列IDが重複しています：" & id
                    seen.Add id, True
                    choice = UCase$(Trim$(CStr(ws.Cells(r, 2).Value2)))
                    Select Case choice
                        Case "YES": result.Add id
                        Case "NO", ""
                        Case Else: Err.Raise vbObjectError + 2613, "SelectedSeriesIds", "設定_対象系列の「実行」はYESまたはNOです：" & id
                    End Select
                End If
            Next r
        Case "ALL_INPUT"
            ' 取り込んだ全系列。判定は modPublicImport.ImportedSeriesIds に一本化（「全系列の実行」と同じ対象）。
            Set result = modPublicImport.ImportedSeriesIds()
            If result.Count = 0 Then Err.Raise vbObjectError + 2614, "SelectedSeriesIds", "取り込んだ系列がありません。先に原指数を取り込んでください。"
        Case Else
            Err.Raise vbObjectError + 2615, "SelectedSeriesIds", "X13_SERIES_SCOPEはSELECTEDまたはALL_INPUTです：" & scope
    End Select
    If result.Count = 0 Then Err.Raise vbObjectError + 2616, "SelectedSeriesIds", "計算する系列がありません（設定_対象系列の「実行」をYESにしてください）。"
    Set SelectedSeriesIds = result
End Function

' 目的: 試験・自動化用。設定_対象系列の「実行」を、指定した系列だけYESにする（他はNO）。
Public Sub MarkSelection(ByVal ids As Collection)
    Dim ws As Worksheet, r As Long, wanted As Object, id As Variant
    Set ws = ThisWorkbook.Worksheets(modSheetNames.SeriesSelectionSheet())
    Set wanted = CreateObject("Scripting.Dictionary")
    For Each id In ids
        wanted(CStr(id)) = True
    Next id
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Cells(r, 2).Value2 = IIf(wanted.Exists(CStr(ws.Cells(r, 1).Value2)), "YES", "NO")
    Next r
End Sub

Private Function CellText(ByVal table As Variant, ByVal row As Long, ByVal column As Long) As String
    If row = 0 Then Exit Function
    If column > UBound(table, 2) Then Exit Function
    If IsError(table(row, column)) Then Err.Raise vbObjectError + 2607, "CellText", "設定_系列別のセルがエラー値です（行" & row & "）"
    CellText = Trim$(CStr(table(row, column)))
End Function
