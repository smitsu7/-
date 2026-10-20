Attribute VB_Name = "modLog"
Option Explicit

' 役割      : 操作とエラーをVBA_Logシートへ1行ずつ記録する。
' 主な入口  : WriteLog
' 依存先    : modMain、modSheetNames
' 列        : 日時、処理名、重要度（INFO/WARN/ERROR/SKIP）、内容、Excel利用者名、ブック名

' 目的  : ログを1行追記する。
' 副作用: VBA_Logシートを作成・追記する。
' エラー: ログ自体の書込みに失敗した場合は、イミディエイトウィンドウへ出力して処理を続ける。
'         ログの失敗で本来の計算エラーの報告を上書きしないため。
Public Sub WriteLog(ByVal procedureName As String, ByVal levelName As String, ByVal messageText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo Fallback
    Set ws = GetLogSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).NumberFormat = "@"
    ws.Cells(nextRow, 2).Resize(1, 5).Value2 = Array(OneLine(procedureName), UCase$(OneLine(levelName)), _
        OneLine(messageText), Application.UserName, ThisWorkbook.Name)
    Exit Sub
Fallback:
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), procedureName, levelName, messageText
End Sub

Private Function GetLogSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = modMain.WorksheetOrNothing(modSheetNames.LogSheet())
    If ws Is Nothing Then
        Set ws = modMain.EnsureWorksheet(modSheetNames.LogSheet())
        ws.Cells(1, 1).Resize(1, 6).Value2 = Array("日時", "処理名", "重要度", "内容", "Excel利用者", "ブック")
        ws.Rows(1).Font.Bold = True
        ws.Columns("D").ColumnWidth = 90
    End If
    Set GetLogSheet = ws
End Function

Private Function OneLine(ByVal text As String) As String
    OneLine = Replace(Replace(text, vbCr, " "), vbLf, " ")
End Function
