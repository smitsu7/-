Attribute VB_Name = "modMain"
Option Explicit

' 役割      : Controlシートの読み書き、シートの取得・作成、ファイル名と時刻の共通処理。
'             モデル選択や数値計算は行わない。パッケージ内の場所の解決はmodPackagePathsが担当。
' 主な入口  : GetControlValue / GetControlBoolean / SetControlValue / EnsureWorksheet
' 依存先    : modSheetNames（Controlシート名）
' Controlの形式: A列=設定キー、B列=値、C列=説明。キーはA列で完全一致により探す。

' Controlの見出し行。設定はこの次の行から始まる。
Private Const CONTROL_HEADER_ROW As Long = 1

' WindowsのSYSTEMTIME構造体（UTCの取得に使う）。
Private Type UtcSystemTime
    yearValue As Integer
    monthValue As Integer
    weekdayValue As Integer
    dayValue As Integer
    hourValue As Integer
    minuteValue As Integer
    secondValue As Integer
    millisecondValue As Integer
End Type
#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef systemTime As UtcSystemTime)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef systemTime As UtcSystemTime)
#End If

' 目的  : Controlシートから設定値を文字列で読む。
' 入力  : keyName=設定キー、defaultValue=キーが無い・空欄の場合の値
' 出力  : 前後の空白を除いた文字列
' 副作用: なし
' エラー: セルがExcelのエラー値の場合は停止する（誤った値で計算しないため）。
Public Function GetControlValue(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    Dim row As Long
    Dim cellValue As Variant

    row = ControlRow(keyName)
    If row = 0 Then
        GetControlValue = defaultValue
        Exit Function
    End If
    cellValue = ThisWorkbook.Worksheets(modSheetNames.ControlSheet()).Cells(row, 2).Value2
    If IsError(cellValue) Then Err.Raise vbObjectError + 100, "GetControlValue", "Controlの値がエラーです：" & keyName
    GetControlValue = Trim$(CStr(cellValue))
    If Len(GetControlValue) = 0 Then GetControlValue = defaultValue
End Function

' 目的  : はい/いいえの設定を読む。
' 入力  : keyName、defaultValue（空欄時）
' 出力  : True/False
' エラー: YES/NO等として解釈できない値は停止する。打ち間違いを「いいえ」と扱わないため。
Public Function GetControlBoolean(ByVal keyName As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    GetControlBoolean = ParseYesNo(GetControlValue(keyName, IIf(defaultValue, "YES", "NO")), keyName)
End Function

' 目的  : はい/いいえを表す文字列をBooleanへ変換する。系列別設定でも使う。
Public Function ParseYesNo(ByVal text As String, ByVal keyName As String) As Boolean
    Select Case UCase$(Trim$(text))
        Case "YES", "TRUE", "ON", "1", "はい"
            ParseYesNo = True
        Case "NO", "FALSE", "OFF", "0", "いいえ"
            ParseYesNo = False
        Case Else
            Err.Raise vbObjectError + 101, "ParseYesNo", keyName & "はYESまたはNOで指定してください：" & text
    End Select
End Function

' 目的  : Controlの値を書き換える（無いキーは末尾に追加）。
' 副作用: Controlシートを変更する。値は文字列として保存する。
Public Sub SetControlValue(ByVal keyName As String, ByVal valueText As String, Optional ByVal note As String = "")
    Dim ws As Worksheet
    Dim row As Long

    Set ws = EnsureControlSheet()
    row = ControlRow(keyName)
    If row = 0 Then
        row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(row, 1).Value2 = keyName
        ws.Cells(row, 3).Value2 = note
    End If
    ws.Cells(row, 2).NumberFormat = "@"
    ws.Cells(row, 2).Value2 = valueText
End Sub

' 目的  : 設定キーの行番号を返す。無ければ0。
Public Function ControlRow(ByVal keyName As String) As Long
    Dim ws As Worksheet
    Dim found As Range

    Set ws = WorksheetOrNothing(modSheetNames.ControlSheet())
    If ws Is Nothing Then Exit Function
    Set found = ws.Columns(1).Find(What:=keyName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If Not found Is Nothing Then ControlRow = found.Row
End Function

' 目的  : Controlシートを返す。無ければ見出しだけの新しいシートを作る。
Public Function EnsureControlSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = WorksheetOrNothing(modSheetNames.ControlSheet())
    If ws Is Nothing Then
        Set ws = EnsureWorksheet(modSheetNames.ControlSheet())
        ws.Cells(CONTROL_HEADER_ROW, 1).Resize(1, 3).Value2 = Array("設定キー", "値", "説明")
        ws.Rows(CONTROL_HEADER_ROW).Font.Bold = True
    End If
    Set EnsureControlSheet = ws
End Function

' 目的  : 名前でシートを探す。無ければNothing。
' 補足  : シートの有無を調べる目的に限り、エラーを無視する1行だけの範囲を使う。
Public Function WorksheetOrNothing(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set WorksheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' 目的  : シートを返す。無ければブックの末尾に作る。
' 副作用: シートを追加することがある。
Public Function EnsureWorksheet(ByVal sheetName As String) As Worksheet
    Set EnsureWorksheet = WorksheetOrNothing(sheetName)
    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

' 目的  : ファイル名に使えない文字を下線に置き換える。
Public Function SafeFileStem(ByVal rawName As String) As String
    Dim forbidden As Variant
    Dim item As Variant
    Dim result As String

    result = Trim$(rawName)
    forbidden = Array("/", "\", ":", "*", "?", Chr$(34), "<", ">", "|", " ")
    For Each item In forbidden
        result = Replace(result, CStr(item), "_")
    Next item
    If Len(result) = 0 Then result = "series"
    If Len(result) > 80 Then result = Left$(result, 80)
    SafeFileStem = result
End Function

' 目的  : フォルダとファイル名を区切り文字で結ぶ。
Public Function CombinePath(ByVal folderPath As String, ByVal leafName As String) As String
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        CombinePath = folderPath & leafName
    Else
        CombinePath = folderPath & "\" & leafName
    End If
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
End Function

' 目的  : 実行ID・試行ID等に使う時刻文字列（ローカル時刻）。
Public Function TimeStamp() As String
    TimeStamp = Format$(Now, "yyyymmdd_hhnnss")
End Function

' 目的  : 協定世界時の日時（ISO 8601、末尾Z）。取得日時の記録に使う。
' 補足  : WindowsのGetSystemTimeでUTCを直接読む（r10と同じ方法）。
Public Function UtcNowText() As String
    Dim clock As UtcSystemTime
    GetSystemTime clock
    UtcNowText = Format$(clock.yearValue, "0000") & "-" & Format$(clock.monthValue, "00") & "-" & _
        Format$(clock.dayValue, "00") & "T" & Format$(clock.hourValue, "00") & ":" & _
        Format$(clock.minuteValue, "00") & ":" & Format$(clock.secondValue, "00") & "Z"
End Function
