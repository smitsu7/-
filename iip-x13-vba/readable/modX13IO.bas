Attribute VB_Name = "modX13IO"
Option Explicit

' 役割      : 年月の表記変換、数値の読取り・固定小数の書出し、テキスト/バイトの入出力、
'             CRC32、X-13の保存表（D10等）の読取り、記録ファイル（キー=値）の読書き、
'             シートへの一括書込みを担当する。業務判断は行わない。
' 主な入口  : MonthIndex / MonthText / FixedDecimalText / ReadText / WriteNewText / FileCRC / AppendRows
'             ReadMonthlyTable / ReadMetaFile / WriteMetaFile / RequireAsciiFile
' 依存先    : modMain
' 年月の表し方（月番号）:
'   このツールでは年月を「月番号 = 年 × 12 + 月 − 1」の整数で扱う。
'   例：2014年1月 = 2014*12+0 = 24168、2025年12月 = 24311。差+1が月数（144）。
'   r10では二通りの番号（+0と−1）が混在していたため、r11では−1の形に統一した。

' DATと計算表に書く数値の正規表現（小数点はピリオド）。
Private Const NUMBER_PATTERN As String = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([Ee][+-]?[0-9]+)?$"

' ---- 年月 ----

' 目的: 年と月から月番号を作る。範囲外は停止。
Public Function MonthIndex(ByVal yearValue As Long, ByVal monthValue As Long) As Long
    If yearValue < 1900 Or yearValue > 2999 Or monthValue < 1 Or monthValue > 12 Then _
        Err.Raise vbObjectError + 1501, "MonthIndex", "年月が範囲外です：" & yearValue & "-" & monthValue
    MonthIndex = yearValue * 12 + monthValue - 1
End Function

Public Function YearOf(ByVal monthNumber As Long) As Long
    YearOf = monthNumber \ 12
End Function

Public Function MonthOf(ByVal monthNumber As Long) As Long
    MonthOf = monthNumber Mod 12 + 1
End Function

' 目的: "YYYY-MM" を月番号へ。形式違いは停止。
Public Function MonthIndexFromText(ByVal text As String) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]{4}-[0-9]{2}$"
    If Not re.Test(text) Then Err.Raise vbObjectError + 1502, "MonthIndexFromText", "年月はYYYY-MMで指定してください：" & text
    MonthIndexFromText = MonthIndex(CLng(Left$(text, 4)), CLng(Right$(text, 2)))
End Function

' 目的: 月番号を "YYYY-MM" へ。
Public Function MonthText(ByVal monthNumber As Long) As String
    MonthText = Format$(YearOf(monthNumber), "0000") & "-" & Format$(MonthOf(monthNumber), "00")
End Function

' 目的: 月番号を X-13の保存表と同じ "YYYYMM" へ。
Public Function MonthCompact(ByVal monthNumber As Long) As String
    MonthCompact = Format$(YearOf(monthNumber), "0000") & Format$(MonthOf(monthNumber), "00")
End Function

' 目的: SPCの開始時点 "YYYY.M" へ。
Public Function X13StartText(ByVal monthNumber As Long) As String
    X13StartText = CStr(YearOf(monthNumber)) & "." & CStr(MonthOf(monthNumber))
End Function

' 目的: 外れ値等の回帰変数に使う "YYYY.mon"（例 2020.apr）へ。
Public Function X13MonthToken(ByVal monthNumber As Long) As String
    X13MonthToken = CStr(YearOf(monthNumber)) & "." & _
        Choose(MonthOf(monthNumber), "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
End Function

' ---- 数値 ----

' 目的: ピリオド小数の文字列をDoubleへ。指数表記可。形式違いは停止（地域設定に依存しない）。
Public Function NumberValue(ByVal text As String) As Double
    If Not NumberRegExp().Test(Trim$(text)) Then Err.Raise vbObjectError + 1500, "NumberValue", "数値として読めません：" & text
    NumberValue = Val(Trim$(text))
End Function

Public Function IsNumberText(ByVal text As String) As Boolean
    IsNumberText = NumberRegExp().Test(Trim$(text))
End Function

' NUMBER_PATTERN の正規表現。1回だけ作って使い回す（保存表・UDGの読取りで数値1つごとに作り直さないため）。
' 決まった1つのパターンでTestするだけなので、使い回しても結果は変わらない。
Private Function NumberRegExp() As Object
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = NUMBER_PATTERN
    End If
    Set NumberRegExp = re
End Function

' 目的: 設定値等を短い文字列にする（例 1.5 → "1.5"）。Str$は常にピリオド小数。
Public Function NumberText(ByVal value As Double) As String
    NumberText = LTrim$(Str$(value))
End Function

' 目的  : 小数点以下decimals桁の固定小数文字列を作る（DATの書出しに使う）。
' 入力  : value、decimals（0～15）
' 出力  : 例 FixedDecimalText(0.000001, 10) = "0.0000010000"
' エラー: 地域設定により想定外の文字が入った場合は停止する。
Public Function FixedDecimalText(ByVal value As Double, ByVal decimals As Long) As String
    Dim pattern As String
    Dim text As String

    If decimals < 0 Or decimals > 15 Then Err.Raise vbObjectError + 1503, "FixedDecimalText", "小数桁は0～15です"
    pattern = "0"
    If decimals > 0 Then pattern = "0." & String$(decimals, "0")
    text = Format$(value, pattern)
    ' Format$はWindowsの地域設定の小数点を使う。DATは常にピリオドにする。
    text = Replace(text, ",", ".")
    If Not IsNumberText(text) Then Err.Raise vbObjectError + 1504, "FixedDecimalText", "数値を固定小数にできません：" & text
    FixedDecimalText = text
End Function

' ---- 識別子とパス ----

' 目的: 識別子に改行・タブ・制御文字が無いことを確かめる（ファイル名・表の照合に使うため）。
Public Sub RequireIdentity(ByVal text As String)
    Dim i As Long
    If Len(text) = 0 Then Err.Raise vbObjectError + 1505, "RequireIdentity", "識別子が空です"
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) >= 0 And AscW(Mid$(text, i, 1)) < 32 Then _
            Err.Raise vbObjectError + 1506, "RequireIdentity", "識別子に制御文字があります：" & text
    Next i
End Sub

' 目的: ドライブ文字から始まるローカルの絶対パスであることを確かめる。
Public Sub RequireLocalPath(ByVal path As String)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z]:[\\/]"
    If Not re.Test(path) Then Err.Raise vbObjectError + 1507, "RequireLocalPath", "ローカルの絶対パスを指定してください：" & path
End Sub

' ---- テキストとバイト ----

' 目的  : テキストファイルを指定の文字コードで読む。
' 入力  : encoding = "UTF8"、"CP932" または "LATIN1"
'         LATIN1 は X-13 HTML版の出力（charset=ISO-8859-1 と宣言）を読むときに使う。1バイトが1文字になる。
' エラー: 読めない場合は元のエラーを送出する。
Public Function ReadText(ByVal path As String, ByVal encoding As String) As String
    Dim stream As Object
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    Select Case UCase$(encoding)
        Case "UTF8": stream.Charset = "utf-8"
        Case "CP932": stream.Charset = "shift_jis"
        Case "LATIN1": stream.Charset = "iso-8859-1"
        Case Else: Err.Raise vbObjectError + 1508, "ReadText", "文字コードはUTF8・CP932・LATIN1のいずれかです"
    End Select
    stream.Open
    stream.LoadFromFile path
    ReadText = stream.ReadText
    stream.Close
    If Left$(ReadText, 1) = ChrW$(&HFEFF) Then ReadText = Mid$(ReadText, 2)
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    CloseQuietly stream
    Err.Raise errorNumber, "ReadText", errorText & "：" & path
End Function

' 目的  : UTF-8（BOMなし）で新しいテキストファイルを書く。既存ファイルは上書きしない。
' エラー: 同名ファイルがある場合は停止する（過去の結果を上書きしないため）。
Public Sub WriteNewText(ByVal path As String, ByVal text As String)
    Dim stream As Object, binary As Object
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    If modMain.FileExists(path) Then Err.Raise vbObjectError + 1509, "WriteNewText", "同名のファイルがあります：" & path
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1
    ' ADODBが付けるBOM（3バイト）を除く。
    If stream.Size >= 3 Then stream.Position = 3
    Set binary = CreateObject("ADODB.Stream")
    binary.Type = 1
    binary.Open
    If Len(text) > 0 Then binary.Write stream.Read
    binary.SaveToFile path, 1
    binary.Close
    stream.Close
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    CloseQuietly binary
    CloseQuietly stream
    Err.Raise errorNumber, "WriteNewText", errorText
End Sub

' 目的: ファイルをバイト配列で読む。空ファイルは停止。
Public Function ReadBytes(ByVal path As String) As Byte()
    Dim stream As Object
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    If stream.Size = 0 Then Err.Raise vbObjectError + 1510, "ReadBytes", "空のファイルです：" & path
    ReadBytes = stream.Read(-1)
    stream.Close
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    CloseQuietly stream
    Err.Raise errorNumber, "ReadBytes", errorText
End Function

' 目的: バイト配列を新しいファイルへ書く。既存ファイルは上書きしない。
Public Sub WriteNewBytes(ByVal path As String, ByRef bytes() As Byte)
    Dim stream As Object
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    If modMain.FileExists(path) Then Err.Raise vbObjectError + 1511, "WriteNewBytes", "同名のファイルがあります：" & path
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile path, 1
    stream.Close
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    CloseQuietly stream
    Err.Raise errorNumber, "WriteNewBytes", errorText
End Sub

' 目的: 完了印などの小さな記録ファイルを読み、末尾の改行を除いた本文を返す。
Public Function ReadMarker(ByVal path As String) As String
    Dim text As String
    text = ReadText(path, "UTF8")
    Do While Right$(text, 1) = vbLf Or Right$(text, 1) = vbCr
        text = Left$(text, Len(text) - 1)
    Loop
    ReadMarker = text
End Function

' 目的: 後始末でストリームを閉じる。閉じる対象が無い・既に閉じている場合の失敗は無視する。
Private Sub CloseQuietly(ByVal stream As Object)
    If stream Is Nothing Then Exit Sub
    On Error Resume Next
    stream.Close
    On Error GoTo 0
End Sub

' 目的  : ファイルを検査付きでコピーする。コピー前後で大きさとCRC32が一致しなければ停止。
' 副作用: destinationを新規作成する（既存なら停止）。
Public Sub CopyFileVerified(ByVal source As String, ByVal destination As String)
    Dim fso As Object, crc As String, size As Double
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(destination) Then Err.Raise vbObjectError + 1512, "CopyFileVerified", "コピー先が既にあります：" & destination
    crc = FileCRC(source)
    size = fso.GetFile(source).Size
    fso.CopyFile source, destination, False
    If fso.GetFile(destination).Size <> size Or FileCRC(destination) <> crc Or FileCRC(source) <> crc Then _
        Err.Raise vbObjectError + 1513, "CopyFileVerified", "コピーの内容が一致しません：" & source
End Sub

' ---- CRC32（偶発的な内容変化の検出。改ざん防止の認証ではない） ----

Private Function ShiftRightOne(ByVal value As Long) As Long
    ShiftRightOne = (value And &H7FFFFFFF) \ 2
    If value < 0 Then ShiftRightOne = ShiftRightOne Or &H40000000
End Function

Private Function ShiftRightEight(ByVal value As Long) As Long
    ShiftRightEight = (value And &H7FFFFFFF) \ &H100&
    If value < 0 Then ShiftRightEight = ShiftRightEight Or &H800000
End Function

' 目的: ファイルのCRC32を16進8桁で返す（r10と同じ値）。
Public Function FileCRC(ByVal path As String) As String
    Static table(0 To 255) As Long, initialized As Boolean
    Dim n As Long, bitIndex As Long, crc As Long, handle As Integer, bytes() As Byte
    Dim remaining As Long, count As Long, i As Long
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    If Not initialized Then
        For n = 0 To 255
            crc = n
            For bitIndex = 1 To 8
                If (crc And 1) <> 0 Then
                    crc = ShiftRightOne(crc) Xor &HEDB88320
                Else
                    crc = ShiftRightOne(crc)
                End If
            Next bitIndex
            table(n) = crc
        Next n
        initialized = True
    End If
    crc = &HFFFFFFFF
    handle = FreeFile
    Open path For Binary Access Read As #handle
    remaining = LOF(handle)
    Do While remaining > 0
        count = remaining
        If count > 65536 Then count = 65536
        ReDim bytes(0 To count - 1)
        Get #handle, , bytes
        For i = 0 To count - 1
            crc = ShiftRightEight(crc) Xor table((crc Xor CLng(bytes(i))) And &HFF&)
        Next i
        remaining = remaining - count
    Loop
    Close #handle
    handle = 0
    FileCRC = Right$("00000000" & Hex$(Not crc), 8)
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    If handle > 0 Then Close #handle
    Err.Raise errorNumber, "FileCRC", errorText & "：" & path
End Function

' ---- X-13の保存表と記録ファイル ----

' 目的  : X-13の保存表の1行を読む。「年月6桁  値」の2列だけをデータ行とする。
' 入力  : text = 1行。monthNumber・value は結果の受け取り先
' 出力  : データ行ならTrue（見出し行・空行はFalse）
' エラー: 年月6桁で始まるのに列数・年月・数値が不正な行、欠損値（-99999）は停止する。
'         壊れた行を読み飛ばすと、月がずれた結果を保存してしまうため。
' 補足  : 2つの正規表現（空白の並び・年月6桁）は1回だけ作って使い回す（保存表の1行ごとに作り直さないため）。
Public Function ParseMonthlyLine(ByVal text As String, ByRef monthNumber As Long, ByRef value As Double) As Boolean
    Static reBlank As Object, reMonth As Object
    Dim fields As Variant, yearValue As Long, monthValue As Long
    If reBlank Is Nothing Then
        Set reBlank = CreateObject("VBScript.RegExp")
        reBlank.Global = True
        reBlank.Pattern = "\s+"
        Set reMonth = CreateObject("VBScript.RegExp")
        reMonth.Pattern = "^[0-9]{6}$"
    End If
    text = Trim$(reBlank.Replace(text, " "))
    If Len(text) = 0 Then Exit Function
    fields = Split(text, " ")
    If Not reMonth.Test(CStr(fields(0))) Then Exit Function
    If UBound(fields) <> 1 Then Err.Raise vbObjectError + 1530, "ParseMonthlyLine", "保存表の行は年月と値の2列です：" & text
    yearValue = CLng(fields(0)) \ 100
    monthValue = CLng(fields(0)) Mod 100
    monthNumber = MonthIndex(yearValue, monthValue)
    value = NumberValue(CStr(fields(1)))
    If value = -99999# Then Err.Raise vbObjectError + 1531, "ParseMonthlyLine", "保存表に欠損値があります：" & text
    ParseMonthlyLine = True
End Function

' 目的  : X-13の保存表（D10・D11・D16・D18等）を読み、行数と年月の連続を確かめる。
' 入力  : path、firstMonth（最初の月番号）、expectedRows（期待する行数。D11は実績月数、
'         D10・D16・D18は実績＋予測の月数）
' 出力  : 2次元配列 (1 To 行数, 1 To 2)。1列目=月番号、2列目=値
' エラー: 開始月の違い、重複・欠落・逆順、行数の過不足は停止する。
Public Function ReadMonthlyTable(ByVal path As String, ByVal firstMonth As Long, ByVal expectedRows As Long) As Variant
    Dim lines As Variant, line As Variant, monthNumber As Long, value As Double
    Dim data() As Variant, count As Long
    If expectedRows < 1 Then Err.Raise vbObjectError + 1532, "ReadMonthlyTable", "期待行数が不正です"
    ReDim data(1 To expectedRows, 1 To 2)
    lines = Split(Replace(ReadText(path, "UTF8"), vbCr, ""), vbLf)
    For Each line In lines
        If ParseMonthlyLine(CStr(line), monthNumber, value) Then
            count = count + 1
            If count > expectedRows Then Err.Raise vbObjectError + 1533, "ReadMonthlyTable", _
                "保存表の行数が期待（" & expectedRows & "行）より多いです：" & path
            If monthNumber <> firstMonth + count - 1 Then Err.Raise vbObjectError + 1534, "ReadMonthlyTable", _
                "保存表の年月が連続していません（" & MonthText(firstMonth + count - 1) & "の位置に" & MonthText(monthNumber) & "）：" & path
            data(count, 1) = monthNumber
            data(count, 2) = value
        End If
    Next line
    If count <> expectedRows Then Err.Raise vbObjectError + 1535, "ReadMonthlyTable", _
        "保存表の行数が" & count & "行です（期待" & expectedRows & "行）：" & path
    ReadMonthlyTable = data
End Function

' 目的  : 「キー=値」形式の記録ファイルを読む。
' エラー: 形式違い・キーの重複は停止する。
Public Function ReadMetaFile(ByVal path As String) As Object
    Dim result As Object, lines As Variant, line As Variant, pos As Long
    Set result = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(ReadText(path, "UTF8"), vbCr, ""), vbLf)
    For Each line In lines
        If Len(line) > 0 Then
            pos = InStr(line, "=")
            If pos < 2 Then Err.Raise vbObjectError + 1536, "ReadMetaFile", "記録ファイルの形式が不正です：" & path
            If result.Exists(Left$(line, pos - 1)) Then Err.Raise vbObjectError + 1537, "ReadMetaFile", "記録ファイルのキーが重複しています：" & path
            result.Add Left$(line, pos - 1), Mid$(line, pos + 1)
        End If
    Next line
    Set ReadMetaFile = result
End Function

' 目的  : Dictionaryを「キー=値」形式の新しいファイルへ書く（UTF-8）。既存ファイルは上書きしない。
' エラー: 値に改行・制御文字がある場合は停止する（1行1項目の形式を守るため）。
Public Sub WriteMetaFile(ByVal path As String, ByVal meta As Object)
    Dim key As Variant, body As String
    For Each key In meta.Keys
        RequireIdentity CStr(key)
        If Len(CStr(meta(key))) > 0 Then RequireIdentity CStr(meta(key))
        body = body & CStr(key) & "=" & CStr(meta(key)) & vbCrLf
    Next key
    WriteNewText path, body
End Sub

' 目的  : ファイルがASCII文字（タブ・改行を含む）だけか確かめる。SPCとMTAの検査に使う。
' エラー: ASCII以外があれば停止する。
Public Sub RequireAsciiFile(ByVal path As String)
    Dim bytes() As Byte, i As Long
    bytes = ReadBytes(path)
    For i = LBound(bytes) To UBound(bytes)
        If Not (bytes(i) = 9 Or bytes(i) = 10 Or bytes(i) = 13 Or (bytes(i) >= 32 And bytes(i) <= 126)) Then _
            Err.Raise vbObjectError + 1538, "RequireAsciiFile", "ASCII以外の文字があります（" & i + 1 & "バイト目）：" & path
    Next i
End Sub

' ---- シート ----

' 目的: シートの最後の使用行の次の行番号（空のシートは2）。
Public Function NextRow(ByVal ws As Worksheet) As Long
    Dim last As Range
    Set last = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    NextRow = 2
    If Not last Is Nothing Then
        If last.Row >= ws.Rows.Count Then Err.Raise vbObjectError + 1514, "NextRow", "シートの行数が上限です：" & ws.Name
        If last.Row >= 2 Then NextRow = last.Row + 1
    End If
End Function

' 目的  : 1行目の見出しを確かめる。空のシートには見出しを書く。
' エラー: 既存の見出しが違う場合は停止する（別の用途のシートへ書き込まないため）。
Public Sub Headers(ByVal ws As Worksheet, ByVal names As Variant)
    Dim c As Long
    If ws.ProtectContents Then Err.Raise vbObjectError + 1515, "Headers", "シートが保護されています：" & ws.Name
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        ws.Cells(1, 1).Resize(1, UBound(names) + 1).Value2 = names
        ws.Rows(1).Font.Bold = True
    End If
    For c = 0 To UBound(names)
        If IsError(ws.Cells(1, c + 1).Value2) Then Err.Raise vbObjectError + 1516, "Headers", "見出しがエラー値です：" & ws.Name
        If CStr(ws.Cells(1, c + 1).Value2) <> CStr(names(c)) Then _
            Err.Raise vbObjectError + 1517, "Headers", "見出しが想定と違います：" & ws.Name & " " & CStr(names(c))
    Next c
End Sub

' 目的  : シートを空にして見出しを書き直す（集計表など、毎回作り直す表に使う）。
Public Sub ResetSheet(ByVal ws As Worksheet, ByVal names As Variant)
    If ws.ProtectContents Then Err.Raise vbObjectError + 1518, "ResetSheet", "シートが保護されています：" & ws.Name
    ws.Cells.Clear
    ws.Cells(1, 1).Resize(1, UBound(names) + 1).Value2 = names
    ws.Rows(1).Font.Bold = True
End Sub

' 目的  : 2次元配列を末尾へ追記する。書いた開始行と行数を返す。
' 入力  : data(1 To n, 1 To m)。numericColumns は数値として書く列番号の配列（省略時は全列を文字列書式）。
' エラー: 失敗した場合は書いた範囲を消してから停止する。
Public Sub AppendRows(ByVal ws As Worksheet, ByVal data As Variant, ByRef startRow As Long, ByRef count As Long, _
                      Optional ByVal numericColumns As Variant)
    Dim width As Long, c As Variant
    Dim errorNumber As Long, errorText As String

    On Error GoTo Failed
    If ws.ProtectContents Then Err.Raise vbObjectError + 1519, "AppendRows", "シートが保護されています：" & ws.Name
    count = UBound(data, 1)
    width = UBound(data, 2)
    startRow = NextRow(ws)
    If count > ws.Rows.Count - startRow + 1 Then Err.Raise vbObjectError + 1520, "AppendRows", "シートの行数が足りません：" & ws.Name
    ws.Cells(startRow, 1).Resize(count, width).NumberFormat = "@"
    If Not IsMissing(numericColumns) Then
        For Each c In numericColumns
            ws.Cells(startRow, CLng(c)).Resize(count, 1).NumberFormat = "General"
        Next c
    End If
    ws.Cells(startRow, 1).Resize(count, width).Value2 = data
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    If startRow > 1 And count > 0 And width > 0 Then ws.Cells(startRow, 1).Resize(count, width).ClearContents
    count = 0
    Err.Raise errorNumber, "AppendRows", errorText
End Sub

' 目的: Collection（各要素は0始まりの1次元配列）を、シートへ書ける2次元配列へ変換する。
Public Function RowsToArray(ByVal rows As Collection, ByVal width As Long) As Variant
    Dim result() As Variant, i As Long, c As Long, item As Variant
    If rows.Count = 0 Then Err.Raise vbObjectError + 1521, "RowsToArray", "書き込む行がありません"
    ReDim result(1 To rows.Count, 1 To width)
    For i = 1 To rows.Count
        item = rows(i)
        For c = 1 To width
            result(i, c) = item(c - 1)
        Next c
    Next i
    RowsToArray = result
End Function
