Attribute VB_Name = "modIIPTextTable"
Option Explicit

' 役割      : CSV・TSVを、指定または判別した文字コードで厳密に読む。ExcelにCSVを開かせないため、
'             列の型の自動変換や数式の実行は起きない。
' 主な入口  : ReadTextTable（取込用の表）/ DecodeFile / DetectEncoding / DecodeBytes / Utf8Length
' 依存先    : modX13IO
' 文字コード: UTF8（BOM有無とも可）、CP932（Shift_JIS）。
'             CP932は、読んだ文字列を同じ文字コードへ戻して元のバイト列と一致することを確かめる。
'             UTF-8は、バイト列が規格どおりか（不正な並び・NUL・サロゲートが無いか）を確かめる。

Private Const MAX_BYTES As Double = 67108864#   ' 64MiB
Private Const MAX_CELLS As Double = 2000000#
Private Const MAX_ROWS As Long = 250001
Private Const MAX_COLS As Long = 2048

' 目的  : 取込設定で指定された区切り文字・文字コードで表を読み、2次元配列（1始まり）を返す。
' エラー: 文字コード違い、引用符の不正、上限超過は停止。
Public Function ReadTextTable(ByVal path As String, ByVal encoding As String, ByVal delimiter As String) As Variant
    Dim text As String, records As Collection, record As Variant, table() As Variant
    Dim row As Long, col As Long, width As Long, separator As String
    If delimiter = "comma" Then
        separator = ","
    ElseIf delimiter = "tab" Then
        separator = vbTab
    Else
        Fail "区切り文字はcommaまたはtabを指定してください。"
    End If
    text = DecodeFile(path, UCase$(encoding))
    Set records = ParseRecords(text, separator)
    For Each record In records
        If UBound(record) + 1 > width Then width = UBound(record) + 1
    Next record
    If CDbl(records.Count) * width > MAX_CELLS Then Fail "表の大きさが上限（200万セル）を超えています。"
    ReDim table(1 To records.Count, 1 To width)
    For row = 1 To records.Count
        record = records(row)
        For col = 0 To UBound(record): table(row, col + 1) = record(col): Next col
    Next row
    ReadTextTable = table
End Function

' 目的  : ファイルを読み、文字コードを指定して文字列にする。encoding="AUTO"なら判別する。
Public Function DecodeFile(ByVal path As String, ByVal encoding As String) As String
    Dim bytes() As Byte, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Fail "ファイルが見つかりません：" & path
    If fso.GetFile(path).Size > MAX_BYTES Then Fail "64MiBを超えるファイルは読めません：" & path
    bytes = modX13IO.ReadBytes(path)
    If UCase$(encoding) = "AUTO" Then encoding = DetectEncoding(bytes)
    DecodeFile = DecodeBytes(bytes, encoding)
End Function

' 目的  : バイト列の文字コードを判別する。
' 出力  : "UTF8_BOM"（先頭がEF BB BF）、"UTF8"（BOM無しで規格どおりのUTF-8）、"CP932"（それ以外でCP932として往復一致）
' エラー: どれにも当てはまらなければ停止（推測で読まない）。
' 補足  : 日本語を含むCP932のファイルは、UTF-8としては不正な並びになるため取り違えない。
'         ASCIIだけのファイルはUTF8と判定される（どちらで読んでも同じ文字列）。
Public Function DetectEncoding(ByRef bytes() As Byte) As String
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            If Utf8Length(bytes, 3) < 0 Then Fail "BOM付きUTF-8として不正なバイト列があります。"
            DetectEncoding = "UTF8_BOM"
            Exit Function
        End If
    End If
    If Utf8Length(bytes, 0) >= 0 Then
        DetectEncoding = "UTF8"
    ElseIf Cp932RoundTrips(bytes) Then
        DetectEncoding = "CP932"
    Else
        Fail "文字コードをUTF-8・CP932のどちらとも判別できません。"
    End If
End Function

' 目的  : バイト列を指定の文字コードで文字列にする。
' 入力  : encoding = "UTF8"（BOMがあれば除く）、"UTF8_BOM"、"CP932"
' エラー: 指定と内容が合わない場合は停止。
Public Function DecodeBytes(ByRef bytes() As Byte, ByVal encoding As String) As String
    Dim start As Long, expectedLength As Long, text As String
    Select Case UCase$(encoding)
        Case "UTF8", "UTF8_BOM"
            If UBound(bytes) >= 2 Then
                If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then start = 3
            End If
            If UCase$(encoding) = "UTF8_BOM" And start = 0 Then Fail "BOM付きUTF-8ではありません。"
            expectedLength = Utf8Length(bytes, start)
            If expectedLength < 0 Then Fail "UTF-8として不正なバイト列があります。文字コードを確認してください。"
            text = StreamText(bytes, "utf-8")
            If Left$(text, 1) = ChrW$(&HFEFF) Then text = Mid$(text, 2)
            If Len(text) <> expectedLength Then Fail "UTF-8の読み取り結果の文字数が一致しません。"
        Case "CP932"
            If Not Cp932RoundTrips(bytes) Then Fail "CP932の文字を正確に復元できません。文字コードを確認してください。"
            text = StreamText(bytes, "shift_jis")
        Case Else
            Fail "文字コードはUTF8またはCP932を指定してください：" & encoding
    End Select
    If InStr(text, ChrW$(0)) > 0 Then Fail "ファイルにNUL文字が含まれています。"
    DecodeBytes = text
End Function

' 目的  : UTF-8のバイト列を厳密に検査し、UTF-16での文字数を返す。不正なら-1。
' 規則  : NUL、冗長表現、サロゲート領域、U+10FFFF超、途中で切れた文字を不正とする。
Public Function Utf8Length(ByRef bytes() As Byte, ByVal start As Long) As Long
    Dim i As Long, j As Long, k As Long, c As Long, lower As Long, upper As Long, n As Long, count As Long
    n = UBound(bytes)
    i = start
    Do While i <= n
        c = bytes(i)
        lower = &H80: upper = &HBF
        If c = 0 Then
            Utf8Length = -1: Exit Function
        ElseIf c < &H80 Then
            k = 0
        ElseIf c >= &HC2 And c <= &HDF Then
            k = 1
        ElseIf c >= &HE0 And c <= &HEF Then
            k = 2
            If c = &HE0 Then lower = &HA0
            If c = &HED Then upper = &H9F
        ElseIf c >= &HF0 And c <= &HF4 Then
            k = 3
            If c = &HF0 Then lower = &H90
            If c = &HF4 Then upper = &H8F
        Else
            Utf8Length = -1: Exit Function
        End If
        If i + k > n Then Utf8Length = -1: Exit Function
        For j = 1 To k
            c = bytes(i + j)
            If j > 1 Then lower = &H80: upper = &HBF
            If c < lower Or c > upper Then Utf8Length = -1: Exit Function
        Next j
        If k = 3 Then count = count + 2 Else count = count + 1
        i = i + k + 1
    Loop
    Utf8Length = count
End Function

' CP932として読み、CP932へ戻したバイト列が元と一致するか。
Private Function Cp932RoundTrips(ByRef bytes() As Byte) As Boolean
    Dim text As String, stream As Object, roundTrip() As Byte, i As Long
    text = StreamText(bytes, "shift_jis")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "shift_jis": stream.Open
    stream.WriteText text
    stream.Position = 0: stream.Type = 1
    roundTrip = stream.Read(-1)
    stream.Close
    If UBound(roundTrip) <> UBound(bytes) Then Exit Function
    For i = 0 To UBound(bytes)
        If bytes(i) <> roundTrip(i) Then Exit Function
    Next i
    Cp932RoundTrips = True
End Function

' バイト列をADODB.Streamで文字列にする（BOMの扱いは呼出側）。
Private Function StreamText(ByRef bytes() As Byte, ByVal charset As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open
    stream.Write bytes
    stream.Position = 0: stream.Type = 2: stream.Charset = charset
    StreamText = stream.ReadText(-1)
    stream.Close
End Function

' RFC4180の引用符に対応して1レコードずつ分割する。列数が違うタイトル行や空行も保持する。
Private Function ParseRecords(ByVal text As String, ByVal separator As String) As Collection
    Dim records As New Collection, fields() As String, count As Long
    Dim pos As Long, length As Long, first As Long, field As String, ch As String, escaped As Boolean
    length = Len(text): pos = 1
    Do While pos <= length
        ReDim fields(0 To MAX_COLS - 1): count = 0
        Do
            first = pos: escaped = False
            If Mid$(text, pos, 1) = """" Then
                pos = pos + 1: first = pos
                Do
                    If pos > length Then Fail "引用符が閉じていません（" & records.Count + 1 & "行目）。"
                    If Mid$(text, pos, 1) = """" Then
                        If Mid$(text, pos + 1, 1) <> """" Then Exit Do
                        escaped = True: pos = pos + 2
                    Else
                        pos = pos + 1
                    End If
                Loop
                field = Mid$(text, first, pos - first)
                If escaped Then field = Replace(field, """""", """")
                pos = pos + 1: ch = Mid$(text, pos, 1)
                If Len(ch) > 0 And ch <> separator And ch <> vbCr And ch <> vbLf Then _
                    Fail "閉じ引用符の後に不正な文字があります（" & records.Count + 1 & "行目）。"
            Else
                Do While pos <= length
                    ch = Mid$(text, pos, 1)
                    If ch = separator Or ch = vbCr Or ch = vbLf Then Exit Do
                    If ch = """" Then Fail "引用符の使い方が不正です（" & records.Count + 1 & "行目）。"
                    pos = pos + 1
                Loop
                field = Mid$(text, first, pos - first)
            End If
            If count >= MAX_COLS Then Fail "列数の上限（2048列）を超えています。"
            fields(count) = field: count = count + 1
            If pos > length Then Exit Do
            ch = Mid$(text, pos, 1): pos = pos + 1
            If ch = vbCr Then
                If Mid$(text, pos, 1) <> vbLf Then Fail "改行にはCRLFまたはLFを使用してください。"
                pos = pos + 1: Exit Do
            ElseIf ch = vbLf Then
                Exit Do
            End If
        Loop
        If records.Count >= MAX_ROWS Then Fail "行数の上限（250001行）を超えています。"
        ReDim Preserve fields(0 To count - 1): records.Add fields
    Loop
    If records.Count = 0 Then Fail "表にデータがありません。"
    Set ParseRecords = records
End Function

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2400, "modIIPTextTable", message
End Sub
