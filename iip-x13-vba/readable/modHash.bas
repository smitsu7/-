Attribute VB_Name = "modHash"
Option Explicit

' 役割      : ファイルのSHA-256（16進64桁、小文字）を計算する。祝日CSVの原本と、
'             保存したコピーが同一であることの確認に使う。
' 主な入口  : FileSHA256 / BytesSHA256 / SelfTestPassed
' 依存先    : modX13IO（ReadBytes）
' 方式      : FIPS 180-4のSHA-256をVBAのLong（符号付き32ビット）で実装した。
'             32ビットの加算・論理シフトは、桁あふれを起こさない補助関数で行う。
'             初回の計算前に既知の答え（"abc"等）と照合し、一致しなければ停止する。
'             .NETやPowerShellを呼ばないため、セキュリティ設定に依存しない。

Private mPow2(0 To 30) As Long
Private mRoundConstants(0 To 63) As Long
Private mReady As Boolean
Private mSelfTested As Boolean

' 目的  : ファイルのSHA-256を返す。
' エラー: 自己点検に失敗した場合、ファイルが読めない場合は停止する。
Public Function FileSHA256(ByVal path As String) As String
    Dim bytes() As Byte
    RequireSelfTest
    bytes = modX13IO.ReadBytes(path)
    FileSHA256 = DigestHex(bytes, UBound(bytes) - LBound(bytes) + 1)
End Function

' 目的: バイト配列のSHA-256を返す。長さ0の場合はlength=0を渡す。
Public Function BytesSHA256(ByRef bytes() As Byte, ByVal length As Long) As String
    RequireSelfTest
    BytesSHA256 = DigestHex(bytes, length)
End Function

' 目的: 既知の答えと照合する（FIPS 180-2 付録の例と空文字列）。
Public Function SelfTestPassed() As Boolean
    Dim noBytes() As Byte, shortText() As Byte, twoBlockText() As Byte
    ReDim noBytes(0 To 0)
    shortText = AsciiBytes("abc")
    twoBlockText = AsciiBytes("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq")
    Prepare
    SelfTestPassed = (DigestHex(shortText, 3) = _
            "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad") And _
        (DigestHex(noBytes, 0) = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855") And _
        (DigestHex(twoBlockText, 56) = _
            "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1")
End Function

Private Sub RequireSelfTest()
    If mSelfTested Then Exit Sub
    If Not SelfTestPassed() Then Err.Raise vbObjectError + 1590, "modHash", "SHA-256の自己点検に失敗しました。計算を中止します。"
    mSelfTested = True
End Sub

Private Function AsciiBytes(ByVal text As String) As Byte()
    Dim result() As Byte, i As Long
    ReDim result(0 To Len(text) - 1)
    For i = 1 To Len(text)
        result(i - 1) = CByte(Asc(Mid$(text, i, 1)))
    Next i
    AsciiBytes = result
End Function

' 2の累乗表と、SHA-256の64個の丸め定数を用意する。
Private Sub Prepare()
    Dim i As Long, constants As Variant
    If mReady Then Exit Sub
    mPow2(0) = 1
    For i = 1 To 30
        mPow2(i) = mPow2(i - 1) * 2
    Next i
    constants = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    For i = 0 To 63
        mRoundConstants(i) = CLng(constants(i))
    Next i
    mReady = True
End Sub

' 論理右シフト（0 < n < 32）。符号ビットを外して割り、外したビットを正しい位置へ戻す。
Private Function ShiftRight(ByVal x As Long, ByVal n As Long) As Long
    ShiftRight = (x And &H7FFFFFFF) \ mPow2(n)
    If x < 0 Then ShiftRight = ShiftRight Or mPow2(31 - n)
End Function

' 論理左シフト（0 < n < 31）。押し出される上位ビットを捨て、符号ビットになる1ビットを別に扱う。
Private Function ShiftLeft(ByVal x As Long, ByVal n As Long) As Long
    ShiftLeft = (x And (mPow2(31 - n) - 1)) * mPow2(n)
    If (x And mPow2(31 - n)) <> 0 Then ShiftLeft = ShiftLeft Or &H80000000
End Function

Private Function RotateRight(ByVal x As Long, ByVal n As Long) As Long
    RotateRight = ShiftRight(x, n) Or ShiftLeft(x, 32 - n)
End Function

' 2^32を法とする加算。下位16ビットと上位16ビットに分け、桁上がりを手で処理する。
Private Function Add32(ByVal a As Long, ByVal b As Long) As Long
    Dim low As Long, high As Long
    low = (a And &HFFFF&) + (b And &HFFFF&)
    high = (ShiftRight(a, 16) + ShiftRight(b, 16) + low \ &H10000) And &HFFFF&
    If high >= &H8000& Then
        Add32 = ((high - &H10000) * &H10000) Or (low And &HFFFF&)
    Else
        Add32 = (high * &H10000) Or (low And &HFFFF&)
    End If
End Function

Private Function BigEndianWord(ByRef block() As Byte, ByVal offset As Long) As Long
    BigEndianWord = (CLng(block(offset) And &H7F) * &H1000000) Or (CLng(block(offset + 1)) * &H10000) Or _
        (CLng(block(offset + 2)) * &H100&) Or CLng(block(offset + 3))
    If (block(offset) And &H80) <> 0 Then BigEndianWord = BigEndianWord Or &H80000000
End Function

' 目的: 長さlengthバイトのメッセージのSHA-256を16進で返す。
Private Function DigestHex(ByRef message() As Byte, ByVal length As Long) As String
    Dim padded() As Byte, paddedLength As Long, bitLength As Double, i As Long, blockStart As Long
    Dim w(0 To 63) As Long, h(0 To 7) As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, choice As Long, majority As Long, temp1 As Long, temp2 As Long, t As Long
    Dim base As Long

    Prepare
    base = 0
    If length > 0 Then base = LBound(message)
    ' 末尾に0x80、0の詰め物、ビット長（64ビット、ビッグエンディアン）を付けて64バイトの倍数にする。
    paddedLength = ((length + 8) \ 64 + 1) * 64
    ReDim padded(0 To paddedLength - 1)
    For i = 0 To length - 1
        padded(i) = message(base + i)
    Next i
    padded(length) = &H80
    bitLength = CDbl(length) * 8#
    For i = 0 To 7
        padded(paddedLength - 1 - i) = CByte(bitLength - Int(bitLength / 256#) * 256#)
        bitLength = Int(bitLength / 256#)
    Next i

    h(0) = &H6A09E667: h(1) = &HBB67AE85: h(2) = &H3C6EF372: h(3) = &HA54FF53A
    h(4) = &H510E527F: h(5) = &H9B05688C: h(6) = &H1F83D9AB: h(7) = &H5BE0CD19

    For blockStart = 0 To paddedLength - 1 Step 64
        For t = 0 To 15
            w(t) = BigEndianWord(padded, blockStart + t * 4)
        Next t
        For t = 16 To 63
            s0 = RotateRight(w(t - 15), 7) Xor RotateRight(w(t - 15), 18) Xor ShiftRight(w(t - 15), 3)
            s1 = RotateRight(w(t - 2), 17) Xor RotateRight(w(t - 2), 19) Xor ShiftRight(w(t - 2), 10)
            w(t) = Add32(Add32(w(t - 16), s0), Add32(w(t - 7), s1))
        Next t
        a = h(0): b = h(1): c = h(2): d = h(3): e = h(4): f = h(5): g = h(6): hh = h(7)
        For t = 0 To 63
            s1 = RotateRight(e, 6) Xor RotateRight(e, 11) Xor RotateRight(e, 25)
            choice = (e And f) Xor ((Not e) And g)
            temp1 = Add32(Add32(Add32(hh, s1), Add32(choice, mRoundConstants(t))), w(t))
            s0 = RotateRight(a, 2) Xor RotateRight(a, 13) Xor RotateRight(a, 22)
            majority = (a And b) Xor (a And c) Xor (b And c)
            temp2 = Add32(s0, majority)
            hh = g: g = f: f = e
            e = Add32(d, temp1)
            d = c: c = b: b = a
            a = Add32(temp1, temp2)
        Next t
        h(0) = Add32(h(0), a): h(1) = Add32(h(1), b): h(2) = Add32(h(2), c): h(3) = Add32(h(3), d)
        h(4) = Add32(h(4), e): h(5) = Add32(h(5), f): h(6) = Add32(h(6), g): h(7) = Add32(h(7), hh)
    Next blockStart

    For i = 0 To 7
        DigestHex = DigestHex & LCase$(Right$("00000000" & Hex$(h(i)), 8))
    Next i
End Function
