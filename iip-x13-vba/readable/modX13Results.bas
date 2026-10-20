Attribute VB_Name = "modX13Results"
Option Explicit

' 役割      : X-13の出力ファイルを読み、比較に使う値を取り出す。値の判断（順位・採否）はしない。
'             UDG（キー: 値の診断要約）… AIC・AICC・BIC・HQ、最終ARIMA次数、収束、回帰係数とt値、
'                                         予測誤差率（aape）、X-13の結合F検定、自動検出の外れ値
'             OUT（主出力）              … 最終ARIMA次数、変換、統計量、自動選択の探索履歴
'             RCM（回帰係数の共分散）    … 変数群のWald検定
' 主な入口  : ReadUdgValues / ReadOutputEvidence / ParseSearchHistory / RegressionCoefficients
'             DetectedOutliers / WaldGroupTest / ExpectedCalendarGroups / CalendarGroup / EnsureD18
'             ReadValue / NumericValue / Numbers
' 依存先    : modX13IO、modMain
' 注意      : 自動選択の探索中に表示されるBIC2は、探索用の値として探索履歴にだけ記録する。
'             最終候補のAIC・BIC・AICC・HQの代わりには使わない（取得できない値は空のまま）。

' 目的  : UDGを読み、キー→値のDictionaryを返す。
' 規則  : 同じキーが2回以上現れた場合（自動選択の探索記録の繰返し等）は、どちらか一方を選ばずに
'         キーごと除く（あいまいな値で比較しないため）。
Public Function ReadUdgValues(ByVal path As String) As Object
    Dim values As Object, ambiguous As Object, lines As Variant, line As Variant, pos As Long, key As String
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = vbTextCompare
    Set ambiguous = CreateObject("Scripting.Dictionary")
    ambiguous.CompareMode = vbTextCompare
    lines = Split(Replace(modX13IO.ReadText(path, "UTF8"), vbCr, ""), vbLf)
    For Each line In lines
        pos = InStr(CStr(line), ":")
        If pos > 1 Then
            key = Trim$(Left$(CStr(line), pos - 1))
            If values.Exists(key) Then
                values.Remove key
                ambiguous.Add key, True
            ElseIf Not ambiguous.Exists(key) Then
                values.Add key, Trim$(Mid$(CStr(line), pos + 1))
            End If
        End If
    Next line
    Set ReadUdgValues = values
End Function

' 目的: UDGの値を文字列で返す。無ければ空文字。
Public Function ReadValue(ByVal values As Object, ByVal key As String) As String
    If values.Exists(key) Then ReadValue = Trim$(CStr(values(key)))
End Function

' 目的: UDGの値を数値で返す。無い・数値でない場合はEmpty（0にしない）。
Public Function NumericValue(ByVal values As Object, ByVal key As String) As Variant
    If Not values.Exists(key) Then Exit Function
    If modX13IO.IsNumberText(CStr(values(key))) Then NumericValue = modX13IO.NumberValue(CStr(values(key)))
End Function

' 目的: 空白で区切られた数値の並びを配列にする。数値でない要素があればEmpty。
Public Function Numbers(ByVal text As String) As Variant
    Dim re As Object, parts As Variant, i As Long, result() As Double
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"
    parts = Split(re.Replace(Trim$(text), " "), " ")
    ReDim result(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Not modX13IO.IsNumberText(CStr(parts(i))) Then Exit Function
        result(i) = modX13IO.NumberValue(CStr(parts(i)))
    Next i
    Numbers = result
End Function

' 目的  : OUTとSPCから、最終ARIMA次数・実際の変換・統計量を取り出す。
' 出力  : Dictionary（X13_ACTUAL_MODEL, X13_ACTUAL_TRANSFORM, X13_ACTUAL_AIC, X13_ACTUAL_AICC, X13_ACTUAL_BIC, X13_ACTUAL_HQ）
'         見つからない項目は入れない。transform=autoの場合は、OUTの「prefers log/no transformation」から決める。
Public Function ReadOutputEvidence(ByVal folder As String, ByVal stem As String) As Object
    Dim output As String, spec As String, lines As Variant, line As Variant, result As Object, re As Object
    Dim matches As Object, model As String, transform As String, name As Variant, pattern As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If Not modMain.FileExists(modMain.CombinePath(folder, stem & ".out")) Then
        Set ReadOutputEvidence = result
        Exit Function
    End If
    output = modX13IO.ReadText(modMain.CombinePath(folder, stem & ".out"), "UTF8")
    spec = modX13IO.ReadText(modMain.CombinePath(folder, stem & ".spc"), "UTF8")
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "transform\s*\{\s*function\s*=\s*(log|none|auto)\b"
    Set matches = re.Execute(spec)
    If matches.Count > 0 Then transform = LCase$(CStr(matches(0).SubMatches(0)))
    If transform = "auto" Then
        transform = ""
        If InStr(1, output, "prefers log transformation", vbTextCompare) > 0 Then transform = "log"
        If InStr(1, output, "prefers no transformation", vbTextCompare) > 0 Then transform = "none"
    End If
    If transform = "log" Or transform = "none" Then result.Add "X13_ACTUAL_TRANSFORM", transform
    lines = Split(Replace(output, vbCr, ""), vbLf)
    For Each line In lines
        re.Pattern = "^\s*ARIMA Model:\s*(\([0-9 ]+\)\s*\([0-9 ]+\))\s*$"
        Set matches = re.Execute(CStr(line))
        If matches.Count > 0 Then model = Trim$(CStr(matches(0).SubMatches(0)))
        For Each name In Array("AIC", "AICC", "BIC", "HQ")
            Select Case CStr(name)
                Case "AICC": pattern = "^\s*AICC\s*\(F-corrected-AIC\)\s+([+-]?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?)\s*$"
                Case "HQ": pattern = "^\s*Hannan Quinn\s+([+-]?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?)\s*$"
                Case Else: pattern = "^\s*" & CStr(name) & "\s+([+-]?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?)\s*$"
            End Select
            re.Pattern = pattern
            Set matches = re.Execute(CStr(line))
            If matches.Count > 0 Then result("X13_ACTUAL_" & CStr(name)) = CStr(matches(0).SubMatches(0))
        Next name
    Next line
    If Len(model) > 0 Then result.Add "X13_ACTUAL_MODEL", model
    Set ReadOutputEvidence = result
End Function

' 目的  : 自動選択（automdl）の詳細出力を読み、探索中に推定した各ARIMA次数を順に返す。
' 出力  : Collection。要素はDictionary（Sequence, Stage, Pass, Model, D, SD, State, BIC, BIC2, Rank, SourceLine）
'         BIC2の順位は、同じ探索段階・探索回・差分次数の中で付ける。探索中に出ないAIC等は持たない。
Public Function ParseSearchHistory(ByVal outputText As String) As Collection
    Dim records As New Collection, item As Object, current As Object, previous As Variant
    Dim lines As Variant, line As String, i As Long, re As Object, matches As Object
    Dim stage As String, unitPass As Long, armaPass As Long, pass As Long, rank As Long
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    lines = Split(Replace(outputText, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        line = CStr(lines(i))
        If InStr(line, "ARIMA Estimates (H-R) for Unit Root Identification") > 0 Then
            stage = "差分次数の判定"
            re.Pattern = "Model No\.\s*1\s*$"
            If re.Test(line) Then unitPass = unitPass + 1
            pass = unitPass
            Set current = Nothing
        ElseIf InStr(line, "ARIMA Estimates and Likelihood Values for ARMA Order Identification") > 0 Then
            stage = "ARMA次数の探索"
            armaPass = armaPass + 1
            pass = armaPass
            Set current = Nothing
        ElseIf InStr(line, "Best Five ARIMA Models") > 0 Or InStr(line, "Final Checks for Identified Model") > 0 Then
            Set current = Nothing
        End If
        re.Pattern = "^\s*(Model Estimated\s*:|Estimation errors cause model)\s*\(\s*([0-9]+)\s+([0-9]+)\s+([0-9]+)\s*\)(\s*\(\s*([0-9]+)\s+([0-9]+)\s+([0-9]+)\s*\))?"
        Set matches = re.Execute(line)
        If matches.Count > 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "Sequence", records.Count + 1
            item.Add "Stage", IIf(Len(stage) = 0, "その他のモデル試行", stage)
            item.Add "Pass", pass
            item.Add "Model", "(" & matches(0).SubMatches(1) & " " & matches(0).SubMatches(2) & " " & matches(0).SubMatches(3) & ")(" & _
                CLng(Val(CStr(matches(0).SubMatches(5)))) & " " & CLng(Val(CStr(matches(0).SubMatches(6)))) & " " & _
                CLng(Val(CStr(matches(0).SubMatches(7)))) & ")"
            item.Add "D", CLng(matches(0).SubMatches(2))
            item.Add "SD", CLng(Val(CStr(matches(0).SubMatches(6))))
            item.Add "State", IIf(InStr(line, "Estimation errors") > 0, "推定失敗", "推定済")
            item.Add "BIC", Empty
            item.Add "BIC2", Empty
            item.Add "Rank", Empty
            item.Add "SourceLine", i + 1
            item.Add "Group", item("Stage") & "|" & pass & "|" & item("D") & "|" & item("SD")
            records.Add item
            Set current = item
        ElseIf Not current Is Nothing Then
            re.Pattern = "^\s*(BIC2|BIC)\s*:\s*([+-]?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?)\s*$"
            Set matches = re.Execute(line)
            If matches.Count > 0 Then current(CStr(matches(0).SubMatches(0))) = modX13IO.NumberValue(CStr(matches(0).SubMatches(1)))
        End If
    Next i
    For Each previous In records
        If Not IsEmpty(previous("BIC2")) And previous("State") = "推定済" Then
            rank = 1
            For Each item In records
                If item("Group") = previous("Group") And Not IsEmpty(item("BIC2")) Then
                    If CDbl(item("BIC2")) < CDbl(previous("BIC2")) Then rank = rank + 1
                End If
            Next item
            previous("Rank") = rank
        End If
    Next previous
    Set ParseSearchHistory = records
End Function

' 目的  : UDGの回帰係数（nregの後に続く「名前$変数: 係数 標準誤差 t値」）を取り出す。
' 出力  : Dictionary（キー=UDGのキー、値=Array(係数, 標準誤差, t値)）。completeに件数がnregと一致したか。
'         派生係数（nregderivedの後、例 Sat/Sun）は推定した係数ではないため含めない。
Public Function RegressionCoefficients(ByVal values As Object, ByRef complete As Boolean) As Object
    Dim result As Object, key As Variant, remaining As Long, expected As Long, inBlock As Boolean, nums As Variant
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    complete = False
    If Not values.Exists("nreg") Then GoTo Done
    If Not modX13IO.IsNumberText(CStr(values("nreg"))) Then GoTo Done
    expected = CLng(modX13IO.NumberValue(CStr(values("nreg"))))
    If expected < 0 Then GoTo Done
    For Each key In values.Keys
        If CStr(key) = "nreg" Then
            remaining = expected
            inBlock = remaining > 0
        ElseIf CStr(key) = "nregderived" Or CStr(key) = "nonseasonaldiff" Or CStr(key) = "seasonaldiff" Or CStr(key) = "nmodel" Then
            inBlock = False
        ElseIf inBlock And InStr(CStr(key), "$") > 0 Then
            nums = Numbers(CStr(values(key)))
            If IsArray(nums) Then
                If UBound(nums) = 2 Then
                    result.Add CStr(key), nums
                    remaining = remaining - 1
                    inBlock = remaining > 0
                End If
            End If
        End If
    Next key
    complete = (result.Count = expected)
Done:
    Set RegressionCoefficients = result
End Function

' 目的  : UDGから自動検出された外れ値を取り出す。
' 出力  : Collection。要素は Array(種類, 年月(YYYY-MM), 係数, 標準誤差, t値, UDGのキー)
Public Function DetectedOutliers(ByVal values As Object) As Collection
    Dim result As New Collection, key As Variant, token As String, nums As Variant, re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(AO|LS|TC|SO)([0-9]{4})\.([A-Za-z]{3})$"
    For Each key In values.Keys
        If Left$(CStr(key), 12) = "AutoOutlier$" Then
            token = Mid$(CStr(key), 13)
            Set matches = re.Execute(token)
            If matches.Count > 0 Then
                nums = Numbers(CStr(values(key)))
                If IsArray(nums) Then
                    If UBound(nums) = 2 Then
                        result.Add Array(UCase$(CStr(matches(0).SubMatches(0))), _
                            CStr(matches(0).SubMatches(1)) & "-" & Format$(MonthFromAbbrev(CStr(matches(0).SubMatches(2))), "00"), _
                            nums(0), nums(1), nums(2), CStr(key))
                    End If
                End If
            End If
        End If
    Next key
    Set DetectedOutliers = result
End Function

Private Function MonthFromAbbrev(ByVal abbrev As String) As Long
    MonthFromAbbrev = (InStr("janfebmaraprmayjunjulaugsepoctnovdec", LCase$(abbrev)) + 2) \ 3
    If MonthFromAbbrev < 1 Or MonthFromAbbrev > 12 Then Err.Raise vbObjectError + 3100, "MonthFromAbbrev", "月の略号を解析できません：" & abbrev
End Function

' 目的  : 比較群に使う外れ値の署名（種類と月を小文字にして並べたもの）。自動・指定のどちらの表記でも同じ署名になる。
Public Function OutlierSignature(ByVal values As Object) As String
    Dim re As Object, key As Variant, matches As Object, tokens As Object, items() As String, i As Long, j As Long, swap As String
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(AO|LS|TC|SO)[0-9]{4}\.[A-Za-z]{3}$"
    Set tokens = CreateObject("Scripting.Dictionary")
    For Each key In values.Keys
        If InStr(CStr(key), "$") > 0 Then
            Set matches = re.Execute(CStr(key))
            If matches.Count > 0 Then tokens(LCase$(CStr(matches(0).Value))) = True
        End If
    Next key
    If tokens.Count = 0 Then Exit Function
    ReDim items(1 To tokens.Count)
    For Each key In tokens.Keys
        i = i + 1
        items(i) = CStr(key)
    Next key
    For i = 1 To UBound(items)
        For j = i + 1 To UBound(items)
            If items(j) < items(i) Then swap = items(i): items(i) = items(j): items(j) = swap
        Next j
    Next i
    For i = 1 To UBound(items)
        OutlierSignature = OutlierSignature & ";" & items(i)
    Next i
End Function

' 目的  : 保存したSPCから、結合検定が必要な変数群（weekday／leap／holiday）を決める。
'         現在のシートの設定ではなく、その計算に使ったSPCを根拠にする。
Public Function ExpectedCalendarGroups(ByVal specPath As String) As Object
    Dim result As Object, re As Object, text As String, matches As Object, token As Variant
    Set result = CreateObject("Scripting.Dictionary")
    text = modX13IO.ReadText(specPath, "UTF8")
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "variables\s*=\s*\(([^)]*)\)"
    Set matches = re.Execute(text)
    If matches.Count > 0 Then
        For Each token In Split(LCase$(Trim$(Replace(Replace(CStr(matches(0).SubMatches(0)), vbCr, " "), vbLf, " "))), " ")
            Select Case CStr(token)
                Case "tdnolpyear", "td1nolpyear": result("weekday") = True
                Case "lpyear": result("leap") = True
            End Select
        Next token
    End If
    re.Pattern = "usertype\s*=\s*holiday\b"
    If re.Test(text) Then result("holiday") = True
    Set ExpectedCalendarGroups = result
End Function

' 目的: UDGの係数キーがどの変数群に属するか（weekday／leap／holiday、該当なしは空文字）。
Public Function CalendarGroup(ByVal key As String) As String
    If InStr(1, key, "Trading Day", vbTextCompare) > 0 Then CalendarGroup = "weekday"
    If InStr(1, key, "User-defined Holiday", vbTextCompare) > 0 Then CalendarGroup = "holiday"
    If InStr(1, key, "Leap Year", vbTextCompare) > 0 Then CalendarGroup = "leap"
End Function

' 目的  : 最終推定の係数とRCM（回帰係数の共分散行列）から、変数群のWaldカイ二乗検定を計算する。
'         統計量 = b' V^-1 b（bは群の係数、Vはその共分散）、自由度 = 群の係数の数。
' 出力  : Dictionary（Available, Reason, Statistic, Df, P）。計算できない場合は Available=False と理由。
' 注意  : RCMの変数の並びが係数の並びと一致しない場合は計算しない（取り違えを防ぐため）。
Public Function WaldGroupTest(ByVal coefficients As Object, ByVal group As String, ByVal rcmPath As String) As Object
    Dim result As Object, lines As Variant, line As Variant, fields As Variant, keys As Variant
    Dim indices As New Collection, i As Long, j As Long, row As Long, n As Long, k As Long, statistic As Double
    Dim covariance() As Double, submatrix() As Double, inverse As Variant, beta() As Double, nums As Variant
    Dim reason As String
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Available", False
    result.Add "Reason", ""
    n = coefficients.Count
    If Not modMain.FileExists(rcmPath) Then result("Reason") = "RCMファイルがありません": GoTo Done
    If n = 0 Then result("Reason") = "回帰係数がありません": GoTo Done
    keys = coefficients.Keys
    ReDim covariance(1 To n, 1 To n)
    lines = Split(Replace(modX13IO.ReadText(rcmPath, "UTF8"), vbCr, ""), vbLf)
    For Each line In lines
        If Len(Trim$(CStr(line))) > 0 Then
            fields = Split(CStr(line), vbTab)
            If UBound(fields) = n Then
                If CStr(fields(0)) <> "variable" And Left$(CStr(fields(0)), 1) <> "-" And Len(CStr(fields(0))) > 0 Then
                    row = row + 1
                    If row > n Then reason = "RCMの行数が係数より多い": GoTo Unavailable
                    If CStr(fields(0)) <> Mid$(CStr(keys(row - 1)), InStrRev(CStr(keys(row - 1)), "$") + 1) Then _
                        reason = "RCMの変数の並びが係数と一致しない（" & CStr(fields(0)) & "）": GoTo Unavailable
                    For j = 1 To n
                        If Not modX13IO.IsNumberText(CStr(fields(j))) Then reason = "RCMに数値でない値がある": GoTo Unavailable
                        covariance(row, j) = modX13IO.NumberValue(CStr(fields(j)))
                    Next j
                End If
            End If
        End If
    Next line
    If row <> n Then reason = "RCMの行数が係数と一致しない（" & row & "／" & n & "）": GoTo Unavailable
    For i = 0 To n - 1
        If CalendarGroup(CStr(keys(i))) = group Then indices.Add i + 1
    Next i
    k = indices.Count
    If k = 0 Then reason = "この変数群の係数がない": GoTo Unavailable
    ReDim submatrix(1 To k, 1 To k)
    ReDim beta(1 To k)
    For i = 1 To k
        nums = coefficients(keys(CLng(indices(i)) - 1))
        beta(i) = CDbl(nums(0))
        For j = 1 To k
            submatrix(i, j) = covariance(CLng(indices(i)), CLng(indices(j)))
        Next j
    Next i
    If k = 1 Then
        If submatrix(1, 1) <= 0 Then reason = "係数の分散が0以下": GoTo Unavailable
        statistic = beta(1) * beta(1) / submatrix(1, 1)
    Else
        inverse = InverseMatrix(submatrix, k, reason)
        If Not IsArray(inverse) Then GoTo Unavailable
        For i = 1 To k
            For j = 1 To k
                statistic = statistic + beta(i) * CDbl(inverse(i, j)) * beta(j)
            Next j
        Next i
    End If
    If statistic < 0 Then reason = "Wald統計量が負": GoTo Unavailable
    result.Add "Statistic", statistic
    result.Add "Df", k
    result.Add "P", Application.WorksheetFunction.ChiSq_Dist_RT(statistic, k)
    result("Available") = True
    GoTo Done
Unavailable:
    result("Reason") = reason
Done:
    Set WaldGroupTest = result
End Function

' 目的  : k×k行列の逆行列をガウス・ジョルダン法（部分ピボット）で求める。
' 出力  : 逆行列（1始まりの2次元Double配列）。特異なら Empty を返し、reasonに理由を入れる。
Private Function InverseMatrix(ByRef matrix() As Double, ByVal k As Long, ByRef reason As String) As Variant
    Dim a() As Double, i As Long, j As Long, r As Long, pivot As Long, factor As Double, temp As Double, best As Double
    ReDim a(1 To k, 1 To 2 * k)
    For i = 1 To k
        For j = 1 To k
            a(i, j) = matrix(i, j)
        Next j
        a(i, k + i) = 1#
    Next i
    For j = 1 To k
        pivot = j
        best = Abs(a(j, j))
        For r = j + 1 To k
            If Abs(a(r, j)) > best Then best = Abs(a(r, j)): pivot = r
        Next r
        If best < 1E-300 Then reason = "共分散行列が特異": Exit Function
        If pivot <> j Then
            For i = 1 To 2 * k
                temp = a(j, i): a(j, i) = a(pivot, i): a(pivot, i) = temp
            Next i
        End If
        factor = a(j, j)
        For i = 1 To 2 * k
            a(j, i) = a(j, i) / factor
        Next i
        For r = 1 To k
            If r <> j Then
                factor = a(r, j)
                For i = 1 To 2 * k
                    a(r, i) = a(r, i) - factor * a(j, i)
                Next i
            End If
        Next r
    Next j
    Dim result() As Double
    ReDim result(1 To k, 1 To k)
    For i = 1 To k
        For j = 1 To k
            result(i, j) = a(i, k + j)
        Next j
    Next i
    InverseMatrix = result
End Function

' 目的  : カレンダー調整を指定しない計算ではX-13がD18を出さないため、中立値のD18を作る。
'         乗法型は1、加法型は0。VBAが作ったことを _d18_origin.txt に記録する。
' 入力  : basePath（拡張子なしのジョブファイル名）、specText、firstMonth、rows（実績＋予測の月数）
' エラー: カレンダー変数・祝日変数を指定したのにD18が無い場合は停止する（X-13の失敗を中立値で隠さない）。
Public Sub EnsureD18(ByVal basePath As String, ByVal specText As String, ByVal firstMonth As Long, ByVal rows As Long)
    Dim re As Object, values As Object, mode As String, neutral As Double, m As Long, body As String
    If modMain.FileExists(basePath & ".d18") Then Exit Sub
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "\b(tdnolpyear|td1nolpyear|lpyear|td|td1coef)\b|usertype\s*=\s*holiday"
    If re.Test(specText) Then Err.Raise vbObjectError + 3101, "EnsureD18", "カレンダー調整を指定した計算のD18がありません：" & basePath
    Set values = ReadUdgValues(basePath & ".udg")
    mode = ReadValue(values, "samode")
    If InStr(mode, "multiplicative") > 0 Then
        neutral = 1
    ElseIf InStr(mode, "additive") > 0 Then
        neutral = 0
    Else
        Err.Raise vbObjectError + 3102, "EnsureD18", "D18の中立値は乗法型・加法型だけに対応しています（samode=" & mode & "）"
    End If
    body = "date" & vbTab & "neutral_calendar_factor" & vbCrLf & "------" & vbTab & "------" & vbCrLf
    For m = firstMonth To firstMonth + rows - 1
        body = body & modX13IO.MonthCompact(m) & vbTab & modX13IO.NumberText(neutral) & vbCrLf
    Next m
    modX13IO.WriteNewText basePath & ".d18", body
    modX13IO.WriteNewText modMain.CombinePath(CreateObject("Scripting.FileSystemObject").GetParentFolderName(basePath), "_d18_origin.txt"), _
        "このD18はX-13の出力ではありません。カレンダー調整を指定しない計算のため、VBAが中立値（" & modX13IO.NumberText(neutral) & _
        "）で作成しました。" & vbCrLf
End Sub
