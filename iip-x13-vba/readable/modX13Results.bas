Attribute VB_Name = "modX13Results"
Option Explicit

' 役割      : X-13の出力ファイルを読み、比較に使う値を取り出す。値の判断（順位・採否）はしない。
'             UDG（キー: 値の診断要約）… AIC・AICC・BIC・HQ、最終ARIMA次数、収束、回帰係数とt値、
'                                         予測誤差率（aape）、X-13の結合F検定、自動検出の外れ値
'             主出力（ASCII版 .out／HTML版 .html）… 最終ARIMA次数、変換、統計量、自動選択の探索履歴
'             RCM（回帰係数の共分散）    … 変数群のWald検定
'             あわせて、X-13の出力形式（ASCII版／HTML版）を判定し、主出力とエラー記録を検査する。
' 主な入口  : ReadUdgValues / ReadOutputEvidence / ParseSearchHistory / RegressionCoefficients
'             DetectedOutliers / WaldGroupTest / ExpectedCalendarGroups / CalendarGroup / EnsureD18
'             ReadValue / NumericValue / Numbers
'             出力形式：BatchLogFile / DetectReport / InspectReport / ReportText / ReportTextPath / HtmlToText
'                       EvidenceFromText / HtmlEvidenceProblem
' 依存先    : modX13IO、modMain
' 注意      : 自動選択の探索中に表示されるBIC2は、探索用の値として探索履歴にだけ記録する。
'             最終候補のAIC・BIC・AICC・HQの代わりには使わない（取得できない値は空のまま）。
' 出力形式  : 公式X-13にはASCII版（x13as_ascii.exe）とHTML版（x13as_html.exe）があり、計算は同じで出力の形式だけが違う。
'             D10等の保存表・UDG・RCMは両版で同じ形式。違うのは主出力・エラー記録・ログのファイル名と書式。
'               ASCII版 … <stem>.out（主出力）、<stem>.err（エラー記録）、バッチのログ jobs.log
'               HTML版  … <stem>.html（主出力）、<stem>_err.html（エラー記録）、バッチのログ jobs_log.html
'             （HTML版の名前は公式Build62の実出力 probe_err.html・probe_log.html と同版マニュアル2.6節による）
'             形式は実行ファイルの名前では決めない（名前を変えたexeでも動くように）。実行後にできたファイルで決め、
'             両方の形式のファイルがある・エラー記録が2つある等、どちらか決められない場合は停止する。
' HTMLの読み方: ブラウザやHTMLのスクリプトは使わず、文字列の走査でタグを除く（HtmlToText）。
'             得たテキストに、ASCII版と同じ解析規則（ParseSearchHistory・EvidenceFromText）を当てる。
'             HTMLから作ったテキストは、試行フォルダに REPORT_TEXT_FILE として残す。VBAが作った内部用の
'             テキストであり、X-13の出力ではない（人が確認する原本は <stem>.html）。

' 出力形式の名前（_trial.meta の REPORT_FORMAT に書く値）。
Public Const REPORT_ASCII As String = "ASCII"
Public Const REPORT_HTML As String = "HTML"
' HTML版の主出力から作った内部用テキストのファイル名（試行・最終フォルダに置く）。
Public Const REPORT_TEXT_FILE As String = "_report_text_from_html.txt"
' -m jobs で一括実行したときに、X-13がバッチフォルダへ書くログと索引。
Private Const BATCH_LOG_ASCII As String = "jobs.log"
Private Const BATCH_LOG_HTML As String = "jobs_log.html"
Private Const BATCH_INDEX_HTML As String = "jobs_mta.html"

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

' 目的  : 主出力（ASCII版 .out／HTML版 .html）とSPCから、最終ARIMA次数・実際の変換・統計量を取り出す。
' 出力  : EvidenceFromText と同じDictionary。主出力が無ければ空のDictionary。
Public Function ReadOutputEvidence(ByVal folder As String, ByVal stem As String) As Object
    Dim result As Object
    If Not HasReport(folder, stem) Then
        Set result = CreateObject("Scripting.Dictionary")
        result.CompareMode = vbTextCompare
        Set ReadOutputEvidence = result
        Exit Function
    End If
    Set ReadOutputEvidence = EvidenceFromText(ReportText(folder, stem), _
        modX13IO.ReadText(modMain.CombinePath(folder, stem & ".spc"), "UTF8"))
End Function

' 目的  : 主出力のテキスト（ASCII版はそのまま、HTML版は HtmlToText の結果）とSPCから値を取り出す。
' 出力  : Dictionary（X13_ACTUAL_MODEL, X13_ACTUAL_TRANSFORM, X13_ACTUAL_AIC, X13_ACTUAL_AICC, X13_ACTUAL_BIC, X13_ACTUAL_HQ,
'         X13_REPORT_AAPE_1～3・X13_REPORT_AAPE_0（標本内予測の平均絶対誤差率。Last year～Last-2 year、Last three years））
'         見つからない項目は入れない。同じ項目が何度も出る場合は最後（最終推定）の値。
'         transform=autoの場合は、主出力の「prefers log/no transformation」から決める。
' 注意  : 予測誤差の値はHTML版の読取りの照合（HtmlEvidenceProblem）だけに使う。比較表の値はUDGのaape。
Public Function EvidenceFromText(ByVal output As String, ByVal spec As String) As Object
    Dim lines As Variant, line As Variant, result As Object, re As Object, aape As Variant, k As Long
    Dim matches As Object, model As String, transform As String, name As Variant, pattern As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
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
    ' 標本内予測の平均絶対誤差率。HTML版は1項目が1行、ASCII版は1行に並ぶことがあるため、行に限らず探す。
    re.Global = True
    aape = Array("Last year", "Last-1 year", "Last-2 year", "Last three years")
    For k = 0 To 3
        re.Pattern = aape(k) & "\s*:\s*([+-]?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?)"
        Set matches = re.Execute(output)
        If matches.Count > 0 Then result("X13_REPORT_AAPE_" & IIf(k = 3, 0, k + 1)) = CStr(matches(matches.Count - 1).SubMatches(0))
    Next k
    Set EvidenceFromText = result
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

' ================= 出力形式（ASCII版／HTML版）の判定と検査 =================

' 目的  : バッチフォルダにX-13が書いたログのファイル名を返す（ASCII版 jobs.log、HTML版 jobs_log.html）。
'         このログの形式を、そのバッチで動いたX-13の出力形式とみなす（実行ファイルの名前は見ない）。
' エラー: どちらも無い、または両方ある場合は停止する（どちらの版で動いたか決められないため）。
Public Function BatchLogFile(ByVal batchFolder As String) As String
    Dim hasAscii As Boolean, hasHtml As Boolean
    hasAscii = modMain.FileExists(modMain.CombinePath(batchFolder, BATCH_LOG_ASCII))
    hasHtml = modMain.FileExists(modMain.CombinePath(batchFolder, BATCH_LOG_HTML))
    If hasAscii And hasHtml Then Err.Raise vbObjectError + 3110, "BatchLogFile", _
        "X-13のバッチのログが2つの形式で見つかりました（" & BATCH_LOG_ASCII & "と" & BATCH_LOG_HTML & "）。どちらの版の結果か決められません：" & batchFolder
    If hasAscii Then
        BatchLogFile = BATCH_LOG_ASCII
    ElseIf hasHtml Then
        BatchLogFile = BATCH_LOG_HTML
    Else
        Err.Raise vbObjectError + 3111, "BatchLogFile", "X-13のバッチのログ（" & BATCH_LOG_ASCII & "または" & BATCH_LOG_HTML & "）がありません：" & batchFolder
    End If
End Function

' 目的: バッチのログのファイル名から出力形式（REPORT_ASCII／REPORT_HTML）を返す。
Public Function BatchLogFormat(ByVal logName As String) As String
    Select Case LCase$(logName)
        Case BATCH_LOG_ASCII: BatchLogFormat = REPORT_ASCII
        Case BATCH_LOG_HTML: BatchLogFormat = REPORT_HTML
        Case Else: Err.Raise vbObjectError + 3112, "BatchLogFormat", "X-13のバッチのログの名前が想定外です：" & logName
    End Select
End Function

' 目的: HTML版がバッチフォルダに書くメタファイルの索引（jobs_mta.html）の名前。無ければ空文字。
'       マニュアル2.5.1節の記載による。保存するだけで、検査には使わない。
Public Function BatchIndexFile(ByVal batchFolder As String) As String
    If modMain.FileExists(modMain.CombinePath(batchFolder, BATCH_INDEX_HTML)) Then BatchIndexFile = BATCH_INDEX_HTML
End Function

' HTML版のエラー記録の名前の候補。公式Build62の実物は <stem>_err.html。
' マニュアルの文面からは下線か点か読み取れないため、他の版が書き得る <stem>.err.html も候補にする（両方あれば停止）。
Private Function HtmlErrorSuffixes() As Variant
    HtmlErrorSuffixes = Array("_err.html", ".err.html")
End Function

' 目的: フォルダに主出力（.out または .html）があるか。
Public Function HasReport(ByVal folder As String, ByVal stem As String) As Boolean
    HasReport = modMain.FileExists(modMain.CombinePath(folder, stem & ".out")) Or _
                modMain.FileExists(modMain.CombinePath(folder, stem & ".html"))
End Function

' 目的  : フォルダにあるファイルから、X-13の出力形式と主出力・エラー記録のファイル名を決める。中身は読まない。
' 出力  : Dictionary（Format＝REPORT_ASCII／REPORT_HTML、Main＝主出力のファイル名、Error＝エラー記録のファイル名）
' エラー: 次の場合は停止する。
'         ASCII版（.out・.err）とHTML版（.html・_err.html）のファイルが同じフォルダにある（どちらの結果か決められない）
'         主出力もエラー記録も無い、片方だけある、HTML版のエラー記録の候補が2つある
Public Function DetectReport(ByVal folder As String, ByVal stem As String) As Object
    Dim base As String, hasOut As Boolean, hasErr As Boolean, hasHtml As Boolean, suffix As Variant
    Dim htmlErrors As New Collection, result As Object
    base = modMain.CombinePath(folder, stem)
    hasOut = modMain.FileExists(base & ".out")
    hasErr = modMain.FileExists(base & ".err")
    hasHtml = modMain.FileExists(base & ".html")
    For Each suffix In HtmlErrorSuffixes()
        If modMain.FileExists(base & CStr(suffix)) Then htmlErrors.Add stem & CStr(suffix)
    Next suffix
    If (hasOut Or hasErr) And (hasHtml Or htmlErrors.Count > 0) Then Err.Raise vbObjectError + 3113, "DetectReport", _
        "ASCII版（.out・.err）とHTML版（.html・_err.html）の出力が同じフォルダにあります。どちらの結果か決められません：" & base
    Set result = CreateObject("Scripting.Dictionary")
    If hasOut Or hasErr Then
        If Not hasErr Then Err.Raise vbObjectError + 3114, "DetectReport", "X-13のエラー記録（" & stem & ".err）がありません"
        If Not hasOut Then Err.Raise vbObjectError + 3115, "DetectReport", "X-13の主出力（" & stem & ".out）がありません。エラー記録（" & stem & ".err）を確認してください"
        result.Add "Format", REPORT_ASCII
        result.Add "Main", stem & ".out"
        result.Add "Error", stem & ".err"
    ElseIf hasHtml Or htmlErrors.Count > 0 Then
        If htmlErrors.Count > 1 Then Err.Raise vbObjectError + 3116, "DetectReport", _
            "HTML版のエラー記録が2つあります（" & htmlErrors(1) & "と" & htmlErrors(2) & "）。どちらが正しいか決められません"
        If htmlErrors.Count = 0 Then Err.Raise vbObjectError + 3117, "DetectReport", "X-13のエラー記録（" & stem & "_err.html）がありません"
        If Not hasHtml Then Err.Raise vbObjectError + 3118, "DetectReport", "X-13の主出力（" & stem & ".html）がありません。エラー記録（" & htmlErrors(1) & "）を確認してください"
        result.Add "Format", REPORT_HTML
        result.Add "Main", stem & ".html"
        result.Add "Error", htmlErrors(1)
    Else
        Err.Raise vbObjectError + 3119, "DetectReport", "X-13の主出力（" & stem & ".out または " & stem & ".html）がありません（入力エラーで停止した可能性）"
    End If
    Set DetectReport = result
End Function

' 目的  : X-13の主出力とエラー記録を検査し、形式とファイル名（と必要なら解析用テキスト）を返す。
' 入力  : folder、stem、expectedFormat（そのバッチ・試行に記録された形式。空欄なら照合しない）、
'         withText（Trueなら Text に解析用テキストを入れる。Falseなら空文字）
' 出力  : Dictionary（Format, Main, Error, Text）。Text は、ASCII版は主出力そのもの、HTML版は HtmlToText の結果。
' 検査  : 1) 形式の判定（DetectReport）と、記録された形式との一致
'         2) エラー記録：ASCII版は大文字の「ERROR」を含まない（r11初版と同じ判定）。
'            HTML版は </html> で終わるHTMLで、画面に表示される本文に大文字の「ERROR」を含まない。
'            head（title・meta）・タグの属性・リンク先・コメントは見ないため、定型の見出し
'            「Error messages generated …」やリンク名「Error Messages」では不合格にしない。
'            警告（WARNING:）は不合格にしない。X-13は入力エラーでも終了コード0を返すため、この記録で判定する。
'         3) 主出力：空でない。HTML版は </html> で終わる（途中で切れたファイルを拒否）。
'            ASCII版は中身がHTMLでない（名前を変えたファイルの取り違えを拒否）。
' エラー: 不合格なら最初に見つかった理由で停止する。
Public Function InspectReport(ByVal folder As String, ByVal stem As String, ByVal expectedFormat As String, _
                              ByVal withText As Boolean) As Object
    Dim result As Object, mainPath As String, errorPath As String, raw As String, text As String
    Set result = DetectReport(folder, stem)
    If Len(expectedFormat) > 0 And result("Format") <> expectedFormat Then Err.Raise vbObjectError + 3120, "InspectReport", _
        "X-13の出力が" & result("Format") & "版の形式ですが、記録は" & expectedFormat & "版です（" & result("Main") & "）"
    mainPath = modMain.CombinePath(folder, CStr(result("Main")))
    errorPath = modMain.CombinePath(folder, CStr(result("Error")))
    If result("Format") = REPORT_ASCII Then
        raw = modX13IO.ReadText(errorPath, "UTF8")
        If InStr(1, raw, "ERROR", vbBinaryCompare) > 0 Then Err.Raise vbObjectError + 3121, "InspectReport", _
            "X-13がERRORを記録しました（" & result("Error") & "）：" & ErrorExcerpt(raw)
        If FileLen(mainPath) = 0 Then Err.Raise vbObjectError + 3122, "InspectReport", "X-13の主出力（" & result("Main") & "）が空です"
        raw = modX13IO.ReadText(mainPath, "UTF8")
        If LooksLikeHtml(raw) Then Err.Raise vbObjectError + 3123, "InspectReport", _
            "主出力（" & result("Main") & "）の中身がHTMLです。ASCII版の出力ではありません"
        If withText Then text = raw
    Else
        raw = modX13IO.ReadText(errorPath, "LATIN1")
        RequireHtmlDocument raw, "X-13のエラー記録（" & result("Error") & "）"
        text = HtmlToText(raw)
        If InStr(1, text, "ERROR", vbBinaryCompare) > 0 Then Err.Raise vbObjectError + 3121, "InspectReport", _
            "X-13がERRORを記録しました（" & result("Error") & "）：" & ErrorExcerpt(text)
        text = ""
        If FileLen(mainPath) = 0 Then Err.Raise vbObjectError + 3122, "InspectReport", "X-13の主出力（" & result("Main") & "）が空です"
        raw = modX13IO.ReadText(mainPath, "LATIN1")
        RequireHtmlDocument raw, "X-13の主出力（" & result("Main") & "）"
        If withText Then text = HtmlToText(raw)
    End If
    result.Add "Text", text
    Set InspectReport = result
End Function

' エラー記録のうち、最初の「ERROR」から200文字までを1行にまとめる（失敗の理由として試行_ジョブに残すため）。
Private Function ErrorExcerpt(ByVal text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"
    ErrorExcerpt = Trim$(re.Replace(Mid$(text, InStr(1, text, "ERROR", vbBinaryCompare), 400), " "))
    If Len(ErrorExcerpt) > 200 Then ErrorExcerpt = Left$(ErrorExcerpt, 200) & "…"
End Function

' 目的  : 主出力を解析用テキストで返す。ASCII版は .out そのもの。
'         HTML版は保存済みの内部用テキスト（REPORT_TEXT_FILE）を読み、無ければ <stem>.html から作る。
Public Function ReportText(ByVal folder As String, ByVal stem As String) As String
    Dim report As Object
    Set report = DetectReport(folder, stem)
    If report("Format") = REPORT_ASCII Then
        ReportText = modX13IO.ReadText(modMain.CombinePath(folder, CStr(report("Main"))), "UTF8")
    ElseIf modMain.FileExists(modMain.CombinePath(folder, REPORT_TEXT_FILE)) Then
        ReportText = modX13IO.ReadText(modMain.CombinePath(folder, REPORT_TEXT_FILE), "UTF8")
    Else
        ReportText = HtmlToText(modX13IO.ReadText(modMain.CombinePath(folder, CStr(report("Main"))), "LATIN1"))
    End If
End Function

' 目的: 探索履歴の「元の出力の行」が指すファイルのパス。ASCII版は <stem>.out、HTML版は内部用テキスト。
Public Function ReportTextPath(ByVal folder As String, ByVal stem As String) As String
    Dim report As Object
    Set report = DetectReport(folder, stem)
    If report("Format") = REPORT_ASCII Then
        ReportTextPath = modMain.CombinePath(folder, CStr(report("Main")))
    Else
        ReportTextPath = modMain.CombinePath(folder, REPORT_TEXT_FILE)
    End If
End Function

' 目的  : HTML版の主出力から読んだ値を、同じ計算のUDGと照合する（HTMLの読み違い・取り違えを見つけるため）。
' 入力  : text（HtmlToTextの結果）、specText（SPC）、values（UDG）、automdl（自動選択の候補ならTrue）
' 出力  : 問題が無ければ空文字。あれば最初に見つかった理由。
' 規則  : 最終ARIMA次数は必須で、UDGに arimamdl があれば一致すること。
'         AIC・AICC・BIC・HQ は、UDGに数値があれば主出力にも必須で、差が0.01以内（主出力は小数4桁で印字）。
'         予測誤差（Last year 等）は、両方にあれば差が0.006以内（主出力は小数2桁で印字）。
'         自動選択の候補は、探索履歴に推定できた次数が1件以上あること（空の探索履歴を正常としない）。
' 注意  : 照合に使うだけで、比較表の統計量はUDGの値のまま。探索中のBIC2は照合にも比較にも使わない。
Public Function HtmlEvidenceProblem(ByVal text As String, ByVal specText As String, ByVal values As Object, _
                                    ByVal automdl As Boolean) As String
    Dim evidence As Object, names As Variant, keys As Variant, i As Long, k As Long, item As Variant, found As Boolean
    Set evidence = EvidenceFromText(text, specText)
    If Not evidence.Exists("X13_ACTUAL_MODEL") Then HtmlEvidenceProblem = "主出力（HTML）から最終ARIMA次数を読めません": Exit Function
    If Len(ReadValue(values, "arimamdl")) > 0 Then
        If CompactModel(ReadValue(values, "arimamdl")) <> CompactModel(CStr(evidence("X13_ACTUAL_MODEL"))) Then
            HtmlEvidenceProblem = "主出力（HTML）の最終ARIMA次数 " & evidence("X13_ACTUAL_MODEL") & " がUDGの " & ReadValue(values, "arimamdl") & " と違います"
            Exit Function
        End If
    End If
    names = Array("AIC", "AICC", "BIC", "HQ")
    keys = Array("aic", "aicc", "bic", "hq")
    For i = 0 To 3
        If Not IsEmpty(NumericValue(values, CStr(keys(i)))) Then
            If Not evidence.Exists("X13_ACTUAL_" & names(i)) Then HtmlEvidenceProblem = "主出力（HTML）から" & names(i) & "を読めません": Exit Function
            If Abs(modX13IO.NumberValue(CStr(evidence("X13_ACTUAL_" & names(i)))) - CDbl(NumericValue(values, CStr(keys(i))))) > 0.01 Then
                HtmlEvidenceProblem = "主出力（HTML）の" & names(i) & "（" & evidence("X13_ACTUAL_" & names(i)) & "）がUDGの値と違います"
                Exit Function
            End If
        End If
    Next i
    For k = 0 To 3
        If Not IsEmpty(NumericValue(values, "aape." & k)) And evidence.Exists("X13_REPORT_AAPE_" & k) Then
            If Abs(modX13IO.NumberValue(CStr(evidence("X13_REPORT_AAPE_" & k))) - CDbl(NumericValue(values, "aape." & k))) > 0.006 Then
                HtmlEvidenceProblem = "主出力（HTML）の予測誤差（" & evidence("X13_REPORT_AAPE_" & k) & "）がUDGの aape." & k & " と違います"
                Exit Function
            End If
        End If
    Next k
    If automdl Then
        For Each item In ParseSearchHistory(text)
            If item("State") = "推定済" Then found = True
        Next item
        If Not found Then HtmlEvidenceProblem = "自動選択の探索履歴を主出力（HTML）から読めません（推定した次数が0件）"
    End If
End Function

' ARIMA次数の空白を除いて比べるための形（例 "(0 1 1)(0 1 1)" → "(011)(011)"）。
Private Function CompactModel(ByVal model As String) As String
    CompactModel = Replace(Replace(Replace(model, " ", ""), vbTab, ""), ChrW$(160), "")
End Function

' ================= HTMLからテキストへ =================

' 目的  : X-13 HTML版の出力から、画面に表示される本文だけをテキストにする。ブラウザやスクリプトは使わない。
' 入力  : html（ファイルの中身。改行はCRLF・LFのどちらでもよい）
' 出力  : テキスト（改行はLF）。規則は次のとおり。
'         - head・style・script・title の中身と、コメント（<!-- … -->）は捨てる
'         - 段落・改行・見出し・表の行・表・リスト等の区切り（br・p・div・tr・table・h1～h6・li 等）は改行にする
'         - 表のセル（td・th）の終わりはタブにする（「AIC<タブ>946.6621」のように項目と値を同じ行に残すため）
'         - それ以外のタグ（a・abbr・strong 等）は除き、中の文字だけを残す。タグの属性は読まない
'           （引用符の中の「>」はタグの終わりとみなさない）
'         - 文字参照（&nbsp; &amp; &lt; &gt; &quot; &apos; &#十進; &#x十六進;）を文字に戻す。&nbsp;（U+00A0）は空白にする。
'           知らない名前の参照はそのまま残す
'         - 「<」の後が英字・/・!・? でなければ、タグではなく文字の「<」とする
' エラー: タグ・コメント・head 等が閉じていない（ファイルが途中で切れている）場合は停止する。
' 補足  : 出力は入力より長くならない（タグは1文字以下、文字参照は1～2文字になる）ため、先に入力と同じ長さの
'         領域を取り、Mid$で書き込む（文字列の連結を繰り返すと大きなファイルで遅くなるため）。
Public Function HtmlToText(ByVal html As String) As String
    Dim buffer As String, size As Long, pos As Long, lt As Long, gt As Long, n As Long, closeAt As Long
    Dim name As String, closing As Boolean, nextChar As String
    html = Replace(Replace(html, vbCrLf, vbLf), vbCr, vbLf)
    n = Len(html)
    buffer = Space$(n + 1)
    pos = 1
    Do While pos <= n
        lt = InStr(pos, html, "<")
        If lt = 0 Then lt = n + 1
        If lt > pos Then AppendBuffer buffer, size, DecodeEntities(Mid$(html, pos, lt - pos))
        If lt > n Then Exit Do
        nextChar = Mid$(html, lt + 1, 1)
        If Mid$(html, lt, 4) = "<!--" Then
            closeAt = InStr(lt + 4, html, "-->")
            If closeAt = 0 Then Err.Raise vbObjectError + 3130, "HtmlToText", "HTMLのコメントが閉じていません（ファイルが途中で切れている可能性）"
            pos = closeAt + 3
        ElseIf Not (nextChar Like "[A-Za-z/!?]") Then
            AppendBuffer buffer, size, "<"
            pos = lt + 1
        Else
            gt = TagEnd(html, lt)
            name = TagName(Mid$(html, lt + 1, gt - lt - 1), closing)
            pos = gt + 1
            Select Case name
                Case "head", "style", "script", "title"
                    If Not closing Then
                        closeAt = InStr(pos, html, "</" & name, vbTextCompare)
                        If closeAt = 0 Then Err.Raise vbObjectError + 3131, "HtmlToText", "HTMLの" & name & "が閉じていません（ファイルが途中で切れている可能性）"
                        pos = TagEnd(html, closeAt) + 1
                        AppendBuffer buffer, size, vbLf
                    End If
                Case "br", "p", "div", "tr", "table", "caption", "h1", "h2", "h3", "h4", "h5", "h6", "li", "ul", "ol", _
                     "pre", "hr", "dl", "dt", "dd", "tbody", "thead", "tfoot", "body", "html", "blockquote"
                    AppendBuffer buffer, size, vbLf
                Case "td", "th"
                    If closing Then AppendBuffer buffer, size, vbTab
            End Select
        End If
    Loop
    HtmlToText = Replace(Left$(buffer, size), ChrW$(160), " ")
End Function

' 文字列の領域（buffer）の size 文字目の後ろへ text を書く。足りなければ領域を広げる。
Private Sub AppendBuffer(ByRef buffer As String, ByRef size As Long, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If size + Len(text) > Len(buffer) Then buffer = buffer & Space$(Len(buffer) + Len(text))
    Mid$(buffer, size + 1, Len(text)) = text
    size = size + Len(text)
End Sub

' 目的  : lt の位置の「<」に対応するタグの終わり「>」の位置を返す。引用符（" または '）の中の「>」は飛ばす。
' 補足  : 探す範囲を次の「>」までに限り、その間に引用符があれば閉じる引用符の後から探し直す
'         （ファイル全体を1文字ずつ見るより速い）。
' エラー: 「>」または閉じる引用符が無ければ停止する。
Private Function TagEnd(ByVal html As String, ByVal lt As Long) As Long
    Dim k As Long, gt As Long, segment As String, dq As Long, sq As Long, q As Long, closeQuote As Long
    k = lt + 1
    Do
        gt = InStr(k, html, ">")
        If gt = 0 Then Err.Raise vbObjectError + 3132, "TagEnd", "HTMLのタグが閉じていません（ファイルが途中で切れている可能性）"
        segment = Mid$(html, k, gt - k)
        dq = InStr(segment, """")
        sq = InStr(segment, "'")
        If dq = 0 And sq = 0 Then
            TagEnd = gt
            Exit Function
        End If
        If dq = 0 Or (sq > 0 And sq < dq) Then q = sq Else q = dq
        q = k + q - 1
        closeQuote = InStr(q + 1, html, Mid$(html, q, 1))
        If closeQuote = 0 Then Err.Raise vbObjectError + 3133, "TagEnd", "HTMLのタグの引用符が閉じていません（ファイルが途中で切れている可能性）"
        k = closeQuote + 1
    Loop
End Function

' 目的: タグの中身（「<」「>」を除いた部分）から、小文字のタグ名と終了タグかどうかを返す。例 "/TD" → "td"、closing=True。
Private Function TagName(ByVal inner As String, ByRef closing As Boolean) As String
    Dim i As Long
    inner = LTrim$(Replace(Replace(inner, vbLf, " "), vbTab, " "))
    closing = (Left$(inner, 1) = "/")
    If closing Then inner = LTrim$(Mid$(inner, 2))
    i = 1
    If Left$(inner, 1) = "!" Or Left$(inner, 1) = "?" Then i = 2
    Do While i <= Len(inner)
        If Not (Mid$(inner, i, 1) Like "[A-Za-z0-9]") Then Exit Do
        i = i + 1
    Loop
    TagName = LCase$(Left$(inner, i - 1))
End Function

' 目的: 文字参照を文字に戻す。形の合わない「&」はそのまま残す。
Private Function DecodeEntities(ByVal text As String) As String
    Dim pos As Long, amp As Long, semi As Long, replaced As String, ok As Boolean, result As String
    If InStr(text, "&") = 0 Then
        DecodeEntities = text
        Exit Function
    End If
    pos = 1
    Do
        amp = InStr(pos, text, "&")
        If amp = 0 Then
            result = result & Mid$(text, pos)
            Exit Do
        End If
        result = result & Mid$(text, pos, amp - pos)
        semi = InStr(amp + 1, text, ";")
        ok = False
        If semi > 0 And semi - amp <= 11 Then replaced = EntityText(Mid$(text, amp + 1, semi - amp - 1), ok)
        If ok Then
            result = result & replaced
            pos = semi + 1
        Else
            result = result & "&"
            pos = amp + 1
        End If
    Loop
    DecodeEntities = result
End Function

' 目的: 文字参照の名前（「&」と「;」の間）を文字にする。知らない名前・範囲外の番号は ok=False。
Private Function EntityText(ByVal name As String, ByRef ok As Boolean) As String
    Dim re As Object, code As Double, i As Long
    ok = True
    Select Case LCase$(name)
        Case "nbsp": EntityText = " "
        Case "amp": EntityText = "&"
        Case "lt": EntityText = "<"
        Case "gt": EntityText = ">"
        Case "quot": EntityText = """"
        Case "apos": EntityText = "'"
        Case Else
            ok = False
            Set re = CreateObject("VBScript.RegExp")
            re.Pattern = "^#([0-9]{1,7}|[xX][0-9A-Fa-f]{1,6})$"
            If Not re.Test(name) Then Exit Function
            If LCase$(Mid$(name, 2, 1)) = "x" Then
                For i = 3 To Len(name)
                    code = code * 16 + InStr("0123456789abcdef", LCase$(Mid$(name, i, 1))) - 1
                Next i
            Else
                code = CDbl(Mid$(name, 2))
            End If
            If code <= 0 Or code > 1114111 Then Exit Function
            If code <= 65535 Then
                EntityText = ChrW$(CLng(code))
            Else
                code = code - 65536
                EntityText = ChrW$(CLng(55296 + Int(code / 1024))) & ChrW$(CLng(56320 + (code - Int(code / 1024) * 1024)))
            End If
            ok = True
    End Select
End Function

' 目的  : HTMLとして完結しているか確かめる（<html があり、最後が </html> で終わる）。
' エラー: 空、HTMLでない、</html> が無い・その後に文字がある（途中で切れた・壊れた）場合は停止する。
Private Sub RequireHtmlDocument(ByVal raw As String, ByVal label As String)
    Dim re As Object, closeAt As Long
    If Len(BlankStripped(raw)) = 0 Then Err.Raise vbObjectError + 3134, "RequireHtmlDocument", label & "が空です"
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "<html\b"
    If Not re.Test(Left$(raw, 4096)) Then Err.Raise vbObjectError + 3135, "RequireHtmlDocument", label & "がHTMLではありません"
    closeAt = InStrRev(raw, "</html>", -1, vbTextCompare)
    If closeAt = 0 Then Err.Raise vbObjectError + 3136, "RequireHtmlDocument", label & "の終わり（</html>）がありません。途中で切れている可能性があります"
    If Len(BlankStripped(Mid$(raw, closeAt + 7))) > 0 Then Err.Raise vbObjectError + 3136, "RequireHtmlDocument", _
        label & "の終わり（</html>）の後に文字があります。壊れている可能性があります"
End Sub

' 空白・タブ・改行を除いた文字列（空かどうかの判定に使う）。
Private Function BlankStripped(ByVal text As String) As String
    BlankStripped = Replace(Replace(Replace(Replace(text, vbCr, ""), vbLf, ""), vbTab, ""), " ", "")
End Function

' 目的: 文字列がHTMLの文書として始まっているか（先頭の空白の後が <!DOCTYPE html または <html）。
Private Function LooksLikeHtml(ByVal raw As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^\s*(<!DOCTYPE\s+html|<html\b)"
    LooksLikeHtml = re.Test(Left$(raw, 4096))
End Function
