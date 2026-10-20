Attribute VB_Name = "modX13Interventions"
Option Explicit

' 役割      : 利用者が「設定_外れ値指定」に書いた外れ値・水準変化を検査し、SPCの回帰変数へ変換する。
'             自動検出（outlierブロック）とは別で、指定した月と種類をそのまま推定に入れる。
'             初回の自動検出で確定した外れ値を、再推定の候補へ固定する処理（FrozenOutliers）もここに置く。
' 主な入口  : SetupInterventionSheet / InterventionVariables / InterventionToken / FrozenOutliers / MergeRegressors
' 依存先    : modMain、modSheetNames、modX13IO、modX13Options
' 表の形    : series_id | setting_id | type | start_month | end_month | 説明
'             series_id・setting_id に「*」を書くと全系列・全設定IDに適用。月はYYYY-MM。
' 種類と書式: AO・LS・TC・SO … 開始月だけ（例 ao2020.apr）
'             TL・RAMP・QUADRATIC_INCREASE・QUADRATIC_DECREASE … 開始月と終了月（例 tl2020.apr-2020.jun、rp…、qi…、qd…）
'             （Census X-13 Reference Manual §7.13 regression の変数名）

' 目的  : 設定_外れ値指定 シートを作る（無い場合）。種類の一覧選択を付ける。
' エラー: 既存の見出しが違う場合は停止する。
Public Sub SetupInterventionSheet()
    Dim ws As Worksheet
    Set ws = modMain.EnsureWorksheet(modSheetNames.InterventionSheet())
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        ws.Cells(1, 1).Resize(1, 6).Value2 = InterventionHeaders()
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:E").NumberFormat = "@"
        ws.Cells(1, 8).Value2 = "記入例：industry_1000000000 | standard | LS | 2020-04 |  （LS・AO・TC・SOは終了月を空欄）"
        ws.Cells(2, 8).Value2 = "TL・RAMP・QUADRATIC_INCREASE・QUADRATIC_DECREASEは開始月と終了月を書きます。series_idの「*」は全系列、setting_idの「*」と空欄は全設定ID。"
        ws.Cells(3, 8).Value2 = "外れ値の扱い（outlier_mode）が「自動＋指定」「指定のみ」の系列だけに適用します。「自動」の系列では使いません。"
        ws.Columns("A").ColumnWidth = 28
        ws.Columns("B:E").ColumnWidth = 14
        ws.Columns("F").ColumnWidth = 30
    Else
        modX13IO.Headers ws, InterventionHeaders()
    End If
    modX13Options.AddChoice ws.Range("C2:C2000"), "AO,LS,TC,SO,TL,RAMP,QUADRATIC_INCREASE,QUADRATIC_DECREASE"
End Sub

Public Function InterventionHeaders() As Variant
    InterventionHeaders = Array("series_id", "setting_id", "type", "start_month", "end_month", "説明")
End Function

' 目的  : ある系列・設定IDに当てはまる指定を、SPCの変数名（空白区切り）で返す。
' 入力  : seriesId、settingId、firstMonth・lastMonth（実績期間の月番号）
' 出力  : 例 "ls2020.apr tl2021.jan-2021.mar"。該当なしは空文字。
' エラー: 種類・月の不正、期間外、重複、series_id空欄の行は、行番号を付けて停止する。
'         他の系列の行は読まないため、ある系列だけに書いた指定が他の系列へ波及しない。
Public Function InterventionVariables(ByVal seriesId As String, ByVal settingId As String, _
                                      ByVal firstMonth As Long, ByVal lastMonth As Long) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, matchSeries As String, matchSetting As String
    Dim token As String, seen As Object
    Set ws = modMain.WorksheetOrNothing(modSheetNames.InterventionSheet())
    If ws Is Nothing Then Exit Function
    modX13IO.Headers ws, InterventionHeaders()
    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For r = 2 To lastRow
        matchSeries = Trim$(CStr(ws.Cells(r, 1).Value2))
        matchSetting = Trim$(CStr(ws.Cells(r, 2).Value2))
        If Len(matchSeries) = 0 Then
            If Len(Trim$(CStr(ws.Cells(r, 3).Value2))) > 0 Or Len(Trim$(CStr(ws.Cells(r, 4).Value2))) > 0 Then _
                Err.Raise vbObjectError + 2950, "InterventionVariables", "設定_外れ値指定の" & r & "行目：series_idが空欄です。"
        ElseIf (matchSeries = "*" Or matchSeries = seriesId) And (matchSetting = "*" Or matchSetting = settingId Or Len(matchSetting) = 0) Then
            token = InterventionToken(Trim$(CStr(ws.Cells(r, 3).Value2)), Trim$(CStr(ws.Cells(r, 4).Value2)), _
                Trim$(CStr(ws.Cells(r, 5).Value2)), firstMonth, lastMonth, r)
            If seen.Exists(LCase$(token)) Then Err.Raise vbObjectError + 2951, "InterventionVariables", _
                "設定_外れ値指定の" & r & "行目：同じ指定が重複しています：" & token
            seen.Add LCase$(token), True
            InterventionVariables = Trim$(InterventionVariables & " " & token)
        End If
    Next r
End Function

' 目的  : 1行の指定をSPCの変数名にする。
' 入力  : effectType（AO等）、startText・endText（YYYY-MM）、実績期間、rowNumber（エラー表示用）
' エラー: 種類が不明、月の形式違い、期間外、終了月が開始月以前、1か月目からのLS・SOは停止する。
Public Function InterventionToken(ByVal effectType As String, ByVal startText As String, ByVal endText As String, _
                                  ByVal firstMonth As Long, ByVal lastMonth As Long, ByVal rowNumber As Long) As String
    Dim code As String, startMonth As Long, endMonth As Long, where As String
    where = "設定_外れ値指定の" & rowNumber & "行目："
    Select Case UCase$(effectType)
        Case "AO", "LS", "TC", "SO", "TL": code = LCase$(effectType)
        Case "RAMP", "RP": code = "rp"
        Case "QUADRATIC_INCREASE", "QI": code = "qi"
        Case "QUADRATIC_DECREASE", "QD": code = "qd"
        Case Else: Err.Raise vbObjectError + 2952, "InterventionToken", where & "種類は AO・LS・TC・SO・TL・RAMP・QUADRATIC_INCREASE・QUADRATIC_DECREASE です：" & effectType
    End Select
    If Len(startText) = 0 Then Err.Raise vbObjectError + 2953, "InterventionToken", where & "開始月が空欄です。"
    startMonth = modX13IO.MonthIndexFromText(startText)
    If startMonth < firstMonth Or startMonth > lastMonth Then Err.Raise vbObjectError + 2954, "InterventionToken", _
        where & "開始月が実績期間（" & modX13IO.MonthText(firstMonth) & "～" & modX13IO.MonthText(lastMonth) & "）の外です：" & startText
    Select Case code
        Case "ao", "ls", "tc", "so"
            If Len(endText) > 0 Then Err.Raise vbObjectError + 2955, "InterventionToken", where & UCase$(code) & "は開始月だけを書きます（終了月は空欄）。"
            If (code = "ls" Or code = "so") And startMonth = firstMonth Then Err.Raise vbObjectError + 2956, "InterventionToken", _
                where & UCase$(code) & "は実績期間の最初の月には置けません。"
            InterventionToken = code & modX13IO.X13MonthToken(startMonth)
        Case Else
            If Len(endText) = 0 Then Err.Raise vbObjectError + 2957, "InterventionToken", where & "終了月が必要です。"
            endMonth = modX13IO.MonthIndexFromText(endText)
            If endMonth <= startMonth Or endMonth > lastMonth Then Err.Raise vbObjectError + 2958, "InterventionToken", _
                where & "終了月は開始月より後で、実績期間内の月を書いてください：" & endText
            InterventionToken = code & modX13IO.X13MonthToken(startMonth) & "-" & modX13IO.X13MonthToken(endMonth)
    End Select
End Function

' 目的  : 初回計算のUDGから、自動検出で確定した外れ値（AutoOutlier$LS2019.Oct 等）を取り出す。
'         探索で現れたARIMA次数を再推定する候補へ、月と種類を固定して渡すために使う（係数は固定しない）。
' 入力  : values（modX13Results.ReadUdgValues の戻り値）
' 出力  : 例 "LS2019.Oct TC2020.May"。無ければ空文字。
' エラー: 想定外の表記は停止する（読み違えた外れ値で再推定しないため）。
Public Function FrozenOutliers(ByVal values As Object) As String
    Dim key As Variant, token As String, re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(AO|LS|TC|SO)[0-9]{4}\.[A-Za-z]{3}$"
    For Each key In values.Keys
        If Left$(CStr(key), 12) = "AutoOutlier$" Then
            token = Mid$(CStr(key), 13)
            If Not re.Test(token) Then Err.Raise vbObjectError + 2959, "FrozenOutliers", "自動検出された外れ値の表記を解析できません：" & CStr(key)
            FrozenOutliers = Trim$(FrozenOutliers & " " & token)
        End If
    Next key
End Function

' 目的: 回帰変数の文字列を、大文字小文字を区別せずに重複を除いて結合する。
Public Function MergeRegressors(ByVal first As String, ByVal second As String) As String
    Dim seen As Object, token As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    For Each token In Split(Trim$(first & " " & second), " ")
        If Len(CStr(token)) > 0 Then
            If Not seen.Exists(LCase$(CStr(token))) Then
                seen.Add LCase$(CStr(token)), True
                MergeRegressors = Trim$(MergeRegressors & " " & CStr(token))
            End If
        End If
    Next token
End Function
