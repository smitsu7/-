Attribute VB_Name = "modX13Calendar"
Option Explicit

' 役割      : 系列別設定（cfg）から、比較するカレンダー変数の組合せ（候補）を作る。
'             使う変数は tdnolpyear（7曜日）・td1nolpyear（平日と土日の2区分）・lpyear（うるう年）・
'             hol（内閣府CSVから作るIIP方式の祝日変数）の4つ。
'             tdnolpyear と td1nolpyear は代わりになる候補で、同じ候補へ同時には入れない。
' 主な入口  : CandidateCalendars / RegressionVariables / UsesHoliday / CalendarLabel
' 依存先    : modSeriesSettings（cfgの読取り）
' 決め方    : X13_CALENDAR_MODE（系列別列 calendar_mode）
'             なし        … カレンダー変数を使わない1候補
'             指定のみ    … X13_FORCED_CALENDAR の変数で1候補。空欄なら td_mode・leap_year_mode・holiday_adjust から作る
'             自動        … X13_CALENDAR_POOL の各変数の有無を組み合わせた全候補（曜日は「なし／7曜日／2曜日」から1つ）
'             自動＋指定  … 自動の各候補に、必ず使う変数（X13_FORCED_CALENDAR）を加えた候補
' 例        : 自動、候補変数が tdnolpyear td1nolpyear lpyear hol なら 3×2×2＝12候補
'             （none／lpyear／hol／lpyear hol／tdnolpyear …／td1nolpyear lpyear hol）

' 目的  : 比較するカレンダーの組合せを返す。
' 入力  : cfg（modSeriesSettings.ResolveAll の戻り値）
' 出力  : Collection。各要素は変数を空白で区切った文字列。変数なしは "none"。
' エラー: 不明な変数、曜日2種類の同時指定、不明な決め方は停止する。
Public Function CandidateCalendars(ByVal cfg As Object) As Collection
    Dim result As New Collection, seen As Object, mode As String, pool As Object, forced As String
    Dim weekdayChoices As New Collection, td As Variant, leap As Long, hol As Long, maxLeap As Long, maxHol As Long
    Dim calendar As String
    mode = Trim$(modSeriesSettings.SettingText(cfg, "X13_CALENDAR_MODE"))
    Set seen = CreateObject("Scripting.Dictionary")
    Select Case mode
        Case "なし"
            result.Add "none"
        Case "指定のみ"
            forced = ForcedCalendar(cfg)
            If Len(forced) = 0 Then forced = "none"
            result.Add forced
        Case "自動", "自動＋指定"
            Set pool = ParseTokens(modSeriesSettings.SettingText(cfg, "X13_CALENDAR_POOL"), False)
            forced = ""
            If mode = "自動＋指定" Then forced = ForcedCalendar(cfg)
            weekdayChoices.Add ""
            If pool.Exists("tdnolpyear") Then weekdayChoices.Add "tdnolpyear"
            If pool.Exists("td1nolpyear") Then weekdayChoices.Add "td1nolpyear"
            If pool.Exists("lpyear") Then maxLeap = 1
            If pool.Exists("hol") Then maxHol = 1
            For Each td In weekdayChoices
                For leap = 0 To maxLeap
                    For hol = 0 To maxHol
                        calendar = CStr(td)
                        If leap = 1 Then calendar = Trim$(calendar & " lpyear")
                        If hol = 1 Then calendar = Trim$(calendar & " hol")
                        calendar = MergeTokens(calendar, forced)
                        ' 必ず使う変数に曜日がある場合、別の曜日を足した組合せは作らない。
                        If IsCompatible(calendar) Then
                            If Len(calendar) = 0 Then calendar = "none"
                            If Not seen.Exists(calendar) Then
                                seen.Add calendar, True
                                result.Add calendar
                            End If
                        End If
                    Next hol
                Next leap
            Next td
        Case Else
            Err.Raise vbObjectError + 2900, "CandidateCalendars", "カレンダーの決め方は 自動・自動＋指定・指定のみ・なし のいずれかです：" & mode
    End Select
    Set CandidateCalendars = result
End Function

' 目的  : 「指定のみ」「自動＋指定」で必ず使う変数を返す。X13_FORCED_CALENDARが空なら3つの切替から作る。
' エラー: 曜日2種類の同時指定は停止する。
Private Function ForcedCalendar(ByVal cfg As Object) As String
    Dim text As String, tokens As Object, key As Variant
    text = Trim$(modSeriesSettings.SettingText(cfg, "X13_FORCED_CALENDAR"))
    If Len(text) = 0 Then
        Select Case Trim$(modSeriesSettings.SettingText(cfg, "X13_TD_MODE"))
            Case "OFF", ""
            Case "2曜日": text = "td1nolpyear"
            Case "7曜日": text = "tdnolpyear"
            Case Else: Err.Raise vbObjectError + 2901, "ForcedCalendar", "曜日調整は OFF・2曜日・7曜日 のいずれかです。"
        End Select
        Select Case UCase$(Trim$(modSeriesSettings.SettingText(cfg, "X13_LEAP_YEAR_MODE")))
            Case "OFF", ""
            Case "ON": text = Trim$(text & " lpyear")
            Case Else: Err.Raise vbObjectError + 2902, "ForcedCalendar", "うるう年調整は ON・OFF のいずれかです。"
        End Select
        Select Case UCase$(Trim$(modSeriesSettings.SettingText(cfg, "X13_HOLIDAY_ADJUST")))
            Case "OFF", ""
            Case "ON": text = Trim$(text & " hol")
            Case Else: Err.Raise vbObjectError + 2903, "ForcedCalendar", "祝日調整は ON・OFF のいずれかです。"
        End Select
    End If
    Set tokens = ParseTokens(text, True)
    For Each key In tokens.Keys
        ForcedCalendar = Trim$(ForcedCalendar & " " & CStr(key))
    Next key
End Function

' 目的  : 空白区切りの変数を検査してDictionary（順序を保つ）にする。
' 入力  : requireCompatible = 曜日2種類の同時指定を拒否するか（必ず使う変数の検査ではTrue）
' エラー: 4つ以外の変数、重複は停止する。
Private Function ParseTokens(ByVal text As String, ByVal requireCompatible As Boolean) As Object
    Dim result As Object, token As Variant, name As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each token In Split(LCase$(Trim$(text)), " ")
        name = Trim$(CStr(token))
        Select Case name
            Case ""
            Case "tdnolpyear", "td1nolpyear", "lpyear", "hol"
                If result.Exists(name) Then Err.Raise vbObjectError + 2904, "ParseTokens", "カレンダー変数が重複しています：" & name
                result.Add name, True
            Case "none"
            Case Else
                Err.Raise vbObjectError + 2905, "ParseTokens", "カレンダー変数は tdnolpyear・td1nolpyear・lpyear・hol から選んでください：" & name
        End Select
    Next token
    If requireCompatible And result.Exists("tdnolpyear") And result.Exists("td1nolpyear") Then _
        Err.Raise vbObjectError + 2906, "ParseTokens", "tdnolpyear と td1nolpyear は代わりになる候補です。同時には指定できません。"
    Set ParseTokens = result
End Function

Private Function MergeTokens(ByVal first As String, ByVal second As String) As String
    Dim seen As Object, token As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    For Each token In Split(Trim$(first & " " & second), " ")
        If Len(CStr(token)) > 0 And CStr(token) <> "none" Then
            If Not seen.Exists(CStr(token)) Then
                seen.Add CStr(token), True
                MergeTokens = Trim$(MergeTokens & " " & CStr(token))
            End If
        End If
    Next token
End Function

' 曜日の変数が2つ以上入っていないか。
Private Function IsCompatible(ByVal calendar As String) As Boolean
    IsCompatible = Not (InStr(" " & calendar & " ", " tdnolpyear ") > 0 And InStr(" " & calendar & " ", " td1nolpyear ") > 0)
End Function

' 目的: SPCのregression.variablesへ入れる変数（holを除く）。"none"や空は空文字。
Public Function RegressionVariables(ByVal calendar As String) As String
    Dim token As Variant
    For Each token In Split(calendar, " ")
        Select Case CStr(token)
            Case "tdnolpyear", "td1nolpyear", "lpyear"
                RegressionVariables = Trim$(RegressionVariables & " " & CStr(token))
        End Select
    Next token
End Function

' 目的: 祝日変数（hol）を使う候補か。
Public Function UsesHoliday(ByVal calendar As String) As Boolean
    UsesHoliday = (InStr(" " & calendar & " ", " hol ") > 0)
End Function

' 目的: 表示用の日本語ラベル（例 "2曜日＋うるう年＋祝日"）。
Public Function CalendarLabel(ByVal calendar As String) As String
    Dim parts As String
    If InStr(" " & calendar & " ", " tdnolpyear ") > 0 Then parts = "7曜日"
    If InStr(" " & calendar & " ", " td1nolpyear ") > 0 Then parts = "2曜日"
    If InStr(" " & calendar & " ", " lpyear ") > 0 Then parts = parts & IIf(Len(parts) > 0, "＋", "") & "うるう年"
    If UsesHoliday(calendar) Then parts = parts & IIf(Len(parts) > 0, "＋", "") & "祝日"
    If Len(parts) = 0 Then parts = "カレンダー調整なし"
    CalendarLabel = parts
End Function
