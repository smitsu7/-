Attribute VB_Name = "modX13Summary"
Option Explicit

' 役割      : 1回の実行の保存済み試行を取りまとめ、比較できる群の中で順位を付け、最終決定（自動／手動）を行い、
'             系列別のExcelへ保存する。値の読取りは modX13Results、保存は modX13Store に任せる。
' 主な入口  : RefreshSummary（取りまとめ）/ AutoFinalize（自動決定）/ FinalizeById・FinalizeSelectedFromUI（手動決定）
'             ExportSummaries（系列別Excel）/ CopySelectedOutliersFromUI（検出外れ値を指定表へ転記する補助）
' 依存先    : modX13Store、modX13Results、modX13Calendar、modX13Provisional、modX13IO、modMain、modSheetNames、
'             modPackagePaths、modUI、modLog
' 書く場所  : 試行_候補比較・試行_変数検定・試行_探索履歴・試行_検出外れ値・試行_実行記録（実行ごとに作り直す）
'             04_results\summaries\<実行ID>\<系列ID>_候補比較_<時刻>.xlsx
' 比較群    : 同じ系列・同じ実行で、DAT（原値とゼロ置換の結果）・期間・変換・季節調整の型・差分次数・
'             外れ値の月と種類（自動検出・固定・指定）が同じ試行。群が違う試行には共通の順位を付けない。
'             固定外れ値とは月と種類の固定であり、係数の固定ではない。
' 統計量    : AIC・BIC・AICC・HQ は最終推定（UDGの aic・bic・aicc・hq）の値。探索履歴のBIC2は使わない。
'             取得できない値は空欄（"NA"と表示）のままにし、0や合格で埋めない。
' 予測誤差  : UDGの aape.0（直近3年）・aape.1～3（1年目～3年目）。aape.mode=withinsample の標本内の値で、
'             独立した将来データでの評価ではない。1か月先・12か月先の予測誤差（history）とは別の指標。

Private mRecords As Collection
Private mRunId As String

' ================= 見出し =================

Public Function ComparisonHeaders() As Variant
    ComparisonHeaders = Array("試行ID", "系列ID", "実行ID", "候補ID", "段階", "カレンダー", "ARIMAの決め方", "ARIMA次数_最終", _
        "外れ値の扱い", "推定した回帰変数", "比較群", "比較群の条件", "AIC", "BIC", "AICC", "HQ", "収束", _
        "予測誤差AAPEの方式", "AAPE_直近3年", "AAPE_1年目", "AAPE_2年目", "AAPE_3年目", "変数の結合検定", _
        "統計量の順位", "条件付き順位", "自動決定の対象", "対象外の理由", "保存フォルダ", "由来")
End Function

Public Function RegressionHeaders() As Variant
    RegressionHeaders = Array("試行ID", "系列ID", "変数・検定", "変数群", "係数", "標準誤差", "t値", "p値", "有意性", "検定の種類", _
        "検定統計量", "自由度", "出典ファイル")
End Function

Public Function SearchHistoryHeaders() As Variant
    SearchHistoryHeaders = Array("試行ID", "系列ID", "候補ID", "探索段階", "探索回", "順序", "ARIMA次数", "状態", "BIC", "BIC2", _
        "探索内のBIC2順位", "元の出力の行", "元の出力ファイル", "注記")
End Function

Public Function DetectedOutlierHeaders() As Variant
    DetectedOutlierHeaders = Array("試行ID", "系列ID", "候補ID", "種類", "年月", "係数", "標準誤差", "t値", "UDGのキー", _
        "設定_外れ値指定への記入例")
End Function

Public Function ExecutionHeaders() As Variant
    ExecutionHeaders = Array("実行ID", "候補ID", "系列ID", "設定ID", "段階", "状態", "理由", "X-13終了コード", "試行ID", "作業フォルダ")
End Function

Public Sub EnsureSummarySheets()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.ComparisonSheet()), ComparisonHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.RegressionSheet()), RegressionHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.SearchHistorySheet()), SearchHistoryHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.DetectedOutlierSheet()), DetectedOutlierHeaders()
    modX13IO.Headers modMain.EnsureWorksheet(modSheetNames.ExecutionRecordSheet()), ExecutionHeaders()
End Sub

' ================= 取りまとめ =================

' 目的  : 実行IDの保存済み試行をすべて検査して読み、5つのシートを作り直す。
' 副作用: 試行_候補比較・試行_変数検定・試行_探索履歴・試行_検出外れ値・試行_実行記録を作り直す。
' エラー: 保存済み試行の検査に失敗した場合は停止する（壊れた試行を比較に入れないため）。
'         設定の誤り（順位の統計量・有意水準）は停止する。
Public Sub RefreshSummary(ByVal runId As String)
    Dim metric As String, alpha As Double, requireCalendar As Boolean, trialId As Variant, meta As Object, values As Object
    Dim evidence As Object, item As Object, comparison As Collection, regression As Collection, history As Collection
    Dim outliers As Collection, calendarOk As Boolean, testText As String, key As Variant, groupLabels As Object
    Dim seriesGroups As Object, record As Variant, reasonText As String

    metric = RankMetric()
    alpha = CalendarAlpha()
    requireCalendar = modMain.GetControlBoolean("X13_REQUIRE_CALENDAR_SIGNIFICANCE", False)
    EnsureSummarySheets
    Set mRecords = New Collection
    mRunId = runId
    Set comparison = New Collection
    Set regression = New Collection
    Set history = New Collection
    Set outliers = New Collection
    Set groupLabels = CreateObject("Scripting.Dictionary")

    For Each trialId In modX13Store.TrialsOfRun(runId)
        Application.StatusBar = "取りまとめ：" & CStr(trialId)
        Set meta = modX13Store.ValidateTrial(CStr(trialId))
        Set values = modX13Results.ReadUdgValues(modMain.CombinePath(meta("FOLDER"), meta("STEM") & ".udg"))
        Set evidence = modX13Results.ReadOutputEvidence(CStr(meta("FOLDER")), CStr(meta("STEM")))
        calendarOk = AppendRegression(regression, meta, values, alpha, testText)
        AppendOutliers outliers, meta, values
        If meta("MODEL_MODE") = "automdl" Then AppendSearchHistory history, meta

        Set item = CreateObject("Scripting.Dictionary")
        item.Add "Trial", CStr(trialId)
        item.Add "Series", CStr(meta("SERIES"))
        item.Add "Stage", CStr(meta("STAGE"))
        item.Add "GroupKey", GroupKey(meta, values)
        item.Add "Group", GroupLabel(groupLabels, CStr(meta("SERIES")), CStr(item("GroupKey")))
        For Each key In Array("AIC", "BIC", "AICC", "HQ")
            item.Add CStr(key), modX13Results.NumericValue(values, LCase$(CStr(key)))
        Next key
        item.Add "Metric", item(metric)
        item.Add "MetricName", metric
        item.Add "Converged", modX13Results.ReadValue(values, "converged")
        item.Add "CalendarOk", calendarOk
        reasonText = ""
        If IsEmpty(item("Metric")) Then reasonText = metric & "を取得できない"
        If item("Converged") <> "yes" Then reasonText = Join2(reasonText, "推定が収束していない（converged=" & item("Converged") & "）")
        If requireCalendar And Not calendarOk Then reasonText = Join2(reasonText, "変数の結合検定が有意でない、または取得できない")
        item.Add "Eligible", (Len(reasonText) = 0)
        item.Add "Reason", reasonText
        item.Add "Rank", Empty
        item.Add "FilteredRank", Empty
        item.Add "Row", Array(trialId, meta("SERIES"), meta("RUN"), meta("CANDIDATE"), meta("STAGE"), _
            modX13Calendar.CalendarLabel(CStr(meta("CALENDAR"))), IIf(meta("MODEL_MODE") = "automdl", "自動選択", "次数指定"), _
            FinalModel(values, evidence), meta("OUTLIER_MODE"), EstimatedVariables(values), item("Group"), GroupDescription(meta, values), _
            ShowNA(item("AIC")), ShowNA(item("BIC")), ShowNA(item("AICC")), ShowNA(item("HQ")), item("Converged"), _
            modX13Results.ReadValue(values, "aape.mode"), ShowNA(modX13Results.NumericValue(values, "aape.0")), _
            ShowNA(modX13Results.NumericValue(values, "aape.1")), ShowNA(modX13Results.NumericValue(values, "aape.2")), _
            ShowNA(modX13Results.NumericValue(values, "aape.3")), testText, "", "", IIf(item("Eligible"), "対象", "対象外"), _
            reasonText, modPackagePaths.RelativeToPackage(CStr(meta("FOLDER"))), OriginText(meta))
        mRecords.Add item
    Next trialId

    CalculateRanks
    For Each record In mRecords
        comparison.Add WithRanks(record)
    Next record
    WriteSheet modSheetNames.ComparisonSheet(), ComparisonHeaders(), comparison, Array(13, 14, 15, 16, 19, 20, 21, 22, 24, 25)
    WriteSheet modSheetNames.RegressionSheet(), RegressionHeaders(), regression, Array(5, 6, 7, 8, 11)
    WriteSheet modSheetNames.SearchHistorySheet(), SearchHistoryHeaders(), history, Array(5, 6, 9, 10, 11, 12)
    WriteSheet modSheetNames.DetectedOutlierSheet(), DetectedOutlierHeaders(), outliers, Array(6, 7, 8)
    WriteSheet modSheetNames.ExecutionRecordSheet(), ExecutionHeaders(), ExecutionRows(runId), Empty
    Application.StatusBar = False
    modLog.WriteLog "RefreshSummary", "INFO", runId & " 試行=" & mRecords.Count & " 統計量=" & metric
End Sub

Private Function RankMetric() As String
    RankMetric = UCase$(modMain.GetControlValue("X13_RANK_METRIC", "AICC"))
    If InStr("|AIC|BIC|AICC|HQ|", "|" & RankMetric & "|") = 0 Then Err.Raise vbObjectError + 3500, "RankMetric", _
        "X13_RANK_METRICは AIC・BIC・AICC・HQ のいずれかです：" & RankMetric
End Function

Private Function CalendarAlpha() As Double
    CalendarAlpha = modX13Options.CheckedNumber(modMain.GetControlValue("X13_CALENDAR_ALPHA", "0.05"), "X13_CALENDAR_ALPHA")
    If CalendarAlpha <= 0 Or CalendarAlpha >= 1 Then Err.Raise vbObjectError + 3501, "CalendarAlpha", _
        "X13_CALENDAR_ALPHAは0より大きく1より小さい値です"
End Function

Private Function Join2(ByVal first As String, ByVal second As String) As String
    If Len(first) = 0 Then Join2 = second Else Join2 = first & "；" & second
End Function

Private Function ShowNA(ByVal value As Variant) As Variant
    If IsEmpty(value) Then ShowNA = "NA" Else ShowNA = value
End Function

' 比較群の照合に使う文字列。保存したファイルと記録だけから作る（現在のシートの設定を見ない）。
Private Function GroupKey(ByVal meta As Object, ByVal values As Object) As String
    GroupKey = meta("RUN") & "|" & meta("SERIES") & "|" & meta("FIRST") & "|" & meta("ACTUAL_MONTHS") & "|" & meta("FORECAST_MONTHS") & _
        "|" & meta("DATA_CRC") & "|" & modX13Results.ReadValue(values, "transform") & "|" & modX13Results.ReadValue(values, "samode") & _
        "|d" & modX13Results.ReadValue(values, "nonseasonaldiff") & "D" & modX13Results.ReadValue(values, "seasonaldiff") & _
        "|" & modX13Results.OutlierSignature(values) & "|" & SortedTokens(CStr(meta("SPECIFIED_OUTLIERS")))
End Function

' 人が読む比較群の条件。
Private Function GroupDescription(ByVal meta As Object, ByVal values As Object) As String
    Dim fixedText As String
    fixedText = Mid$(modX13Results.OutlierSignature(values), 2)
    If Len(fixedText) = 0 Then fixedText = "なし"
    GroupDescription = "期間 " & meta("FIRST") & "から" & meta("ACTUAL_MONTHS") & "か月；変換 " & modX13Results.ReadValue(values, "transform") & _
        "；型 " & modX13Results.ReadValue(values, "samode") & "；差分 d=" & modX13Results.ReadValue(values, "nonseasonaldiff") & _
        " D=" & modX13Results.ReadValue(values, "seasonaldiff") & "；外れ値（月と種類）" & Replace(fixedText, ";", " ") & _
        IIf(Len(CStr(meta("SPECIFIED_OUTLIERS"))) > 0, "；指定 " & meta("SPECIFIED_OUTLIERS"), "") & "；DAT CRC32 " & meta("DATA_CRC")
End Function

' 系列ごとに G1, G2 … の名前を付ける（出てきた順）。
Private Function GroupLabel(ByVal labels As Object, ByVal seriesId As String, ByVal key As String) As String
    Dim count As Long, k As Variant
    If labels.Exists(key) Then
        GroupLabel = labels(key)
        Exit Function
    End If
    For Each k In labels.Keys
        If Left$(CStr(k), Len(mRunId & "|" & seriesId & "|")) = mRunId & "|" & seriesId & "|" Then count = count + 1
    Next k
    GroupLabel = "G" & (count + 1)
    labels.Add key, GroupLabel
End Function

Private Function SortedTokens(ByVal text As String) As String
    Dim parts As Variant, i As Long, j As Long, swap As String
    text = LCase$(Trim$(text))
    If Len(text) = 0 Then Exit Function
    parts = Split(text, " ")
    For i = 0 To UBound(parts)
        For j = i + 1 To UBound(parts)
            If parts(j) < parts(i) Then swap = parts(i): parts(i) = parts(j): parts(j) = swap
        Next j
    Next i
    SortedTokens = Join(parts, " ")
End Function

Private Function FinalModel(ByVal values As Object, ByVal evidence As Object) As String
    FinalModel = modX13Results.ReadValue(values, "arimamdl")
    If Len(FinalModel) = 0 And evidence.Exists("X13_ACTUAL_MODEL") Then FinalModel = CStr(evidence("X13_ACTUAL_MODEL"))
    If Len(FinalModel) = 0 Then FinalModel = "NA"
End Function

Private Function EstimatedVariables(ByVal values As Object) As String
    Dim coefficients As Object, key As Variant, complete As Boolean
    Set coefficients = modX13Results.RegressionCoefficients(values, complete)
    For Each key In coefficients.Keys
        EstimatedVariables = EstimatedVariables & IIf(Len(EstimatedVariables) > 0, "; ", "") & CStr(key)
    Next key
    If Not complete Then EstimatedVariables = Join2(EstimatedVariables, "回帰係数の件数がnregと一致しない")
    If Len(EstimatedVariables) = 0 Then EstimatedVariables = "なし"
End Function

' 由来の列。段階に続けて、X-13の出力形式（ASCII版／HTML版）と主出力のファイル名を示す。
Private Function OriginText(ByVal meta As Object) As String
    OriginText = IIf(meta("STAGE") = "再推定", "再推定（基準試行 " & meta("REFERENCE_TRIAL") & "、固定外れ値 " & _
        IIf(Len(CStr(meta("FROZEN_OUTLIERS"))) > 0, meta("FROZEN_OUTLIERS"), "なし") & "）", "初回")
    If meta.Exists("REPORT_MAIN") Then
        OriginText = OriginText & "；出力 " & meta("REPORT_FORMAT") & "版（" & meta("REPORT_MAIN") & "）"
    Else
        OriginText = OriginText & "；出力 ASCII版（HTML対応前の記録）"
    End If
End Function

' 目的  : 比較群の中で統計量が小さい順に順位を付ける。同じ値は同じ順位。
'         統計量の順位 = 群内で統計量を取得できた全試行の中での順位
'         条件付き順位 = 群内で「自動決定の対象」の試行だけの中での順位（変数検定の合否等を考慮した順位）
Private Sub CalculateRanks()
    Dim item As Variant, other As Variant, rawRank As Long, filteredRank As Long
    For Each item In mRecords
        If Not IsEmpty(item("Metric")) Then
            rawRank = 1: filteredRank = 1
            For Each other In mRecords
                If other("GroupKey") = item("GroupKey") And Not IsEmpty(other("Metric")) Then
                    If CDbl(other("Metric")) < CDbl(item("Metric")) Then
                        rawRank = rawRank + 1
                        If other("Eligible") Then filteredRank = filteredRank + 1
                    End If
                End If
            Next other
            item("Rank") = rawRank
            If item("Eligible") Then item("FilteredRank") = filteredRank
        End If
    Next item
End Sub

Private Function WithRanks(ByVal record As Object) As Variant
    Dim row As Variant
    row = record("Row")
    row(23) = ShowNA(record("Rank"))
    row(24) = ShowNA(record("FilteredRank"))
    WithRanks = row
End Function

' ---- 変数検定 ----

' 目的  : 1試行の回帰係数（t値・p値）、X-13の結合F検定、RCMによるWald検定を行として加える。
' 出力  : 必要な変数群（曜日・うるう年・祝日）の結合検定がすべて取得でき、有意水準未満ならTrue。
'         必要な変数群が無い試行（カレンダー調整なし）はTrue（検定不要）。
'         testText に、変数群ごとの結果の短い説明を返す。
Private Function AppendRegression(ByVal rows As Collection, ByVal meta As Object, ByVal values As Object, _
                                  ByVal alpha As Double, ByRef testText As String) As Boolean
    Dim coefficients As Object, complete As Boolean, needed As Object, tested As Object, key As Variant, nums As Variant
    Dim p As Double, group As String, good As Boolean, source As String, rcm As String, test As Object, base As String
    base = modMain.CombinePath(meta("FOLDER"), meta("STEM"))
    source = modPackagePaths.RelativeToPackage(base & ".udg")
    rcm = base & ".rcm"
    Set coefficients = modX13Results.RegressionCoefficients(values, complete)
    Set needed = modX13Results.ExpectedCalendarGroups(base & ".spc")
    Set tested = CreateObject("Scripting.Dictionary")
    good = True
    testText = ""
    If Not complete Then
        good = False
        rows.Add Array(meta("TRIAL"), meta("SERIES"), "回帰係数の取得", "", "NA", "NA", "NA", "NA", "未取得", _
            "nregと回帰係数の件数が一致しない", "NA", "NA", source)
    End If
    For Each key In coefficients.Keys
        nums = coefficients(key)
        p = 2 * Application.WorksheetFunction.Norm_S_Dist(-Abs(CDbl(nums(2))), True)
        rows.Add Array(meta("TRIAL"), meta("SERIES"), CStr(key), GroupName(modX13Results.CalendarGroup(CStr(key))), nums(0), nums(1), nums(2), p, _
            IIf(p < alpha, "有意", "有意でない"), "t値から求めた正規近似のp値（両側）", "NA", "NA", source)
    Next key
    ' X-13の結合F検定（曜日の変数群）。
    For Each key In values.Keys
        If Left$(CStr(key), 6) = "ftest$" Then
            nums = modX13Results.Numbers(CStr(values(key)))
            If IsArray(nums) Then
                If UBound(nums) = 3 Then
                    p = CDbl(nums(3))
                    group = modX13Results.CalendarGroup(CStr(key))
                    If Len(group) > 0 And needed.Exists(group) Then
                        tested(group) = True
                        If p >= alpha Then good = False
                        testText = Join2(testText, GroupName(group) & " F検定 p=" & Format$(p, "0.0000"))
                    End If
                    rows.Add Array(meta("TRIAL"), meta("SERIES"), CStr(key), GroupName(group), "NA", "NA", "NA", p, _
                        IIf(p < alpha, "有意", "有意でない"), "X-13の結合F検定", nums(2), "df1=" & nums(0) & "; df2=" & nums(1), source)
                End If
            End If
        End If
    Next key
    ' F検定が無い変数群（うるう年・祝日等）は、最終推定の係数と共分散（RCM）からWaldカイ二乗検定を計算する。
    For Each key In needed.Keys
        If Not tested.Exists(CStr(key)) Then
            If complete Then
                Set test = modX13Results.WaldGroupTest(coefficients, CStr(key), rcm)
            Else
                Set test = CreateObject("Scripting.Dictionary")
                test.Add "Available", False
                test.Add "Reason", "回帰係数をすべて取得できない"
            End If
            If test("Available") Then
                p = CDbl(test("P"))
                If p >= alpha Then good = False
                testText = Join2(testText, GroupName(CStr(key)) & " Wald検定 p=" & Format$(p, "0.0000"))
                rows.Add Array(meta("TRIAL"), meta("SERIES"), GroupName(CStr(key)) & "の結合検定", GroupName(CStr(key)), "NA", "NA", "NA", p, _
                    IIf(p < alpha, "有意", "有意でない"), "最終推定の係数と共分散（RCM）によるWaldカイ二乗検定", test("Statistic"), test("Df"), _
                    source & "; " & modPackagePaths.RelativeToPackage(rcm))
            Else
                good = False
                testText = Join2(testText, GroupName(CStr(key)) & " 未取得")
                rows.Add Array(meta("TRIAL"), meta("SERIES"), GroupName(CStr(key)) & "の結合検定", GroupName(CStr(key)), "NA", "NA", "NA", "NA", _
                    "未取得", "Wald検定を計算できない：" & test("Reason"), "NA", "NA", source)
            End If
        End If
    Next key
    If needed.Count = 0 Then testText = "カレンダー変数なし（検定不要）"
    AppendRegression = good
End Function

Private Function GroupName(ByVal group As String) As String
    Select Case group
        Case "weekday": GroupName = "曜日"
        Case "leap": GroupName = "うるう年"
        Case "holiday": GroupName = "祝日"
        Case Else: GroupName = ""
    End Select
End Function

' ---- 検出外れ値・探索履歴 ----

Private Sub AppendOutliers(ByVal rows As Collection, ByVal meta As Object, ByVal values As Object)
    Dim item As Variant
    For Each item In modX13Results.DetectedOutliers(values)
        rows.Add Array(meta("TRIAL"), meta("SERIES"), meta("CANDIDATE"), item(0), item(1), item(2), item(3), item(4), item(5), _
            meta("SERIES") & " | " & meta("SETTING") & " | " & item(0) & " | " & item(1) & " |")
    Next item
End Sub

' 探索履歴の行。「元の出力の行」は、ASCII版は .out の行、HTML版はVBAがHTMLから作った解析用テキストの行。
Private Sub AppendSearchHistory(ByVal rows As Collection, ByVal meta As Object)
    Dim path As String, item As Variant, note As String
    path = modX13Results.ReportTextPath(CStr(meta("FOLDER")), CStr(meta("STEM")))
    note = "自動選択の探索中の値。候補比較のAIC・BIC・AICC・HQとは別"
    If modX13Store.TrialReportFormat(meta) = modX13Results.REPORT_HTML Then _
        note = note & "。行番号はHTML版の原本 " & meta("STEM") & ".html から作った解析用テキストの行"
    For Each item In modX13Results.ParseSearchHistory(modX13Results.ReportText(CStr(meta("FOLDER")), CStr(meta("STEM"))))
        rows.Add Array(meta("TRIAL"), meta("SERIES"), meta("CANDIDATE"), item("Stage"), item("Pass"), item("Sequence"), item("Model"), _
            item("State"), ShowNA(item("BIC")), ShowNA(item("BIC2")), ShowNA(item("Rank")), item("SourceLine"), _
            modPackagePaths.RelativeToPackage(path), note)
    Next item
End Sub

' 試行_ジョブから、この実行の全候補（失敗・実行前エラーを含む）を読む。
Private Function ExecutionRows(ByVal runId As String) As Collection
    Dim ws As Worksheet, r As Long, rows As New Collection
    Set ws = modMain.WorksheetOrNothing(modSheetNames.JobSheet())
    If Not ws Is Nothing Then
        For r = 2 To modX13IO.NextRow(ws) - 1
            If CStr(ws.Cells(r, 1).Value2) = runId Then
                rows.Add Array(ws.Cells(r, 1).Value2, ws.Cells(r, 2).Value2, ws.Cells(r, 3).Value2, ws.Cells(r, 4).Value2, _
                    ws.Cells(r, 5).Value2, ws.Cells(r, 9).Value2, ws.Cells(r, 10).Value2, ws.Cells(r, 11).Value2, _
                    ws.Cells(r, 12).Value2, ws.Cells(r, 8).Value2)
            End If
        Next r
    End If
    Set ExecutionRows = rows
End Function

' シートを作り直して行を書く。numericColumns は数値として書く列。
Private Sub WriteSheet(ByVal sheetName As String, ByVal headers As Variant, ByVal rows As Collection, ByVal numericColumns As Variant)
    Dim ws As Worksheet, startRow As Long, count As Long
    Set ws = modMain.EnsureWorksheet(sheetName)
    modX13IO.ResetSheet ws, headers
    If rows.Count > 0 Then
        If IsArray(numericColumns) Then
            modX13IO.AppendRows ws, modX13IO.RowsToArray(rows, UBound(headers) + 1), startRow, count, numericColumns
        Else
            modX13IO.AppendRows ws, modX13IO.RowsToArray(rows, UBound(headers) + 1), startRow, count
        End If
    End If
    ws.Rows(1).WrapText = True
    ws.Columns.ColumnWidth = 14
End Sub

' ================= 最終決定 =================

' 目的  : 系列ごとに、比較群の中で条件を満たす最上位の試行を最終決定として保存する（X13_FINAL_MODE=自動）。
' 規則  : 比較群は X13_COMPARISON_GROUP。AUTO_REFERENCE は、その系列で最初に保存した試行（初回の最初の候補）の群。
'         G1・G2等を書いた場合は、各系列の同じ名前の群。統計量が同じ値の候補が複数あれば、先に保存した試行を選び、
'         同じ値の件数を最終_決定に記録する。
' エラー: 条件を満たす候補が無い系列は保存せず、ログと試行_実行記録の後ろの注記に残して次の系列へ進む。
Public Sub AutoFinalize(ByVal runId As String)
    Dim series As Object, item As Variant, sid As Variant, configured As String, wanted As String
    Dim best As Object, ties As Long, reason As String, failures As String
    If mRecords Is Nothing Then RefreshSummary runId
    If mRunId <> runId Then RefreshSummary runId
    configured = modMain.GetControlValue("X13_COMPARISON_GROUP", "AUTO_REFERENCE")
    Set series = CreateObject("Scripting.Dictionary")
    For Each item In mRecords
        If Not series.Exists(item("Series")) Then series.Add item("Series"), item("Group")
    Next item
    For Each sid In series.Keys
        If configured = "AUTO_REFERENCE" Then wanted = series(sid) Else wanted = configured
        Set best = Nothing
        ties = 0
        For Each item In mRecords
            If item("Series") = sid And item("Group") = wanted And item("Eligible") Then
                If best Is Nothing Then
                    Set best = item: ties = 1
                ElseIf CDbl(item("Metric")) < CDbl(best("Metric")) Then
                    Set best = item: ties = 1
                ElseIf CDbl(item("Metric")) = CDbl(best("Metric")) Then
                    ties = ties + 1
                End If
            End If
        Next item
        If best Is Nothing Then
            failures = failures & vbLf & CStr(sid) & "：比較群" & wanted & "に自動決定の対象となる試行がありません"
            modLog.WriteLog "AutoFinalize", "WARN", CStr(sid) & " 比較群" & wanted & "に対象の試行がないため最終保存しません"
        Else
            reason = TrySaveDecision(best, "自動", ties, "比較群 " & wanted & " の条件付き順位1位（設定 X13_COMPARISON_GROUP=" & configured & "）")
            If Len(reason) > 0 Then failures = failures & vbLf & CStr(sid) & "：最終保存に失敗しました（" & reason & "）"
        End If
    Next sid
    modX13Provisional.RefreshProvisional
    If Len(failures) > 0 Then modUI.ShowMessage "自動決定できなかった系列があります。" & failures, vbExclamation, "自動決定"
End Sub

' 目的  : 試行IDを指定して最終決定する（利用者の手動決定、または試験・自動化）。
' 入力  : trialId、mode（"手動"等、最終_決定に記録する決定方式）
' 出力  : 保存ID
' エラー: 現在の候補比較に無い試行、自動決定の対象外の試行（統計量なし・未収束・変数検定の条件不成立）は停止する。
Public Function FinalizeById(ByVal trialId As String, Optional ByVal mode As String = "") As String
    Dim item As Variant, found As Object, runId As String, meta As Object
    If Len(mode) = 0 Then mode = "手動"
    Set meta = modX13Store.ValidateTrial(trialId)
    runId = CStr(meta("RUN"))
    If mRecords Is Nothing Then RefreshSummary runId
    If mRunId <> runId Then RefreshSummary runId
    For Each item In mRecords
        If item("Trial") = trialId Then Set found = item
    Next item
    If found Is Nothing Then Err.Raise vbObjectError + 3502, "FinalizeById", "候補比較に試行IDがありません：" & trialId
    If Not found("Eligible") Then Err.Raise vbObjectError + 3503, "FinalizeById", "この試行は最終決定の条件を満たしません：" & found("Reason")
    FinalizeById = SaveDecision(found, mode, 0, "利用者が候補比較から選択")
    modX13Provisional.RefreshProvisional
    ExportSummaries runId, CStr(found("Series"))
End Function

' 目的: 試行_候補比較で選んだ行の試行を、確認のうえ最終決定する。
Public Sub FinalizeSelectedFromUI()
    Dim ws As Worksheet, trialId As String, adoption As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(modSheetNames.ComparisonSheet())
    If Not ActiveSheet Is ws Then Err.Raise vbObjectError + 3504, , "「" & ws.Name & "」で、最終決定する試行の行を選んでから実行してください。"
    If ActiveCell.Row < 2 Then Err.Raise vbObjectError + 3505, , "見出しではなく試行の行を選んでください。"
    trialId = CStr(ws.Cells(ActiveCell.Row, 1).Value2)
    If Len(trialId) = 0 Then Err.Raise vbObjectError + 3506, , "選んだ行に試行IDがありません。"
    If Not modUI.Confirm("系列「" & ws.Cells(ActiveCell.Row, 2).Value2 & "」の試行「" & trialId & "」を最終決定として保存します。" & vbLf & _
        "比較群 " & ws.Cells(ActiveCell.Row, 11).Value2 & "、" & ws.Cells(1, 24).Value2 & " " & ws.Cells(ActiveCell.Row, 24).Value2 & "。" & vbLf & _
        "以前の最終結果のフォルダは残し、表では「過去の採用」にします。続けますか？", "最終決定") Then Exit Sub
    adoption = FinalizeById(trialId, "手動")
    modUI.ShowMessage "最終決定として保存しました。保存ID：" & adoption & vbLf & _
        "月次の結果は「" & modSheetNames.FinalMonthlySheet() & "」、暫定季節指数は「" & modSheetNames.ProvisionalSheet() & "」です。", _
        vbInformation, "最終決定"
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "最終決定"
End Sub

' 目的: 画面から、候補比較に表示中の実行について自動決定を行う。
Public Sub AutoFinalizeFromUI()
    Dim runId As String
    On Error GoTo Failed
    runId = DisplayedRunId()
    If Not modUI.Confirm("実行「" & runId & "」の各系列について、条件付き順位1位の試行を最終決定として保存します。続けますか？", "自動決定") Then Exit Sub
    AutoFinalize runId
    ExportSummaries runId
    ThisWorkbook.Worksheets(modSheetNames.DecisionSheet()).Activate
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "自動決定"
End Sub

' 目的: 画面から、候補比較に表示中の実行の系列別Excelを作り直す（新しいファイルとして保存）。
Public Sub ExportFromUI()
    On Error GoTo Failed
    ExportSummaries DisplayedRunId()
    modUI.ShowMessage "系列別のExcelを保存しました。保存先：" & modPackagePaths.SettingPath("X13_SUMMARY_DIR"), vbInformation, "系列別Excel"
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "系列別Excel"
End Sub

Private Function DisplayedRunId() As String
    Dim ws As Worksheet
    Set ws = modMain.WorksheetOrNothing(modSheetNames.ComparisonSheet())
    If Not ws Is Nothing Then DisplayedRunId = CStr(ws.Cells(2, 3).Value2)
    If Len(DisplayedRunId) = 0 Then Err.Raise vbObjectError + 3507, "DisplayedRunId", "候補比較に試行がありません。先に季節調整を実行してください。"
End Function

' 自動決定で1系列の保存に失敗しても、他の系列を続けるため、理由を文字列で返す。
Private Function TrySaveDecision(ByVal record As Object, ByVal mode As String, ByVal ties As Long, ByVal note As String) As String
    On Error GoTo Failed
    SaveDecision record, mode, ties, note
    Exit Function
Failed:
    TrySaveDecision = Err.Description
    modLog.WriteLog "AutoFinalize", "ERROR", record("Series") & " " & Err.Description
End Function

' 決定を modX13Store へ渡して保存する。
Private Function SaveDecision(ByVal record As Object, ByVal mode As String, ByVal ties As Long, ByVal note As String) As String
    Dim decision As Object
    Set decision = CreateObject("Scripting.Dictionary")
    decision.Add "mode", mode
    decision.Add "metric", record("MetricName")
    decision.Add "metricValue", modX13IO.NumberText(CDbl(record("Metric")))
    decision.Add "group", record("Group")
    decision.Add "rank", CStr(record("Rank"))
    decision.Add "filteredRank", CStr(record("FilteredRank"))
    decision.Add "requireCalendar", modMain.GetControlValue("X13_REQUIRE_CALENDAR_SIGNIFICANCE", "NO")
    decision.Add "alpha", modMain.GetControlValue("X13_CALENDAR_ALPHA", "0.05")
    decision.Add "ties", IIf(ties > 0, CStr(ties), "")
    decision.Add "note", note
    SaveDecision = modX13Store.FinalizeTrial(CStr(record("Trial")), decision)
End Function

' ================= 系列別Excel =================

' 目的  : 系列ごとに、候補比較・変数検定・探索履歴・検出外れ値・実行記録・最終決定・最終月次結果・暫定季節指数を
'         1つのExcelにまとめて保存する。新しいファイル名で保存し、既存ファイルは上書きしない。
' 入力  : runId、onlySeries（空欄なら実行の全系列）
' 副作用: 04_results\summaries\<実行ID>\ に .xlsx を作る。
Public Sub ExportSummaries(ByVal runId As String, Optional ByVal onlySeries As String = "")
    Dim root As String, folder As String, series As Object, item As Variant, sid As Variant, ws As Worksheet, r As Long
    Dim book As Workbook, path As String, n As Long, sheetSpec As Variant, oldAlerts As Boolean, errorNumber As Long, errorText As String
    root = modPackagePaths.RequireFolder("X13_SUMMARY_DIR", True)
    folder = modMain.CombinePath(root, runId)
    modPackagePaths.EnsureFolderTree folder
    Set series = CreateObject("Scripting.Dictionary")
    Set ws = modMain.WorksheetOrNothing(modSheetNames.JobSheet())
    If Not ws Is Nothing Then
        For r = 2 To modX13IO.NextRow(ws) - 1
            If CStr(ws.Cells(r, 1).Value2) = runId Then series(CStr(ws.Cells(r, 3).Value2)) = True
        Next r
    End If
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    For Each sid In series.Keys
        If Len(onlySeries) = 0 Or CStr(sid) = onlySeries Then
            Set book = Application.Workbooks.Add(xlWBATWorksheet)
            n = 0
            For Each sheetSpec In Array( _
                    Array(modSheetNames.ComparisonSheet(), "候補比較", 2), Array(modSheetNames.RegressionSheet(), "変数検定", 2), _
                    Array(modSheetNames.SearchHistorySheet(), "探索履歴", 2), Array(modSheetNames.DetectedOutlierSheet(), "検出外れ値", 2), _
                    Array(modSheetNames.ExecutionRecordSheet(), "実行記録", 3), Array(modSheetNames.DecisionSheet(), "最終決定", 3), _
                    Array(modSheetNames.FinalMonthlySheet(), "最終月次結果", 2), Array(modSheetNames.ProvisionalSheet(), "暫定季節指数", 2))
                n = n + 1
                CopyRowsOfSeries book, n, CStr(sheetSpec(0)), CStr(sheetSpec(1)), CLng(sheetSpec(2)), CStr(sid)
            Next sheetSpec
            AddReadingNotes book
            path = modMain.CombinePath(folder, modMain.SafeFileStem(CStr(sid)) & "_候補比較_" & modMain.TimeStamp() & ".xlsx")
            n = 0
            Do While modMain.FileExists(path)
                n = n + 1
                path = modMain.CombinePath(folder, modMain.SafeFileStem(CStr(sid)) & "_候補比較_" & modMain.TimeStamp() & "_" & n & ".xlsx")
            Loop
            Application.DisplayAlerts = False
            book.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
            Application.DisplayAlerts = oldAlerts
            book.Close SaveChanges:=False
            Set book = Nothing
        End If
    Next sid
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    ' 後始末：作りかけのブックを保存せずに閉じる。
    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errorNumber, "ExportSummaries", errorText
End Sub

' 元のシートから、系列ID列（seriesColumn）が一致する行を、新しいブックの index 番目のシートへ写す（値だけ）。
Private Sub CopyRowsOfSeries(ByVal book As Workbook, ByVal index As Long, ByVal sourceName As String, ByVal targetName As String, _
                             ByVal seriesColumn As Long, ByVal seriesId As String)
    Dim source As Worksheet, target As Worksheet, data As Variant, lastRow As Long, lastColumn As Long, r As Long, c As Long
    Dim out() As Variant, count As Long
    If index = 1 Then
        Set target = book.Worksheets(1)
    Else
        Set target = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    End If
    target.Name = targetName
    Set source = modMain.WorksheetOrNothing(sourceName)
    If source Is Nothing Then Exit Sub
    lastRow = modX13IO.NextRow(source) - 1
    lastColumn = source.Cells(1, source.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Then Exit Sub
    data = source.Range(source.Cells(1, 1), source.Cells(lastRow, lastColumn)).Value2
    If Not IsArray(data) Then Exit Sub
    ReDim out(1 To UBound(data, 1), 1 To lastColumn)
    For r = 1 To UBound(data, 1)
        If r = 1 Or CStr(data(r, seriesColumn)) = seriesId Then
            count = count + 1
            For c = 1 To lastColumn
                out(count, c) = data(r, c)
            Next c
        End If
    Next r
    target.Cells.NumberFormat = "@"
    target.Cells(1, 1).Resize(count, lastColumn).Value2 = out
    target.Rows(1).Font.Bold = True
    target.Columns.ColumnWidth = 14
End Sub

Private Sub AddReadingNotes(ByVal book As Workbook)
    Dim ws As Worksheet, notes As Variant, i As Long
    Set ws = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    ws.Name = "読み方"
    notes = Array( _
        "候補比較：同じ比較群（期間・DAT・変換・型・差分次数・外れ値の月と種類が同じ）の中で、統計量が小さい試行を上位にしています。群が違う試行には共通の順位を付けません。", _
        "統計量の順位は統計量を取得できた全試行の順位、条件付き順位は「自動決定の対象」の試行だけの順位です（変数検定を条件にする設定のときは、その合否を反映）。", _
        "AIC・BIC・AICC・HQはX-13の最終推定の値です。自動選択の探索中に表示されるBIC2は「探索履歴」だけにあり、順位には使いません。", _
        "予測誤差AAPEは、X-13の標本内の値（aape.mode=withinsample）です。直近3年と1年目～3年目の平均絶対誤差率で、独立した将来データでの評価ではありません。1か月先・12か月先の予測誤差とは別の指標です。", _
        "変数検定：個別の係数のp値はt値からの正規近似、曜日の結合検定はX-13のF検定、うるう年・祝日の結合検定は最終推定の係数と共分散（RCM）によるWaldカイ二乗検定です。NAは取得できなかった値です。", _
        "再推定の試行は、初回の基準試行で自動検出された外れ値の月と種類を固定して推定したものです（係数は固定していません）。", _
        "最終決定・最終月次結果は候補比較とは別に保存しています。D18の出所が「VBAの中立値」の行は、カレンダー調整なしのためX-13がD18を出さず、乗法型1・加法型0を補ったものです。", _
        "X-13はASCII版・HTML版のどちらでも計算は同じです。由来の列に、使った版と主出力のファイル名（.out／.html）を示します。HTML版の探索履歴の行番号は、原本の.htmlからVBAが作った解析用テキスト（_report_text_from_html.txt）の行です。", _
        "事後診断（check・spectrum・history・slidingspans）は今回のツールでは実行していません。")
    For i = 0 To UBound(notes)
        ws.Cells(i + 1, 1).Value2 = notes(i)
    Next i
    ws.Columns(1).ColumnWidth = 120
    ws.Columns(1).WrapText = True
End Sub

' ================= 従来手順の補助 =================

' 目的  : 試行_検出外れ値で選んだ行の「種類・年月」を、設定_外れ値指定へ書き写す（利用者が選んだ行だけ）。
'         従来手順（次数指定で外れ値を検出 → 利用者が月と種類を指定表へ記入 → 別の実行でモデル選択）の記入を助ける。
'         次の実行の設定（outlier_mode・model_mode・設定ID）は変えない。利用者が設定_系列別で選ぶ。
' 副作用: 設定_外れ値指定に行を追加する。同じ系列・設定ID・種類・月の行が既にあれば追加しない。
Public Sub CopySelectedOutliersFromUI()
    Dim source As Worksheet, target As Worksheet, area As Range, cell As Range, rows As Object, r As Variant
    Dim settingId As String, added As Long, skipped As Long, key As String, existing As Object, t As Long
    On Error GoTo Failed
    Set source = ThisWorkbook.Worksheets(modSheetNames.DetectedOutlierSheet())
    If Not ActiveSheet Is source Then Err.Raise vbObjectError + 3508, , "「" & source.Name & "」で転記する行を選んでから実行してください。"
    settingId = modUI.AskText("書き写す先の設定ID（setting_id）を入力してください。" & vbLf & _
        "例：次の実行用に step2 等。設定_系列別で、その設定IDの outlier_mode を「自動＋指定」または「指定のみ」にしてください。", _
        "検出外れ値の転記", modMain.GetControlValue("X13_SETTING_ID", "standard"))
    If Len(settingId) = 0 Then Exit Sub
    If Not modPublicImport.IsIdent(settingId) Then Err.Raise vbObjectError + 3509, , "設定IDは半角英数字と下線です：" & settingId
    modX13Interventions.SetupInterventionSheet
    Set target = ThisWorkbook.Worksheets(modSheetNames.InterventionSheet())
    Set existing = CreateObject("Scripting.Dictionary")
    For t = 2 To modX13IO.NextRow(target) - 1
        existing(CStr(target.Cells(t, 1).Value2) & "|" & CStr(target.Cells(t, 2).Value2) & "|" & UCase$(CStr(target.Cells(t, 3).Value2)) & _
            "|" & CStr(target.Cells(t, 4).Value2)) = True
    Next t
    Set rows = CreateObject("Scripting.Dictionary")
    For Each area In Selection.Areas
        For Each cell In area.Rows
            If cell.Row >= 2 Then rows(cell.Row) = True
        Next cell
    Next area
    For Each r In rows.Keys
        If Len(CStr(source.Cells(r, 1).Value2)) > 0 Then
            key = CStr(source.Cells(r, 2).Value2) & "|" & settingId & "|" & UCase$(CStr(source.Cells(r, 4).Value2)) & "|" & CStr(source.Cells(r, 5).Value2)
            If existing.Exists(key) Then
                skipped = skipped + 1
            Else
                t = modX13IO.NextRow(target)
                target.Cells(t, 1).Resize(1, 6).NumberFormat = "@"
                target.Cells(t, 1).Resize(1, 6).Value2 = Array(source.Cells(r, 2).Value2, settingId, UCase$(CStr(source.Cells(r, 4).Value2)), _
                    source.Cells(r, 5).Value2, "", "検出結果から転記（試行 " & source.Cells(r, 1).Value2 & "）")
                existing.Add key, True
                added = added + 1
            End If
        End If
    Next r
    target.Activate
    modUI.ShowMessage added & "件を「" & target.Name & "」へ書き写しました（既にあった" & skipped & "件は追加していません）。" & vbLf & _
        "次の実行で使うには、設定_系列別で設定ID「" & settingId & "」の outlier_mode と model_mode を選び、ControlのX13_SETTING_IDを「" & _
        settingId & "」にしてください。", vbInformation, "検出外れ値の転記"
    Exit Sub
Failed:
    modUI.ShowMessage Err.Description, vbExclamation, "検出外れ値の転記"
End Sub
