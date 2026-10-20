Attribute VB_Name = "modX13Options"
Option Explicit

' 役割      : 設定項目の一覧（キー、系列別列名、初期値、選択肢、説明）を一か所で定義し、
'             値を検査してX-13の正式な指定（SPCの各ブロック）に変換する。
' 主な入口  : Definitions / SetupControl / X11Block / AutoModelBlock / EstimateBlock / OutlierBlock
'             ForecastBlock / FirstMonth / LastMonth / ActualMonths / ForecastMonths
' 依存先    : modMain、modSeriesSettings（系列別に解決済みの設定 cfg を受け取る）、modX13IO
' 設定の優先順位（modSeriesSettings.ResolveAll が適用する）:
'   1) 設定_系列別 の「その系列ID × 設定ID」の行  2) 同じ設定IDの「*」行  3) Control  4) 下の初期値
'   系列別列名が空の項目は、全系列共通（Controlと初期値のみ）。
' 根拠      : Census X-13 Reference Manual §7.2 automdl（maxdiffの季節側は1のみ）、
'             §7.19 x11（trendmaは1より大きく101以下の奇数）、§7.7 forecast（maxleadは0～120、
'             0で予測なし）。
' 範囲外    : check・spectrum・history・slidingspans（事後診断）のブロックは作らない。

' 定義表の列位置（Definitions() の各要素は0始まりの配列）
Public Const DEF_KEY As Long = 0
Public Const DEF_SERIES_COLUMN As Long = 1
Public Const DEF_DEFAULT As Long = 2
Public Const DEF_CHOICES As Long = 3
Public Const DEF_NOTE As Long = 4

' 実績月数。本事業のIIPは連続144か月（12年）で固定する。Controlの値では変えられない。
' 開始年月を動かしても実績の長さは144か月のまま。予測月数はこの外側に別に数える。
Public Const ACTUAL_MONTHS_REQUIRED As Long = 144
' 予測月数の上限。r10から変更なし。
Private Const MAX_FORECAST_MONTHS As Long = 120

' 目的: 全設定項目の定義を返す。表示順もこの順。
' 出力: Collection。各要素は Array(キー, 系列別列名, 初期値, 選択肢, 説明)
Public Function Definitions() As Collection
    Dim d As New Collection
    ' ---- パッケージ内の場所（全系列共通）。相対パスはこのブックの保存先が基準 ----
    d.Add Array("X13_EXECUTABLE_PATH", "", "01_engine\x13as_ascii.exe", "", "X-13の実行ファイル。ASCII版（x13as_ascii.exe）・HTML版（x13as_html.exe）のどちらでもよい。版は実行後の出力ファイルで判定し、ファイル名では決めない。相対パスはこのブックのフォルダが基準")
    d.Add Array("X13_WORK_DIR", "", "04_results\work", "", "X-13の作業領域。実行ごとに新しいフォルダを作る。半角英数字だけのパスが必要（日本語を含む場所ではC:\IIP_X13_work等の絶対パスを指定）")
    d.Add Array("X13_CANDIDATE_DIR", "", "04_results\candidates", "", "比較候補（試行）の保存先。実行ID・系列ID別のフォルダに分ける")
    d.Add Array("X13_FINAL_DIR", "", "04_results\final", "", "最終決定したモデルの保存先。系列ID・保存ID別のフォルダに分け、過去の結果を上書きしない")
    d.Add Array("X13_SUMMARY_DIR", "", "04_results\summaries", "", "系列別の候補比較Excelの保存先")
    d.Add Array("INPUT_PUBLIC_EXCEL_DIR", "", "02_input\public_excel", "", "経産省の公表Excel原本の置き場所（読み取るだけで変更しない）")
    d.Add Array("INPUT_SERIES_DIR", "", "02_input\series", "", "系列ID別の原値CSV（raw.csv）の置き場所。ゼロ置換前の原値")
    d.Add Array("HOLIDAY_CSV_PATH", "", "03_holiday\source\cabinet_holidays.csv", "", "内閣府「国民の祝日」CSV")
    d.Add Array("HOLIDAY_DOWNLOAD_DIR", "", "03_holiday\source", "", "公式URLから取得したCSVの保存先")
    d.Add Array("HOLIDAY_GENERATED_DIR", "", "03_holiday\generated", "", "祝日変数（HOL）と月別の計算根拠の保存先")
    ' ---- 実行の許可と区切り（全系列共通） ----
    d.Add Array("ALLOW_X13_EXECUTION", "", "NO", "YES,NO", "X-13の外部実行を許可するか。内容を確認した利用者が手動でYESにする")
    d.Add Array("X13_BATCH_TIMEOUT_SECONDS", "", "600", "", "X-13の1回の起動を待つ上限秒数（1～3600）")
    d.Add Array("X13_COMPARE_CHUNK", "", "50", "", "1回の起動で計算する候補数（1～500）。系列数の上限ではない")
    ' ---- 期間と予測（全系列共通） ----
    d.Add Array("X13_START_YEAR", "", "2014", "", "実績期間の開始年")
    d.Add Array("X13_START_MONTH", "", "1", "", "実績期間の開始月")
    d.Add Array("X13_END_YEAR", "", "2025", "", "実績期間の終了年。開始から終了までが連続144か月でなければ停止（144か月は固定で変更不可）")
    d.Add Array("X13_END_MONTH", "", "12", "", "実績期間の終了月。開始年月を変えた場合は、144か月目になるよう終了年月も変える")
    d.Add Array("X13_GENERATE_FORECAST", "", "YES", "YES,NO", "予測を作るか。NOは予測月数0。予測は実績144か月の外側に付ける")
    d.Add Array("X13_FORECAST_MONTHS", "", "12", "", "予測月数（0～120）。実績月数には含めない")
    d.Add Array("X13_DAT_DECIMALS", "", "10", "", "DAT（計算用データ）に書く小数点以下の桁数（0～15）。原データは変更しない")
    d.Add Array("X13_OUT_DECIMALS", "", "3", "", "主出力（ASCII版 .out／HTML版 .html）の表に印字する小数桁（0～5）")
    d.Add Array("X13_SAVE_PRECISION", "", "15", "", "D10等の保存ファイルの有効桁（0～15）")
    ' ---- 対象と設定ID ----
    d.Add Array("X13_SERIES_SCOPE", "", "SELECTED", "SELECTED,ALL_INPUT", "SELECTED=設定_対象系列でYESの系列、ALL_INPUT=取り込んだ全系列")
    d.Add Array("X13_SETTING_ID", "", "standard", "", "今回の実行で使う設定ID。設定_系列別・設定_外れ値指定のsetting_id列と照合する")
    ' ---- 変換とゼロ置換（系列別に変更可） ----
    d.Add Array("X13_TRANSFORM_FUNCTION", "transform", "log", "log,none,auto", "対数変換。log=あり、none=なし、auto=X-13が選ぶ（ゼロ置換はlogと同じ扱い）。autoのときはX-13の仕様によりx11のmodeを書かず、季節調整の型もX-13が決める")
    d.Add Array("X13_ZERO_REPLACE", "zero_replace", "YES", "YES,NO", "対数計算のため、数値の0だけを計算用データでepsilonに置き換えるか（原データは変更しない）")
    d.Add Array("X13_ZERO_EPSILON", "zero_epsilon", "0.000001", "", "0の置換値。0より大きい数値")
    ' ---- X-11（系列別に変更可） ----
    d.Add Array("X13_X11_MODE", "x11_mode", "mult", "mult,add", "季節調整の型。mult=乗法型、add=加法型。transform=autoの系列では使わない（X-13がmodeの指定を受け付けないため）")
    d.Add Array("X13_SEASONAL_MA", "seasonal_ma", "x11default", "x11default,msr,s3x1,s3x3,s3x5,s3x9,s3x15,stable", "季節移動平均")
    d.Add Array("X13_TREND_MA", "trend_ma", "", "", "傾向移動平均（Henderson）。空欄はX-13が自動選択。指定は3～101の奇数")
    d.Add Array("X13_SIGMA_LOWER", "sigma_lower", "1.5", "", "管理限界値の下限（sigmalimの1番目）")
    d.Add Array("X13_SIGMA_UPPER", "sigma_upper", "2.5", "", "管理限界値の上限（sigmalimの2番目）。下限より大きい値")
    ' ---- ARIMA（系列別に変更可） ----
    d.Add Array("X13_MODEL_MODE", "model_mode", "自動選択", "自動選択,次数指定", "ARIMA次数の決め方。自動選択=automdl、次数指定=下のARIMA次数")
    d.Add Array("X13_ARIMA_MODEL", "arima_model", "(0 1 1)(0 1 1)", "", "次数指定の場合の次数 (p d q)(P D Q)。エアラインモデルは(0 1 1)(0 1 1)")
    d.Add Array("X13_MAXDIFF_REGULAR", "maxdiff_regular", "1", "1,2", "自動選択の非季節差分の上限（automdl.maxdiffの1番目）")
    d.Add Array("X13_MAXDIFF_SEASONAL", "maxdiff_seasonal", "1", "1", "自動選択の季節差分の上限（automdl.maxdiffの2番目）。マニュアル§7.2により1のみ")
    d.Add Array("X13_MAXORDER_REGULAR", "maxorder_regular", "2", "1,2,3,4", "自動選択の非季節ARMA次数の上限")
    d.Add Array("X13_MAXORDER_SEASONAL", "maxorder_seasonal", "2", "1,2", "自動選択の季節ARMA次数の上限")
    d.Add Array("X13_ESTIMATE_MAXITER", "estimate_maxiter", "500", "", "ARMA推定の最大反復回数（1～100000）")
    d.Add Array("X13_MODEL_SPAN", "model_span", "", "", "係数の推定に使う期間（series.modelspan）。空欄は全期間。例 2014.1,2025.12")
    ' ---- 外れ値（系列別に変更可） ----
    d.Add Array("X13_OUTLIER_MODE", "outlier_mode", "自動", "自動,自動＋指定,指定のみ", "外れ値の扱い。指定は設定_外れ値指定の表を使う")
    d.Add Array("X13_OUTLIER_TYPES", "outlier_types", "ao ls tc", "", "自動検出する種類（ao ls tc の組合せ、またはall）")
    d.Add Array("X13_OUTLIER_CRITICAL", "outlier_critical", "", "", "自動検出の臨界値。空欄はX-13の既定")
    d.Add Array("X13_OUTLIER_SPAN", "outlier_span", "", "", "自動検出する期間。空欄は全期間。例 2014.1,2025.12")
    ' ---- カレンダー（系列別に変更可） ----
    d.Add Array("X13_CALENDAR_MODE", "calendar_mode", "指定のみ", "自動,自動＋指定,指定のみ,なし", "曜日・祝日調整の決め方。自動は候補変数の有無の組合せを実際に推定して比較")
    d.Add Array("X13_CALENDAR_POOL", "calendar_pool", "tdnolpyear td1nolpyear lpyear hol", "", "自動で比較する変数。tdnolpyearとtd1nolpyearは代替候補で、同じ候補に同時には入れない")
    d.Add Array("X13_FORCED_CALENDAR", "forced_calendar", "", "", "自動＋指定・指定のみで必ず使う変数。例 td1nolpyear hol。空欄なら下の3項目から作る")
    d.Add Array("X13_TD_MODE", "td_mode", "OFF", "OFF,2曜日,7曜日", "指定のみ：曜日調整。2曜日=td1nolpyear、7曜日=tdnolpyear")
    d.Add Array("X13_LEAP_YEAR_MODE", "leap_year_mode", "OFF", "OFF,ON", "指定のみ：うるう年調整（lpyear）")
    d.Add Array("X13_HOLIDAY_ADJUST", "holiday_adjust", "OFF", "OFF,ON", "指定のみ：祝日調整（hol）")
    ' ---- 比較と最終決定（全系列共通） ----
    d.Add Array("X13_REFIT_SEARCH_MODELS", "", "YES", "YES,NO", "自動選択の探索で現れたARIMA次数を、初回の外れ値の月と種類を固定して各カレンダー条件で再推定するか")
    d.Add Array("X13_RANK_METRIC", "", "AICC", "AIC,BIC,AICC,HQ", "順位に使う統計量（小さいほど上位）")
    d.Add Array("X13_CALENDAR_ALPHA", "", "0.05", "", "変数検定の有意水準（0より大きく1より小さい）")
    d.Add Array("X13_REQUIRE_CALENDAR_SIGNIFICANCE", "", "NO", "YES,NO", "YESは曜日・うるう年・祝日の結合検定がすべて有意な候補だけを自動決定の対象にする")
    d.Add Array("X13_FINAL_MODE", "", "手動", "手動,自動", "手動=利用者が試行IDを指定、自動=比較群内で条件を満たす最上位")
    d.Add Array("X13_COMPARISON_GROUP", "", "AUTO_REFERENCE", "", "自動決定に使う比較群。AUTO_REFERENCE=系列ごとの基準試行（初回で最初に成立した試行）と同じ群")
    ' ---- 祝日変数（全系列共通）。内閣府CSVから経産省IIP方式の1列を作る ----
    d.Add Array("HOLIDAY_CSV_ENCODING", "", "AUTO", "AUTO,CP932,UTF8", "CSVの文字コード。AUTOはUTF-8（BOM有無）とCP932を判別")
    d.Add Array("HOLIDAY_CSV_ORIGIN", "", "", "", "CSVの取得元。空欄なら同じフォルダの取得記録（provenance.json）を使う。どちらも無ければ未記入として記録")
    d.Add Array("HOLIDAY_CSV_ACQUIRED_AT", "", "", "", "CSVの取得日時（UTC）。空欄なら同じフォルダの取得記録を使う")
    Set Definitions = d
End Function

' 目的  : 定義表からキーの定義を探す。無ければ停止（未定義のキーを読まないため）。
Public Function DefinitionOf(ByVal key As String) As Variant
    Dim item As Variant
    For Each item In Definitions()
        If item(DEF_KEY) = key Then
            DefinitionOf = item
            Exit Function
        End If
    Next item
    Err.Raise vbObjectError + 2500, "DefinitionOf", "未定義の設定キーです：" & key
End Function

' 目的  : Controlに全設定の行を作る。既にある行の値は変更しない。
' 副作用: Controlシートへ行・入力規則を追加する。
Public Sub SetupControl()
    Dim ws As Worksheet, item As Variant, row As Long
    Set ws = modMain.EnsureControlSheet()
    For Each item In Definitions()
        row = modMain.ControlRow(CStr(item(DEF_KEY)))
        If row = 0 Then
            row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(row, 1).Value2 = item(DEF_KEY)
            ws.Cells(row, 2).NumberFormat = "@"
            ws.Cells(row, 2).Value2 = item(DEF_DEFAULT)
        End If
        ws.Cells(row, 3).Value2 = item(DEF_NOTE) & IIf(Len(item(DEF_SERIES_COLUMN)) > 0, _
            "（系列別に変更可：設定_系列別の " & item(DEF_SERIES_COLUMN) & " 列）", "")
        If Len(item(DEF_CHOICES)) > 0 Then AddChoice ws.Cells(row, 2), CStr(item(DEF_CHOICES))
    Next item
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 100
    ws.Columns("C").WrapText = True
End Sub

' 目的: セルに一覧選択の入力規則を付ける。
Public Sub AddChoice(ByVal target As Range, ByVal choices As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=choices
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorMessage = "一覧から選択してください。"
    End With
End Sub

' ---- 全系列共通の数値設定 ----

' 目的  : Controlの整数設定を読み、範囲を検査する。
' エラー: 整数でない、範囲外の場合は停止。
Public Function IntegerSetting(ByVal key As String, ByVal minimum As Long, ByVal maximum As Long) As Long
    IntegerSetting = CheckedInteger(modMain.GetControlValue(key, CStr(DefinitionOf(key)(DEF_DEFAULT))), key, minimum, maximum)
End Function

Public Function CheckedInteger(ByVal text As String, ByVal key As String, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim value As Double
    If Not modX13IO.IsNumberText(text) Then Err.Raise vbObjectError + 2501, "CheckedInteger", key & "は整数で指定してください：" & text
    value = modX13IO.NumberValue(text)
    If value <> Fix(value) Or value < minimum Or value > maximum Then _
        Err.Raise vbObjectError + 2502, "CheckedInteger", key & "は" & minimum & "～" & maximum & "の整数で指定してください：" & text
    CheckedInteger = CLng(value)
End Function

Public Function CheckedNumber(ByVal text As String, ByVal key As String) As Double
    If Not modX13IO.IsNumberText(text) Then Err.Raise vbObjectError + 2503, "CheckedNumber", key & "は数値で指定してください：" & text
    CheckedNumber = modX13IO.NumberValue(text)
End Function

' 実績期間の最初の月番号。
Public Function FirstMonth() As Long
    FirstMonth = modX13IO.MonthIndex(IntegerSetting("X13_START_YEAR", 1900, 2999), IntegerSetting("X13_START_MONTH", 1, 12))
End Function

' 実績期間の最後の月番号。
Public Function LastMonth() As Long
    LastMonth = modX13IO.MonthIndex(IntegerSetting("X13_END_YEAR", 1900, 2999), IntegerSetting("X13_END_MONTH", 1, 12))
End Function

' 目的  : 実績月数（開始～終了の月数）を返す。値は常に ACTUAL_MONTHS_REQUIRED（144）。
' 規則  : 開始・終了年月から数えた月数が144でなければ停止する。月数の検査値を設定で変える項目は設けない
'         （IIPは144か月固定。空欄や別の値で検査を外せないようにするため）。
' エラー: 144か月でない場合は、開始年月から144か月目の終了年月を示して停止する。
Public Function ActualMonths() As Long
    Dim first As Long, last As Long, count As Long
    first = FirstMonth()
    last = LastMonth()
    count = last - first + 1
    If count <> ACTUAL_MONTHS_REQUIRED Then Err.Raise vbObjectError + 2504, "ActualMonths", _
        "実績期間は連続" & ACTUAL_MONTHS_REQUIRED & "か月で固定です。現在の" & modX13IO.MonthText(first) & "～" & _
        modX13IO.MonthText(last) & "は" & count & "か月です。開始を" & modX13IO.MonthText(first) & "とする場合、終了は" & _
        modX13IO.MonthText(first + ACTUAL_MONTHS_REQUIRED - 1) & "です（予測月数は別に数えます）。"
    ActualMonths = count
End Function

' 予測月数。予測を作らない設定なら0。
Public Function ForecastMonths() As Long
    If Not modMain.GetControlBoolean("X13_GENERATE_FORECAST", True) Then Exit Function
    ForecastMonths = IntegerSetting("X13_FORECAST_MONTHS", 0, MAX_FORECAST_MONTHS)
End Function

Public Function PeriodLabel() As String
    PeriodLabel = modX13IO.MonthText(FirstMonth()) & "～" & modX13IO.MonthText(LastMonth()) & "（" & ActualMonths() & "か月）"
End Function

Public Function DatDecimals() As Long
    DatDecimals = IntegerSetting("X13_DAT_DECIMALS", 0, 15)
End Function

' ---- 系列別に解決済みの設定（cfg）から作るSPCの各ブロック ----

' 目的: 変換方式（log/none/auto）。
Public Function TransformFunction(ByVal cfg As Object) As String
    TransformFunction = LCase$(modSeriesSettings.SettingText(cfg, "X13_TRANSFORM_FUNCTION"))
    Select Case TransformFunction
        Case "log", "none", "auto"
        Case Else: Err.Raise vbObjectError + 2506, "TransformFunction", "変換方式はlog・none・autoのいずれかです：" & TransformFunction
    End Select
End Function

' 目的: transformブロック。constantやtrimzeroは使わない（ゼロはmodZeroReplaceで扱う）。
Public Function TransformBlock(ByVal cfg As Object) As String
    TransformBlock = "transform{function=" & TransformFunction(cfg) & "}"
End Function

' 目的  : x11ブロック。乗法/加法、季節・傾向移動平均、管理限界値、保存表、予測の付加。
' 規則  : transform=auto の系列では mode を書かない。X-13は自動の変換選択と mode の同時指定を受け付けず、
'         エラー記録に「Cannot set seasonal adjustment mode when automatic transformation selection is done.」と書いて、
'         終了コード0のまま計算しない（公式Build62の実出力：検証/20261020_HTML対応/reference/probe_failure）。
'         このときの型はX-13に任せる。実際に使われた型はUDGの samode に記録され、比較群の条件とD18の中立値に使う。
'         X13_X11_MODE の値の検査は、autoの系列でも従来どおり行う。
' エラー: 不正な指定は停止。trendmaは3～101の奇数（マニュアル§7.19）。
Public Function X11Block(ByVal cfg As Object, ByVal forecastMonthCount As Long) As String
    Dim mode As String, seasonalMa As String, trendText As String, trendPart As String, modePart As String
    Dim lower As Double, upper As Double, trendLength As Long

    mode = LCase$(modSeriesSettings.SettingText(cfg, "X13_X11_MODE"))
    If mode <> "mult" And mode <> "add" Then Err.Raise vbObjectError + 2507, "X11Block", "季節調整の型はmultまたはaddです：" & mode
    modePart = "mode=" & mode & " "
    If TransformFunction(cfg) = "auto" Then modePart = ""
    seasonalMa = LCase$(modSeriesSettings.SettingText(cfg, "X13_SEASONAL_MA"))
    If InStr("," & DefinitionOf("X13_SEASONAL_MA")(DEF_CHOICES) & ",", "," & seasonalMa & ",") = 0 Then _
        Err.Raise vbObjectError + 2508, "X11Block", "季節移動平均の指定が不正です：" & seasonalMa
    lower = CheckedNumber(modSeriesSettings.SettingText(cfg, "X13_SIGMA_LOWER"), "X13_SIGMA_LOWER")
    upper = CheckedNumber(modSeriesSettings.SettingText(cfg, "X13_SIGMA_UPPER"), "X13_SIGMA_UPPER")
    If lower <= 0 Or upper <= lower Then Err.Raise vbObjectError + 2509, "X11Block", "管理限界値は 0 ＜ 下限 ＜ 上限 で指定してください"
    trendText = modSeriesSettings.SettingText(cfg, "X13_TREND_MA")
    If Len(trendText) > 0 Then
        trendLength = CheckedInteger(trendText, "X13_TREND_MA", 3, 101)
        If trendLength Mod 2 = 0 Then Err.Raise vbObjectError + 2510, "X11Block", "傾向移動平均は3～101の奇数で指定してください：" & trendText
        trendPart = " trendma=" & trendLength
    End If
    ' 2行に分ける（SPCの1行132文字の上限に余裕を持たせるため）。
    X11Block = "x11{" & modePart & "seasonalma=" & seasonalMa & trendPart & _
        " sigmalim=(" & modX13IO.NumberText(lower) & " " & modX13IO.NumberText(upper) & ")" & vbCrLf & _
        " save=(d10 d11 d12 d13 d16 d18) appendfcst=" & IIf(forecastMonthCount > 0, "yes", "no") & "}"
End Function

' 目的  : automdlブロック。maxdiffの季節側は1のみ（マニュアル§7.2）。
Public Function AutoModelBlock(ByVal cfg As Object) As String
    AutoModelBlock = "automdl{maxdiff=(" & _
        CheckedInteger(modSeriesSettings.SettingText(cfg, "X13_MAXDIFF_REGULAR"), "X13_MAXDIFF_REGULAR", 1, 2) & " " & _
        CheckedInteger(modSeriesSettings.SettingText(cfg, "X13_MAXDIFF_SEASONAL"), "X13_MAXDIFF_SEASONAL", 1, 1) & ") maxorder=(" & _
        CheckedInteger(modSeriesSettings.SettingText(cfg, "X13_MAXORDER_REGULAR"), "X13_MAXORDER_REGULAR", 1, 4) & " " & _
        CheckedInteger(modSeriesSettings.SettingText(cfg, "X13_MAXORDER_SEASONAL"), "X13_MAXORDER_SEASONAL", 1, 2) & _
        ") print=all savelog=alldiagnostics}"
End Function

' 目的: ARIMA次数 "(p d q)(P D Q)" を検査して空白を整える。
Public Function ArimaModelText(ByVal model As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\s+"
    model = Trim$(re.Replace(model, " "))
    re.Global = False
    re.Pattern = "^\(\s?[0-9] [0-2] [0-9]\s?\)\s?\(\s?[0-9] [0-1] [0-9]\s?\)$"
    If Not re.Test(model) Then Err.Raise vbObjectError + 2511, "ArimaModelText", "ARIMA次数は (p d q)(P D Q) の形で指定してください：" & model
    ArimaModelText = Replace(Replace(Replace(model, "( ", "("), " )", ")"), ") (", ")(")
End Function

' 目的: ARIMAの決め方（automdl/fixed）。
Public Function ModelMode(ByVal rawValue As String) As String
    Select Case LCase$(Trim$(rawValue))
        Case "automdl", "自動選択": ModelMode = "automdl"
        Case "fixed", "次数指定": ModelMode = "fixed"
        Case Else: Err.Raise vbObjectError + 2512, "ModelMode", "ARIMAの決め方は自動選択または次数指定です：" & rawValue
    End Select
End Function

' 目的: estimateブロック。回帰係数の共分散（rcm）を保存し、情報量と標本内予測誤差をログに出す。
Public Function EstimateBlock(ByVal cfg As Object) As String
    EstimateBlock = "estimate{maxiter=" & _
        CheckedInteger(modSeriesSettings.SettingText(cfg, "X13_ESTIMATE_MAXITER"), "X13_ESTIMATE_MAXITER", 1, 100000) & _
        " save=(regcmatrix) print=(model estimates lkstats) savelog=(aic aicc bic hq afc)}"
End Function

' 目的  : forecastブロック。予測なしはmaxlead=0（マニュアル§7.7：forecastを省くと1年分の予測が作られるため明示する）。
' エラー: 0～120以外は停止。
Public Function ForecastBlock(ByVal forecastMonthCount As Long) As String
    If forecastMonthCount < 0 Or forecastMonthCount > MAX_FORECAST_MONTHS Then _
        Err.Raise vbObjectError + 2518, "ForecastBlock", "予測月数は0～" & MAX_FORECAST_MONTHS & "です：" & forecastMonthCount
    ForecastBlock = "forecast{maxlead=" & forecastMonthCount & "}"
End Function

' 目的: 外れ値の扱い（自動/自動＋指定/指定のみ）を検査する。
Public Function OutlierMode(ByVal rawValue As String) As String
    Select Case Trim$(rawValue)
        Case "自動", "自動＋指定", "指定のみ": OutlierMode = Trim$(rawValue)
        Case Else: Err.Raise vbObjectError + 2513, "OutlierMode", "外れ値の扱いは自動・自動＋指定・指定のみです：" & rawValue
    End Select
End Function

' 目的  : outlierブロック。指定のみの場合は空文字（自動検出しない）。
' 入力  : outlierModeText は候補ごとの扱い（再推定の候補は「指定のみ」）。
Public Function OutlierBlock(ByVal cfg As Object, ByVal outlierModeText As String) As String
    Dim typesText As String, item As Variant, criticalText As String, spanText As String
    If OutlierMode(outlierModeText) = "指定のみ" Then Exit Function
    typesText = LCase$(Trim$(modSeriesSettings.SettingText(cfg, "X13_OUTLIER_TYPES")))
    For Each item In Split(typesText, " ")
        Select Case CStr(item)
            Case "ao", "ls", "tc", "all", ""
            Case Else: Err.Raise vbObjectError + 2514, "OutlierBlock", "自動検出の種類は ao・ls・tc・all です：" & typesText
        End Select
    Next item
    If Len(typesText) = 0 Then Err.Raise vbObjectError + 2515, "OutlierBlock", "自動検出の種類を指定してください"
    OutlierBlock = "outlier{types=(" & typesText & ")"
    criticalText = modSeriesSettings.SettingText(cfg, "X13_OUTLIER_CRITICAL")
    If Len(criticalText) > 0 Then
        If CheckedNumber(criticalText, "X13_OUTLIER_CRITICAL") <= 0 Then Err.Raise vbObjectError + 2516, "OutlierBlock", "臨界値は正の数です"
        OutlierBlock = OutlierBlock & " critical=" & modX13IO.NumberText(modX13IO.NumberValue(criticalText))
    End If
    spanText = modSeriesSettings.SettingText(cfg, "X13_OUTLIER_SPAN")
    If Len(spanText) > 0 Then OutlierBlock = OutlierBlock & " span=(" & SafeSpecText(spanText) & ")"
    OutlierBlock = OutlierBlock & "}"
End Function

' 目的: SPCへ書く利用者入力に、改行・波括弧・引用符が無いことを確かめる。
Public Function SafeSpecText(ByVal text As String) As String
    If InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Or InStr(text, "{") > 0 Or _
       InStr(text, "}") > 0 Or InStr(text, Chr$(34)) > 0 Then _
        Err.Raise vbObjectError + 2517, "SafeSpecText", "SPCに書けない文字（改行・波括弧・引用符）があります：" & text
    SafeSpecText = text
End Function

' 目的: series.modelspan の指定（空欄なら空文字）。
Public Function ModelSpanText(ByVal cfg As Object) As String
    ModelSpanText = SafeSpecText(modSeriesSettings.SettingText(cfg, "X13_MODEL_SPAN"))
End Function
