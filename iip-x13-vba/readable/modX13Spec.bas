Attribute VB_Name = "modX13Spec"
Option Explicit

' 役割      : 1つのジョブ（1系列×1候補）について、X-13へ渡すSPC・DATと、追跡用の記録ファイルを作る。
' 主な入口  : SpecText（SPC本文を作る純粋な処理）/ WriteJobFiles（ジョブフォルダへ書く）
' 依存先    : modX13Options（各ブロック）、modX13Calendar、modHolidayIIP、modZeroReplace、modSeriesSettings、
'             modX13IO、modMain、modPackagePaths
' SPCの形   : series → transform → regression（カレンダー変数・指定外れ値・祝日変数）→ outlier（自動検出）
'             → automdl または arima → estimate → forecast → x11
'             check・spectrum・history・slidingspans は書かない（事後診断は今回の範囲外）。
' ファイル  : <stem>.spc / <stem>.dat（計算用の値。ゼロ置換後）/ _input_values.tsv（原値と元セル）
'             _zero_replacement.tsv / _used_settings.tsv / _candidate.txt / 祝日を使う候補は _holiday.*
' パスの書き方: X-13を「-m jobs」で一括実行すると、SPC内のファイル名はバッチフォルダからの相対位置で解釈される。
'               そのため file="j000001/j000001.dat" のようにジョブフォルダ名を付ける（r10で確認した挙動）。
' 行の長さ  : X-13は1行が132文字を超えるSPCを読めない（133文字の行で「Input record longer than limit」と
'             停止し、終了コードは0のままだった：検証/X13設定検証_20261016T094447Z）。
'             回帰変数の並びは WRAP_COLUMN 文字ごとに空白の位置で改行し、書く前に全行の長さを検査する。

' X-13が読めるSPCの1行の最大文字数と、回帰変数を折り返す目安の文字数。
Public Const MAX_SPC_LINE As Long = 132
Private Const WRAP_COLUMN As Long = 100

' 目的  : SPCの本文を作る。ファイルは書かない。
' 入力  : stem（ジョブ名 j000001）、firstMonth、forecastMonths、cfg（系列別設定）、calendar（"tdnolpyear hol"等）、
'         modelMode（automdl／fixed）、arimaModel（fixedの場合の次数）、outlierMode（自動／自動＋指定／指定のみ）、
'         regressors（指定外れ値・固定外れ値の変数名。空白区切り）
' 出力  : SPCの本文（ASCIIだけ）
' エラー: 設定値の不正は各ブロックの作成処理が停止する。
Public Function SpecText(ByVal stem As String, ByVal firstMonth As Long, ByVal forecastMonths As Long, ByVal cfg As Object, _
                         ByVal calendar As String, ByVal modelMode As String, ByVal arimaModel As String, _
                         ByVal outlierMode As String, ByVal regressors As String) As String
    Dim lines As String, variables As String, modelSpan As String, usesHol As Boolean, outlierText As String
    modX13IO.RequireIdentity stem
    variables = Trim$(modX13Calendar.RegressionVariables(calendar) & " " & modX13Options.SafeSpecText(regressors))
    usesHol = modX13Calendar.UsesHoliday(calendar)
    modelSpan = modX13Options.ModelSpanText(cfg)

    lines = "series{" & vbCrLf & _
        " title=""" & stem & """" & vbCrLf & _
        " file=""" & stem & "/" & stem & ".dat""" & vbCrLf & _
        " format=""free""" & vbCrLf & _
        " start=" & modX13IO.X13StartText(firstMonth) & vbCrLf & _
        " period=12" & vbCrLf & _
        " decimals=" & modX13Options.IntegerSetting("X13_OUT_DECIMALS", 0, 5) & vbCrLf & _
        " saveprecision=" & modX13Options.IntegerSetting("X13_SAVE_PRECISION", 0, 15) & vbCrLf
    If Len(modelSpan) > 0 Then lines = lines & " modelspan=(" & modelSpan & ")" & vbCrLf
    lines = lines & "}" & vbCrLf
    lines = lines & modX13Options.TransformBlock(cfg) & vbCrLf
    If Len(variables) > 0 Or usesHol Then
        lines = lines & "regression{" & vbCrLf
        If Len(variables) > 0 Then lines = lines & WrappedList(" variables=(", variables, ")") & vbCrLf
        If usesHol Then lines = lines & modHolidayIIP.RegressionLines(firstMonth, stem & "/" & modHolidayIIP.HOL_FILE_NAME) & vbCrLf
        lines = lines & "}" & vbCrLf
    End If
    outlierText = modX13Options.OutlierBlock(cfg, outlierMode)
    If Len(outlierText) > 0 Then lines = lines & outlierText & vbCrLf
    Select Case modelMode
        Case "automdl": lines = lines & modX13Options.AutoModelBlock(cfg) & vbCrLf
        Case "fixed": lines = lines & "arima{model=" & modX13Options.ArimaModelText(arimaModel) & "}" & vbCrLf
        Case Else: Err.Raise vbObjectError + 3200, "SpecText", "ARIMAの決め方が不正です：" & modelMode
    End Select
    lines = lines & modX13Options.EstimateBlock(cfg) & vbCrLf
    lines = lines & modX13Options.ForecastBlock(forecastMonths) & vbCrLf
    lines = lines & modX13Options.X11Block(cfg, forecastMonths) & vbCrLf
    RequireLineLength lines
    SpecText = lines
End Function

' 目的  : 空白区切りの並びを、1行が WRAP_COLUMN 文字を超えないよう空白の位置で改行する。
' 例    : WrappedList(" variables=(", "tl2022.jan-2022.mar ... qd2019.jan-2019.jun", ")")
'           →  " variables=(tl2022.jan-2022.mar tl2024.may-2024.nov" & 改行 & " so2019.apr ...)"
'         続きの行の先頭は空白1つ（X-13の検証で読めた形：検証/X13設定検証_20261016T094557Z/specified_fixed）。
Private Function WrappedList(ByVal opening As String, ByVal items As String, ByVal closing As String) As String
    Dim token As Variant, current As String, result As String
    current = opening
    For Each token In Split(Trim$(items), " ")
        If Len(CStr(token)) > 0 Then
            If Len(current) > Len(opening) And Len(current) + 1 + Len(CStr(token)) > WRAP_COLUMN Then
                result = result & current & vbCrLf
                current = " " & CStr(token)
            ElseIf Len(current) = Len(opening) Then
                current = current & CStr(token)
            Else
                current = current & " " & CStr(token)
            End If
        End If
    Next token
    WrappedList = result & current & closing
End Function

' 目的  : SPCのすべての行が MAX_SPC_LINE 文字以下か確かめる。
' エラー: 超える行があれば、行番号と長さを示して停止する（X-13を起動する前に止める）。
Public Sub RequireLineLength(ByVal specBody As String)
    Dim parts As Variant, i As Long
    parts = Split(Replace(specBody, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(parts)
        If Len(parts(i)) > MAX_SPC_LINE Then Err.Raise vbObjectError + 3202, "RequireLineLength", _
            "SPCの" & i + 1 & "行目が" & Len(parts(i)) & "文字です。X-13が読める" & MAX_SPC_LINE & "文字を超えています：" & Left$(parts(i), 60) & "…"
    Next i
End Sub

' 目的  : ジョブフォルダへSPC・DATと記録ファイルを書く。
' 入力  : job（候補の定義。modX13Runが作るDictionary）、seriesContext（系列の設定・原値・計算用の値）、
'         holidayContext（祝日変数の計算結果。祝日を使わない実行ではNothingでよい）、jobFolder、stem
' 出力  : SPCの絶対パス
' 副作用: jobFolderへ新しいファイルを書く（既存ファイルがあれば停止）。
' エラー: 祝日を使う候補なのに祝日変数が無い場合、SPCにASCII以外が入った場合は停止する。
Public Function WriteJobFiles(ByVal job As Object, ByVal seriesContext As Object, ByVal holidayContext As Object, _
                              ByVal jobFolder As String, ByVal stem As String) As String
    Dim prepared As Object, cfg As Object, specPath As String, i As Long, body As String
    Set prepared = seriesContext("prepared")
    Set cfg = seriesContext("cfg")
    body = ""
    For i = 1 To UBound(prepared("datText"))
        body = body & prepared("datText")(i) & vbCrLf
    Next i
    modX13IO.WriteNewText modMain.CombinePath(jobFolder, stem & ".dat"), body
    If modX13Calendar.UsesHoliday(CStr(job("calendar"))) Then
        If holidayContext Is Nothing Then Err.Raise vbObjectError + 3201, "WriteJobFiles", "祝日変数を使う候補ですが、祝日変数が準備されていません。"
        modHolidayIIP.WriteJobHolidayFiles holidayContext, jobFolder
    End If
    specPath = modMain.CombinePath(jobFolder, stem & ".spc")
    modX13IO.WriteNewText specPath, SpecText(stem, seriesContext("first"), seriesContext("forecast"), cfg, CStr(job("calendar")), _
        CStr(job("modelMode")), CStr(job("arima")), CStr(job("outlierMode")), CStr(job("regressors")))
    modX13IO.RequireAsciiFile specPath
    modX13IO.WriteNewText modMain.CombinePath(jobFolder, "_input_values.tsv"), InputValuesText(seriesContext)
    modX13IO.WriteNewText modMain.CombinePath(jobFolder, "_zero_replacement.tsv"), _
        modZeroReplace.ReplacementFileText(CStr(job("series")), CStr(job("setting")), prepared)
    modX13IO.WriteNewText modMain.CombinePath(jobFolder, "_used_settings.tsv"), modSeriesSettings.UsedSettingsText(cfg)
    modX13IO.WriteNewText modMain.CombinePath(jobFolder, "_candidate.txt"), CandidateText(job, seriesContext)
    WriteJobFiles = specPath
End Function

' 目的  : 原値と計算用の値を月ごとに並べた表（タブ区切り）。原値は原データ_原指数の値そのもの。
' 規則  : i行目は実績期間のi番目の月。sources(i)の年月がその月であること、sources(i)の原値と
'         prepared("original")(i)が同じであることを確かめてから書く（位置だけで結ばない）。
' エラー: 年月または原値が対応しない場合は停止する。
Public Function InputValuesText(ByVal seriesContext As Object) As String
    Dim prepared As Object, sources As Collection, body As String, i As Long, item As Variant, expected As String
    Set prepared = seriesContext("prepared")
    Set sources = seriesContext("sources")
    If sources.Count <> UBound(prepared("original")) Then Err.Raise vbObjectError + 3203, "InputValuesText", _
        "元セルの件数（" & sources.Count & "）と計算用の値の件数（" & UBound(prepared("original")) & "）が違います。"
    body = "date" & vbTab & "original_value" & vbTab & "computed_value" & vbTab & "dat_text" & vbTab & _
        "preliminary" & vbTab & "source_file" & vbTab & "source_sheet" & vbTab & "source_cell" & vbTab & "raw_sheet_row" & vbCrLf
    For i = 1 To sources.Count
        item = sources(i)
        expected = modX13IO.MonthText(seriesContext("first") + i - 1)
        If CStr(item(0)) <> expected Then Err.Raise vbObjectError + 3204, "InputValuesText", _
            i & "番目の元セルの年月が" & item(0) & "です（期待 " & expected & "）。"
        If CDbl(item(1)) <> prepared("original")(i) Then Err.Raise vbObjectError + 3205, "InputValuesText", _
            expected & "の原値が元セルと計算用の値で一致しません。"
        body = body & expected & vbTab & modX13IO.NumberText(prepared("original")(i)) & vbTab & _
            modX13IO.NumberText(prepared("computed")(i)) & vbTab & prepared("datText")(i) & vbTab & _
            item(2) & vbTab & item(3) & vbTab & item(4) & vbTab & item(5) & vbTab & item(6) & vbCrLf
    Next i
    InputValuesText = body
End Function

' 候補の定義（キー=値）。保存時の照合と、人が条件を読むために使う。
Private Function CandidateText(ByVal job As Object, ByVal seriesContext As Object) As String
    CandidateText = "RUN=" & job("run") & vbCrLf & _
        "CANDIDATE=" & job("candidate") & vbCrLf & _
        "SERIES=" & job("series") & vbCrLf & _
        "SETTING=" & job("setting") & vbCrLf & _
        "STAGE=" & job("stage") & vbCrLf & _
        "CALENDAR=" & job("calendar") & vbCrLf & _
        "MODEL_MODE=" & job("modelMode") & vbCrLf & _
        "ARIMA=" & job("arima") & vbCrLf & _
        "OUTLIER_MODE=" & job("outlierMode") & vbCrLf & _
        "REGRESSORS=" & job("regressors") & vbCrLf & _
        "REFERENCE_TRIAL=" & job("reference") & vbCrLf & _
        "ORIGIN=" & job("origin") & vbCrLf & _
        "FIRST=" & modX13IO.MonthText(seriesContext("first")) & vbCrLf & _
        "ACTUAL_MONTHS=" & seriesContext("actual") & vbCrLf & _
        "FORECAST_MONTHS=" & seriesContext("forecast") & vbCrLf
End Function
