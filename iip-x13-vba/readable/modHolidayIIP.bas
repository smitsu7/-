Attribute VB_Name = "modHolidayIIP"
Option Explicit

' 役割      : 内閣府「国民の祝日」CSVから、経産省がIIPで公表している方法で祝日変数（1列）を作る。
'             根拠：経済産業省「2020年基準改定の概要」PDF34頁（印刷32頁）「４．祝日変数について」
'               A = 各年各月で、月曜日～金曜日に当たる祝日・振替休日・国民の休日の日数
'                   （CSVの行をそのまま数える。土日・独自の休業日は加えない）
'               B = 実績期間の同じ暦月に当たるAの平均（144か月なら12個、96か月なら8個の平均。
'                   期間が年の途中で始まる・終わる場合は、暦月ごとに実績期間にある個数で割る）
'               祝日変数 = A - B。予測月のBも実績期間だけの平均（予測年を平均に入れない）。
'             2019年の特例や2020・2021年の祝日移動は、CSVに載っている日付をそのまま使う。
'             祝日の規則（第2月曜日等）は実装しない。
' 主な入口  : DownloadOfficialCsv / ChooseLocalCsv / BuildFromUI（画面の操作）
'             PrepareForRun（実行前の検査と計算）/ WriteJobHolidayFiles / RegressionLines / VerifySnapshot
' 依存先    : modIIPTextTable（文字コード）、modHash（SHA-256）、modX13IO、modMain、modSheetNames、
'             modX13Options、modPackagePaths、modUI（画面の通知）、modLog
' 書く場所  : 原データ_祝日CSV（CSVの日付一覧）、計算用_祝日変数（年月・A・B・A-B）、
'             03_holiday\generated（HOLと計算根拠）、各ジョブフォルダの _holiday.hol と記録ファイル
' 計算例    : 同梱の内閣府CSVで2014～2025年を実績とした場合（tests\fixtures\expected_holiday_2014_2026.tsv と同じ値）
'             1月のAは2022年だけ1、他の11年は2。B = 23/12 = 1.91666…
'             2014年1月 = 2 - 1.91666… = 0.08333…、2022年1月 = 1 - 1.91666… = -0.91666…
'             予測の2026年1月もこのB（2026年をBの平均に入れない）。A=2 なので 0.08333…

' 公式の取得先。ネットからの取得はこのURLだけに限る。
Public Const OFFICIAL_CSV_URL As String = "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
' X-13のuser変数名とファイル名
Public Const VARIABLE_NAME As String = "iip_hol"
Public Const HOL_FILE_NAME As String = "_holiday.hol"
Public Const SOURCE_COPY_NAME As String = "_holiday_source.csv"
Public Const CALCULATION_FILE_NAME As String = "_holiday_calculation.tsv"
Public Const META_FILE_NAME As String = "_holiday.meta.txt"
Public Const METHOD_ID As String = "METI_IIP_2020_weekday_holidays_A_minus_actual_month_mean_B"
' 取得記録（JSON）のファイル名。CSVと同じフォルダに置く。
Private Const PROVENANCE_FILE_NAME As String = "provenance.json"
Private Const PROVENANCE_SUFFIX As String = ".provenance.json"
' HOLに書く小数桁（経産省の公表表は3桁。計算には丸めずに近い12桁を渡す）
Public Const HOL_DECIMALS As Long = 12
' ネット取得の待ち時間（ミリ秒）
Private Const HTTP_TIMEOUT_MS As Long = 30000

' ================= 画面の操作 =================

' 目的  : 公式URLからCSVを取得し、新しいファイルとして保存する。成功したらControlのパスを更新する。
' 副作用: HOLIDAY_DOWNLOAD_DIR に syukujitsu_YYYYMMDDThhmmssZ.csv と同名の .provenance.json を作る。
'         HOLIDAY_CSV_PATH を相対パスで更新し、HOLIDAY_CSV_ORIGIN／ACQUIRED_AT を空欄に戻す
'         （取得元・取得日時は取得記録ファイルから読むため）。
' エラー: 取得・検査に失敗した場合はControlを変えずに知らせる。ローカルCSVのパスで続行できる。
'         認証・プロキシ・セキュリティ設定は変更しない。転送（リダイレクト）は受け付けない。
Public Sub DownloadOfficialCsv()
    Dim http As Object, folder As String, path As String, acquired As String, body() As Byte, loaded As Object

    On Error GoTo Failed
    folder = modPackagePaths.RequireFolder("HOLIDAY_DOWNLOAD_DIR", True)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = False          ' 6 = WinHttpRequestOption_EnableRedirects。公式URL以外へ移らない。
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    http.Open "GET", OFFICIAL_CSV_URL, False
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 2800, , "HTTPの応答が200ではありません：" & http.Status
    body = http.ResponseBody
    acquired = modMain.UtcNowText()
    path = modMain.CombinePath(folder, "syukujitsu_" & Replace(Replace(acquired, ":", ""), "-", "") & ".csv")
    modX13IO.WriteNewBytes path, body
    Set loaded = LoadCsv(path, "AUTO")
    modX13IO.WriteNewText path & PROVENANCE_SUFFIX, "{" & vbCrLf & _
        "  ""url"": """ & OFFICIAL_CSV_URL & """," & vbCrLf & _
        "  ""retrieved_utc"": """ & acquired & """," & vbCrLf & _
        "  ""sha256"": """ & loaded("sha256") & """," & vbCrLf & _
        "  ""bytes"": " & loaded("bytes") & vbCrLf & "}" & vbCrLf
    modMain.SetControlValue "HOLIDAY_CSV_PATH", modPackagePaths.RelativeToPackage(path)
    modMain.SetControlValue "HOLIDAY_CSV_ORIGIN", ""
    modMain.SetControlValue "HOLIDAY_CSV_ACQUIRED_AT", ""
    WriteCalendarSheet loaded, ResolveProvenance(path, loaded("sha256"))
    modLog.WriteLog "DownloadOfficialCsv", "INFO", "取得：" & path & " SHA-256=" & loaded("sha256")
    modUI.ShowMessage "内閣府の祝日CSVを取得しました。" & vbLf & "保存先：" & path & vbLf & _
        "収録年：" & loaded("firstYear") & "～" & loaded("lastYear") & "年、" & loaded("rows").Count & "日" & vbLf & _
        "SHA-256：" & loaded("sha256"), vbInformation, "祝日CSVの取得"
    Exit Sub
Failed:
    modLog.WriteLog "DownloadOfficialCsv", "WARN", Err.Description
    modUI.ShowMessage "公式URLからの取得に失敗しました。Controlの設定は変えていません。" & vbLf & Err.Description & vbLf & vbLf & _
        "手元のCSVで続行する場合は「祝日CSVの選択」を使うか、ControlのHOLIDAY_CSV_PATHにパスを記入してください。", _
        vbExclamation, "祝日CSVの取得"
End Sub

' 目的  : 手元の祝日CSVを選び、検査してからControlのパスを更新する。
' 副作用: HOLIDAY_CSV_PATH を更新（パッケージ内なら相対パス）。取得元・取得日時を空欄に戻す。
Public Sub ChooseLocalCsv()
    Dim dialog As Object, path As String, loaded As Object, provenance As Variant
    On Error GoTo Failed
    Set dialog = Application.FileDialog(3)
    dialog.Title = "内閣府「国民の祝日」CSVを選択してください"
    dialog.AllowMultiSelect = False
    dialog.Filters.Clear
    dialog.Filters.Add "CSV", "*.csv"
    If dialog.Show <> -1 Then Exit Sub
    path = CStr(dialog.SelectedItems(1))
    Set loaded = LoadCsv(path, modMain.GetControlValue("HOLIDAY_CSV_ENCODING", "AUTO"))
    modMain.SetControlValue "HOLIDAY_CSV_PATH", modPackagePaths.RelativeToPackage(path)
    modMain.SetControlValue "HOLIDAY_CSV_ORIGIN", ""
    modMain.SetControlValue "HOLIDAY_CSV_ACQUIRED_AT", ""
    provenance = ResolveProvenance(path, loaded("sha256"))
    WriteCalendarSheet loaded, provenance
    If Len(provenance(0)) = 0 Then
        modUI.ShowMessage "祝日CSVを確認しました（収録年 " & loaded("firstYear") & "～" & loaded("lastYear") & "年）。" & vbLf & _
            "同じフォルダに取得記録が無いため、ControlのHOLIDAY_CSV_ORIGINとHOLIDAY_CSV_ACQUIRED_ATに入手元と取得日を記入してください。", _
            vbInformation, "祝日CSVの選択"
    Else
        modUI.ShowMessage "祝日CSVを確認しました（収録年 " & loaded("firstYear") & "～" & loaded("lastYear") & "年）。" & vbLf & _
            "取得元：" & provenance(0) & vbLf & "取得日時：" & provenance(1), vbInformation, "祝日CSVの選択"
    End If
    Exit Sub
Failed:
    modUI.ShowMessage "祝日CSVを使えません。" & vbLf & Err.Description, vbExclamation, "祝日CSVの選択"
End Sub

' 目的  : 現在の実績期間・予測月数で祝日変数を作り、2つのシートと03_holiday\generatedへ保存する。
'         X-13は起動しない。季節調整の実行時には、ジョブごとに同じ計算をやり直して使う。
Public Sub BuildFromUI()
    Dim context As Object, folder As String, stem As String
    On Error GoTo Failed
    Set context = PrepareForRun(modX13Options.FirstMonth(), modX13Options.ActualMonths(), modX13Options.ForecastMonths())
    folder = modPackagePaths.RequireFolder("HOLIDAY_GENERATED_DIR", True)
    stem = WriteGeneratedFiles(context, folder)
    ThisWorkbook.Worksheets(modSheetNames.HolidayCalculationSheet()).Activate
    modUI.ShowMessage "祝日変数を作成しました（" & context("calculation")("rows").Count & "か月）。" & vbLf & _
        "計算表は「" & modSheetNames.HolidayCalculationSheet() & "」、CSVの日付一覧は「" & _
        modSheetNames.HolidayCalendarSheet() & "」です。" & vbLf & "保存先：" & modMain.CombinePath(folder, stem) & ".*", _
        vbInformation, "祝日変数の作成"
    Exit Sub
Failed:
    modUI.ShowMessage "祝日変数を作成できません。" & vbLf & Err.Description, vbExclamation, "祝日変数の作成"
End Sub

' ================= 読込と検査 =================

' 目的  : 祝日CSVを読み、全行を検査する。
' 入力  : path（絶対パス）、encodingSetting（AUTO/CP932/UTF8）
' 出力  : Dictionary
'           "path","encoding","sha256","crc32","bytes","modified","header","firstYear","lastYear"
'           "rows"  : Collection。要素は Array(CSV行番号, 日付(Date), 名称)
'           "dates" : Dictionary。キー=日付のシリアル値(Long)、値=名称
' エラー: 文字コード不明、見出しが2列でない、列数違い、日付の形式違い・存在しない日付、名称の空欄、
'         日付の重複、途中の空行を、行番号を付けて停止する。読み飛ばさない。
Public Function LoadCsv(ByVal path As String, ByVal encodingSetting As String) As Object
    Dim result As Object, bytes() As Byte, encoding As String, text As String, lines As Variant
    Dim lastLine As Long, i As Long, fields As Variant, day As Date, name As String
    Dim rows As New Collection, dates As Object, firstYear As Long, lastYear As Long, fso As Object

    modX13IO.RequireLocalPath path
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Err.Raise vbObjectError + 2802, "LoadCsv", "祝日CSVが見つかりません：" & path
    bytes = modX13IO.ReadBytes(path)
    encoding = UCase$(encodingSetting)
    If encoding = "AUTO" Or Len(encoding) = 0 Then encoding = modIIPTextTable.DetectEncoding(bytes)
    text = modIIPTextTable.DecodeBytes(bytes, encoding)
    text = Replace(text, vbCrLf, vbLf)
    If InStr(text, vbCr) > 0 Then Err.Raise vbObjectError + 2803, "LoadCsv", "改行コードが不正です（CRだけの改行）"
    lines = Split(text, vbLf)
    lastLine = UBound(lines)
    Do While lastLine >= 0
        If Len(lines(lastLine)) > 0 Then Exit Do
        lastLine = lastLine - 1          ' 末尾の改行だけは許す
    Loop
    If lastLine < 1 Then Err.Raise vbObjectError + 2804, "LoadCsv", "祝日CSVにデータ行がありません"
    fields = SplitCsvLine(CStr(lines(0)))
    If UBound(fields) <> 1 Then Err.Raise vbObjectError + 2805, "LoadCsv", "1行目（見出し）が2列ではありません"
    If TryParseDate(CStr(fields(0)), day) Then Err.Raise vbObjectError + 2806, "LoadCsv", "1行目が見出しではなく日付です"

    Set dates = CreateObject("Scripting.Dictionary")
    For i = 1 To lastLine
        fields = SplitCsvLine(CStr(lines(i)))
        If UBound(fields) <> 1 Then Err.Raise vbObjectError + 2807, "LoadCsv", (i + 1) & "行目：列数が2ではありません（空行も不可）"
        If Not TryParseDate(CStr(fields(0)), day) Then Err.Raise vbObjectError + 2808, "LoadCsv", (i + 1) & "行目：日付が不正です：" & fields(0)
        name = Trim$(CStr(fields(1)))
        If Len(name) = 0 Then Err.Raise vbObjectError + 2809, "LoadCsv", (i + 1) & "行目：祝日名が空欄です"
        If dates.Exists(CLng(day)) Then Err.Raise vbObjectError + 2810, "LoadCsv", (i + 1) & "行目：日付が重複しています：" & Format$(day, "yyyy-mm-dd")
        dates.Add CLng(day), name
        rows.Add Array(i + 1, day, name)
        If firstYear = 0 Or Year(day) < firstYear Then firstYear = Year(day)
        If Year(day) > lastYear Then lastYear = Year(day)
    Next i

    Set result = CreateObject("Scripting.Dictionary")
    result.Add "path", path
    result.Add "encoding", encoding
    result.Add "sha256", modHash.FileSHA256(path)
    result.Add "crc32", modX13IO.FileCRC(path)
    result.Add "bytes", fso.GetFile(path).Size
    result.Add "modified", Format$(fso.GetFile(path).DateLastModified, "yyyy-mm-dd hh:nn:ss")
    result.Add "header", CStr(lines(0))
    result.Add "firstYear", firstYear
    result.Add "lastYear", lastYear
    result.Add "rows", rows
    result.Add "dates", dates
    Set LoadCsv = result
End Function

' 1行を区切る。両端を囲む引用符だけ外す（内閣府CSVは引用符を使わない）。区切りを含む引用は列数違いとして止まる。
Private Function SplitCsvLine(ByVal lineText As String) As Variant
    Dim parts As Variant, i As Long
    parts = Split(lineText, ",")
    For i = 0 To UBound(parts)
        parts(i) = Trim$(parts(i))
        If Len(parts(i)) >= 2 And Left$(parts(i), 1) = """" And Right$(parts(i), 1) = """" Then parts(i) = Mid$(parts(i), 2, Len(parts(i)) - 2)
    Next i
    SplitCsvLine = parts
End Function

' "YYYY/M/D" または "YYYY-MM-DD" を日付へ。存在しない日付（2019/2/30等）はFalse。
Private Function TryParseDate(ByVal text As String, ByRef result As Date) As Boolean
    Dim re As Object, m As Object, y As Long, mo As Long, d As Long
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^([0-9]{4})[/-]([0-9]{1,2})[/-]([0-9]{1,2})$"
    If Not re.Test(Trim$(text)) Then Exit Function
    Set m = re.Execute(Trim$(text))(0)
    y = CLng(m.SubMatches(0)): mo = CLng(m.SubMatches(1)): d = CLng(m.SubMatches(2))
    If y < 1900 Or y > 2999 Or mo < 1 Or mo > 12 Or d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, mo, d)
    ' DateSerialは2月30日を3月2日へ繰り上げるため、元の年月日と一致するかで存在を確かめる。
    TryParseDate = (Year(result) = y And Month(result) = mo And Day(result) = d)
End Function

' 目的  : 必要な年がCSVに収録されていることを確かめる。将来の祝日は推測しない。
' 規則  : 必要な最初の年・最後の年がCSVの収録範囲内であること。
'         さらに各年に元日（1月1日）と勤労感謝の日（11月23日）の両方があること
'         （固定日の祝日を、年の途中で切れたCSVや欠けた年を見つける目印にだけ使う。計算には使わない）。
Public Sub RequireCoverage(ByVal loaded As Object, ByVal firstYear As Long, ByVal lastYear As Long)
    Dim y As Long
    If lastYear > loaded("lastYear") Then Err.Raise vbObjectError + 2811, "RequireCoverage", _
        lastYear & "年の祝日がCSVにありません（CSVの最終年は" & loaded("lastYear") & "年）。予測期間が公表済みの範囲を超えています。将来の祝日は推測しません。"
    If firstYear < loaded("firstYear") Then Err.Raise vbObjectError + 2812, "RequireCoverage", _
        firstYear & "年の祝日がCSVにありません（CSVの最初の年は" & loaded("firstYear") & "年）。"
    For y = firstYear To lastYear
        If Not loaded("dates").Exists(CLng(DateSerial(y, 1, 1))) Or Not loaded("dates").Exists(CLng(DateSerial(y, 11, 23))) Then _
            Err.Raise vbObjectError + 2813, "RequireCoverage", y & "年の収録が不完全です（元日または11月23日がありません）。"
    Next y
End Sub

' 目的  : CSVの取得元・取得日時を決める。
' 規則  : 1) CSVと同じフォルダの「<CSV名>.provenance.json」または「provenance.json」で、
'            記録されたSHA-256がCSVと一致するもの  2) ControlのHOLIDAY_CSV_ORIGIN／ACQUIRED_AT
'         どちらも無ければ空文字（呼出側が「未記入」として記録する）。
' 出力  : Array(取得元, 取得日時, 出所の説明)
Public Function ResolveProvenance(ByVal csvPath As String, ByVal sha256 As String) As Variant
    Dim fso As Object, candidate As Variant, text As String, recordedSha As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each candidate In Array(csvPath & PROVENANCE_SUFFIX, fso.BuildPath(fso.GetParentFolderName(csvPath), PROVENANCE_FILE_NAME))
        If fso.FileExists(CStr(candidate)) Then
            text = modX13IO.ReadText(CStr(candidate), "UTF8")
            recordedSha = LCase$(JsonText(text, "sha256"))
            If recordedSha = LCase$(sha256) And Len(JsonText(text, "url")) > 0 Then
                ResolveProvenance = Array(JsonText(text, "url"), JsonText(text, "retrieved_utc"), _
                    "取得記録ファイル（SHA-256一致）：" & fso.GetFileName(CStr(candidate)))
                Exit Function
            End If
        End If
    Next candidate
    ResolveProvenance = Array(modMain.GetControlValue("HOLIDAY_CSV_ORIGIN", ""), modMain.GetControlValue("HOLIDAY_CSV_ACQUIRED_AT", ""), _
        "ControlのHOLIDAY_CSV_ORIGIN／HOLIDAY_CSV_ACQUIRED_AT")
End Function

' 取得記録の簡単なJSONから、"キー": "文字列" の値を取り出す（入れ子は扱わない）。無ければ空文字。
Private Function JsonText(ByVal text As String, ByVal key As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & key & """\s*:\s*""([^""]*)"""
    Set matches = re.Execute(text)
    If matches.Count > 0 Then JsonText = CStr(matches(0).SubMatches(0))
End Function

' ================= 計算 =================

' 目的  : 祝日変数（A-B）を計算する。
' 入力  : loaded（LoadCsvの戻り値）、firstMonthNumber（実績の最初の月番号）、actualMonths（実績月数）、forecastMonths
' 出力  : Dictionary
'           "rows" : Collection。要素は Array(月番号, 区分(実績/予測), A, B, Bの分母, A-B, HOL文字列, 算入した祝日, 算入しない祝日)
'           "first","actual","forecast","bYears"（Bに使った期間の文字列）
' エラー: 必要な年がCSVに無い場合は停止（RequireCoverage）。
Public Function Calculate(ByVal loaded As Object, ByVal firstMonthNumber As Long, ByVal actualMonths As Long, _
                          ByVal forecastMonths As Long) As Object
    Dim total As Long, k As Long, monthNumber As Long, countA() As Long, counted() As String, skipped() As String
    Dim sumByMonth(1 To 12) As Double, countByMonth(1 To 12) As Long, item As Variant, day As Date
    Dim rows As New Collection, b As Double, value As Double, result As Object, calendarMonth As Long

    total = actualMonths + forecastMonths
    RequireCoverage loaded, modX13IO.YearOf(firstMonthNumber), modX13IO.YearOf(firstMonthNumber + total - 1)
    ReDim countA(0 To total - 1): ReDim counted(0 To total - 1): ReDim skipped(0 To total - 1)

    ' A：月曜日～金曜日に当たるCSVの日付を、その年月に数える。
    For Each item In loaded("rows")
        day = item(1)
        monthNumber = modX13IO.MonthIndex(Year(day), Month(day))
        k = monthNumber - firstMonthNumber
        If k >= 0 And k < total Then
            If Weekday(day, vbMonday) <= 5 Then
                countA(k) = countA(k) + 1
                counted(k) = counted(k) & IIf(Len(counted(k)) > 0, "、", "") & Format$(day, "m/d") & item(2)
            Else
                skipped(k) = skipped(k) & IIf(Len(skipped(k)) > 0, "、", "") & Format$(day, "m/d") & item(2) & "（" & WeekdayLabel(day) & "）"
            End If
        End If
    Next item

    ' B：実績期間（予測月を含めない）の同じ暦月のAの平均。
    For k = 0 To actualMonths - 1
        calendarMonth = modX13IO.MonthOf(firstMonthNumber + k)
        sumByMonth(calendarMonth) = sumByMonth(calendarMonth) + countA(k)
        countByMonth(calendarMonth) = countByMonth(calendarMonth) + 1
    Next k

    For k = 0 To total - 1
        calendarMonth = modX13IO.MonthOf(firstMonthNumber + k)
        If countByMonth(calendarMonth) = 0 Then Err.Raise vbObjectError + 2814, "Calculate", "実績期間に" & calendarMonth & "月がありません"
        b = sumByMonth(calendarMonth) / countByMonth(calendarMonth)
        value = countA(k) - b
        rows.Add Array(firstMonthNumber + k, IIf(k < actualMonths, "実績", "予測"), countA(k), b, _
            countByMonth(calendarMonth), value, modX13IO.FixedDecimalText(value, HOL_DECIMALS), counted(k), skipped(k))
    Next k

    Set result = CreateObject("Scripting.Dictionary")
    result.Add "rows", rows
    result.Add "first", firstMonthNumber
    result.Add "actual", actualMonths
    result.Add "forecast", forecastMonths
    result.Add "bYears", modX13IO.MonthText(firstMonthNumber) & "～" & modX13IO.MonthText(firstMonthNumber + actualMonths - 1)
    Set Calculate = result
End Function

Private Function WeekdayLabel(ByVal day As Date) As String
    WeekdayLabel = Choose(Weekday(day, vbMonday), "月", "火", "水", "木", "金", "土", "日")
End Function

' 目的  : 実行前の準備。CSVを読み、期間をすべて覆うことを確かめ、計算して2つのシートに書く。
' 出力  : Dictionary（"loaded"、"calculation"、"origin"、"acquired"、"provenanceSource"）
' 副作用: 原データ_祝日CSV と 計算用_祝日変数 を作り直す。
' エラー: CSV不正・不足は停止。X-13は起動前。
Public Function PrepareForRun(ByVal firstMonthNumber As Long, ByVal actualMonths As Long, ByVal forecastMonths As Long) As Object
    Dim path As String, context As Object, provenance As Variant
    path = modPackagePaths.RequireFile("HOLIDAY_CSV_PATH")
    Set context = CreateObject("Scripting.Dictionary")
    context.Add "loaded", LoadCsv(path, modMain.GetControlValue("HOLIDAY_CSV_ENCODING", "AUTO"))
    context.Add "calculation", Calculate(context("loaded"), firstMonthNumber, actualMonths, forecastMonths)
    provenance = ResolveProvenance(path, context("loaded")("sha256"))
    context.Add "origin", IIf(Len(provenance(0)) > 0, provenance(0), "（未記入）")
    context.Add "acquired", IIf(Len(provenance(1)) > 0, provenance(1), "（未記入）")
    context.Add "provenanceSource", provenance(2)
    If Len(provenance(0)) = 0 Or Len(provenance(1)) = 0 Then _
        modLog.WriteLog "PrepareForRun", "WARN", "祝日CSVの取得元または取得日時が未記入です：" & path
    WriteCalendarSheet context("loaded"), provenance
    WriteCalculationSheet context
    Set PrepareForRun = context
End Function

' ================= シートとファイル =================

' 目的: 原データ_祝日CSV に、CSVの全行と判定（平日か）を書く。右側に原本の情報を書く。
Public Sub WriteCalendarSheet(ByVal loaded As Object, ByVal provenance As Variant)
    Dim ws As Worksheet, data() As Variant, i As Long, item As Variant, day As Date
    Set ws = modMain.EnsureWorksheet(modSheetNames.HolidayCalendarSheet())
    modX13IO.ResetSheet ws, Array("CSV行番号", "日付", "祝日名", "曜日", "月～金に当たるか", "年月")
    ReDim data(1 To loaded("rows").Count, 1 To 6)
    i = 0
    For Each item In loaded("rows")
        i = i + 1
        day = item(1)
        data(i, 1) = item(0)
        data(i, 2) = Format$(day, "yyyy-mm-dd")
        data(i, 3) = item(2)
        data(i, 4) = WeekdayLabel(day)
        data(i, 5) = IIf(Weekday(day, vbMonday) <= 5, "月～金（Aに数える）", "土日（数えない）")
        data(i, 6) = Format$(day, "yyyy-mm")
    Next item
    ws.Range("B:B,F:F").NumberFormat = "@"
    ws.Cells(2, 1).Resize(UBound(data, 1), 6).Value2 = data
    WriteInfoBlock ws, 8, Array( _
        Array("原本ファイル", modPackagePaths.RelativeToPackage(loaded("path"))), _
        Array("取得元", IIf(Len(provenance(0)) > 0, provenance(0), "（未記入）")), _
        Array("取得日時（UTC）", IIf(Len(provenance(1)) > 0, provenance(1), "（未記入）")), _
        Array("取得元・日時の出所", provenance(2)), _
        Array("ファイル更新日時", loaded("modified")), _
        Array("SHA-256", loaded("sha256")), _
        Array("CRC32", loaded("crc32")), _
        Array("バイト数", loaded("bytes")), _
        Array("文字コード", loaded("encoding")), _
        Array("見出し行", loaded("header")), _
        Array("収録年", loaded("firstYear") & "～" & loaded("lastYear")), _
        Array("日付の件数", loaded("rows").Count), _
        Array("公式の取得先", OFFICIAL_CSV_URL))
    ws.Columns("A:F").AutoFit
End Sub

' 目的: 計算用_祝日変数 に年月・A・B・A-B・区分を書く。右側に計算の条件を書く。
Public Sub WriteCalculationSheet(ByVal context As Object)
    Dim ws As Worksheet, calc As Object, data() As Variant, i As Long, item As Variant
    Set calc = context("calculation")
    Set ws = modMain.EnsureWorksheet(modSheetNames.HolidayCalculationSheet())
    modX13IO.ResetSheet ws, Array("年月", "区分", "A_月～金の祝日数", "B_実績期間の同月平均", "Bの分母_実績の同月数", _
        "祝日変数_A-B", "HOL書出値", "Aに数えた祝日", "数えない祝日（土日）")
    ReDim data(1 To calc("rows").Count, 1 To 9)
    i = 0
    For Each item In calc("rows")
        i = i + 1
        data(i, 1) = modX13IO.MonthText(item(0))
        data(i, 2) = item(1)
        data(i, 3) = item(2)
        data(i, 4) = item(3)
        data(i, 5) = item(4)
        data(i, 6) = item(5)
        data(i, 7) = item(6)
        data(i, 8) = item(7)
        data(i, 9) = item(8)
    Next item
    ws.Range("A:B,G:I").NumberFormat = "@"
    ws.Range("D:D,F:F").NumberFormat = "0.000000000000"
    ws.Cells(2, 1).Resize(UBound(data, 1), 9).Value2 = data
    WriteInfoBlock ws, 11, Array( _
        Array("方法", "経産省IIP（2020年基準改定の概要 PDF34頁）：A=月～金の祝日数、B=実績期間の同月平均、変数=A-B"), _
        Array("方法ID", METHOD_ID), _
        Array("実績期間（Bの計算に使う期間）", calc("bYears")), _
        Array("実績月数", calc("actual")), _
        Array("予測月数", calc("forecast")), _
        Array("出力月数", calc("rows").Count), _
        Array("X-13の変数名", VARIABLE_NAME & "（usertype=holiday）"), _
        Array("原本CSV", modPackagePaths.RelativeToPackage(context("loaded")("path"))), _
        Array("原本のSHA-256", context("loaded")("sha256")), _
        Array("作成日時（UTC）", modMain.UtcNowText()))
    ws.Columns("A:G").AutoFit
    ws.Columns("H:I").ColumnWidth = 40
End Sub

Private Sub WriteInfoBlock(ByVal ws As Worksheet, ByVal column As Long, ByVal pairs As Variant)
    Dim i As Long
    ws.Cells(1, column).Resize(1, 2).Value2 = Array("項目", "内容")
    ws.Cells(1, column).Resize(1, 2).Font.Bold = True
    ws.Cells(2, column + 1).Resize(UBound(pairs) + 1, 1).NumberFormat = "@"
    For i = 0 To UBound(pairs)
        ws.Cells(i + 2, column).Value2 = pairs(i)(0)
        ws.Cells(i + 2, column + 1).Value2 = CStr(pairs(i)(1))
    Next i
    ws.Columns(column).ColumnWidth = 30
    ws.Columns(column + 1).ColumnWidth = 70
End Sub

' HOLの本文（1行1か月、実績＋予測）。
Private Function HolText(ByVal calc As Object) As String
    Dim item As Variant, body As String
    For Each item In calc("rows")
        body = body & item(6) & vbCrLf
    Next item
    HolText = body
End Function

' 計算根拠の本文（タブ区切り：年月・区分・A・B・Bの分母・A-B）。
Private Function CalculationText(ByVal calc As Object) As String
    Dim item As Variant, body As String
    body = "date" & vbTab & "kind" & vbTab & "A" & vbTab & "B" & vbTab & "B_denominator" & vbTab & "A_minus_B" & vbCrLf
    For Each item In calc("rows")
        body = body & modX13IO.MonthText(item(0)) & vbTab & IIf(item(1) = "実績", "actual", "forecast") & vbTab & item(2) & vbTab & _
            modX13IO.FixedDecimalText(item(3), HOL_DECIMALS) & vbTab & item(4) & vbTab & item(6) & vbCrLf
    Next item
    CalculationText = body
End Function

' 記録ファイル（キー=値）の本文。
Private Function MetaText(ByVal context As Object) As String
    Dim calc As Object
    Set calc = context("calculation")
    MetaText = "METHOD=" & METHOD_ID & vbCrLf & _
        "VARIABLE=" & VARIABLE_NAME & vbCrLf & _
        "FIRST=" & modX13IO.MonthText(calc("first")) & vbCrLf & _
        "ACTUAL_MONTHS=" & calc("actual") & vbCrLf & _
        "FORECAST_MONTHS=" & calc("forecast") & vbCrLf & _
        "ROWS=" & calc("rows").Count & vbCrLf & _
        "HOL_DECIMALS=" & HOL_DECIMALS & vbCrLf & _
        "SOURCE_PATH=" & modPackagePaths.RelativeToPackage(context("loaded")("path")) & vbCrLf & _
        "SOURCE_ORIGIN=" & context("origin") & vbCrLf & _
        "SOURCE_ACQUIRED_AT=" & context("acquired") & vbCrLf & _
        "SOURCE_PROVENANCE=" & context("provenanceSource") & vbCrLf & _
        "SOURCE_SHA256=" & context("loaded")("sha256") & vbCrLf & _
        "SOURCE_ENCODING=" & context("loaded")("encoding") & vbCrLf & _
        "SOURCE_YEARS=" & context("loaded")("firstYear") & "-" & context("loaded")("lastYear") & vbCrLf
End Function

' 目的  : 03_holiday\generated へ、HOL・計算根拠・記録を新しいファイル名で保存する。
' 出力  : 保存したファイル名の共通部分（例 holiday_201401_202612_20261016T090000Z）
' エラー: 同名ファイルがあれば停止（上書きしない）。
Public Function WriteGeneratedFiles(ByVal context As Object, ByVal folder As String) As String
    Dim calc As Object, stem As String
    Set calc = context("calculation")
    stem = "holiday_" & modX13IO.MonthCompact(calc("first")) & "_" & _
        modX13IO.MonthCompact(calc("first") + calc("rows").Count - 1) & "_" & Replace(Replace(modMain.UtcNowText(), ":", ""), "-", "")
    modX13IO.WriteNewText modMain.CombinePath(folder, stem & ".hol"), HolText(calc)
    modX13IO.WriteNewText modMain.CombinePath(folder, stem & "_calculation.tsv"), CalculationText(calc)
    modX13IO.WriteNewText modMain.CombinePath(folder, stem & ".meta.txt"), MetaText(context)
    WriteGeneratedFiles = stem
End Function

' 目的  : 1つの計算（ジョブ）フォルダに、HOLと追跡用の記録を書く。
' 書くもの: _holiday.hol（X-13が読む値。1行1か月、実績＋予測）
'           _holiday_calculation.tsv（年月・区分・A・B・分母・A-B）
'           _holiday_source.csv（原本CSVのバイト単位の複製。SHA-256で一致を確認）
'           _holiday.meta.txt（方法、期間、原本のパス・取得元・取得日時・SHA-256）
' エラー: 原本が読込後に変わっていた場合は停止（SHA-256の不一致）。
Public Sub WriteJobHolidayFiles(ByVal context As Object, ByVal folder As String)
    Dim copyPath As String
    modX13IO.WriteNewText modMain.CombinePath(folder, HOL_FILE_NAME), HolText(context("calculation"))
    modX13IO.WriteNewText modMain.CombinePath(folder, CALCULATION_FILE_NAME), CalculationText(context("calculation"))
    copyPath = modMain.CombinePath(folder, SOURCE_COPY_NAME)
    modX13IO.CopyFileVerified context("loaded")("path"), copyPath
    If modHash.FileSHA256(copyPath) <> context("loaded")("sha256") Then _
        Err.Raise vbObjectError + 2816, "WriteJobHolidayFiles","祝日CSVが読込後に変わっています。祝日変数を作り直してください。"
    modX13IO.WriteNewText modMain.CombinePath(folder, META_FILE_NAME), MetaText(context)
End Sub

' 目的: SPCのregressionブロックに入れる祝日変数の指定（複数行）。
' 入力: fileReference = SPCから見たHOLのパス（例 "j000001/_holiday.hol"）
Public Function RegressionLines(ByVal firstMonthNumber As Long, ByVal fileReference As String) As String
    RegressionLines = " user=(" & VARIABLE_NAME & ")" & vbCrLf & _
        " usertype=holiday" & vbCrLf & _
        " start=" & modX13IO.X13StartText(firstMonthNumber) & vbCrLf & _
        " file=""" & fileReference & """" & vbCrLf & _
        " format=""free"""
End Function

' 目的  : 保存済みフォルダのHOLが、同じフォルダの原本CSVの複製から作り直した値と一字一句一致するか確かめる。
'         最終保存の前と、保存後の読戻しで使う（出力HOLと原本CSVの追跡関係の確認）。
' エラー: SHA-256、期間、行数、値のいずれかが合わなければ停止。
Public Sub VerifySnapshot(ByVal folder As String)
    Dim meta As Object, loaded As Object, calc As Object, lines As Variant, item As Variant, i As Long
    Set meta = modX13IO.ReadMetaFile(modMain.CombinePath(folder, META_FILE_NAME))
    If meta("METHOD") <> METHOD_ID Then Err.Raise vbObjectError + 2817, "VerifySnapshot", "祝日変数の方法IDが違います"
    Set loaded = LoadCsv(modMain.CombinePath(folder, SOURCE_COPY_NAME), meta("SOURCE_ENCODING"))
    If loaded("sha256") <> meta("SOURCE_SHA256") Then Err.Raise vbObjectError + 2818, "VerifySnapshot", "原本CSVの複製のSHA-256が記録と違います"
    Set calc = Calculate(loaded, modX13IO.MonthIndexFromText(meta("FIRST")), CLng(meta("ACTUAL_MONTHS")), CLng(meta("FORECAST_MONTHS")))
    lines = Split(Replace(modX13IO.ReadText(modMain.CombinePath(folder, HOL_FILE_NAME), "UTF8"), vbCrLf, vbLf), vbLf)
    ' HOLは各行が改行で終わるため、Splitの最後の要素は空文字になる。
    If UBound(lines) <> calc("rows").Count Then Err.Raise vbObjectError + 2819, "VerifySnapshot", "HOLの行数が計算と違います"
    If Len(lines(UBound(lines))) > 0 Then Err.Raise vbObjectError + 2819, "VerifySnapshot", "HOLの末尾が改行で終わっていません"
    i = 0
    For Each item In calc("rows")
        If CStr(lines(i)) <> CStr(item(6)) Then Err.Raise vbObjectError + 2820, "VerifySnapshot", _
            "HOLの値が原本CSVから作り直した値と違います：" & modX13IO.MonthText(item(0))
        i = i + 1
    Next item
End Sub
