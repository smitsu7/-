Attribute VB_Name = "modZeroReplace"
Option Explicit

' 役割      : 原指数（原データ）から、DATへ書く計算用の値を作る。対数計算の準備として、
'             数値の0だけをepsilon（初期値0.000001）に置き換え、その記録を残す。
'             原データのシートと取込元Excelは変更しない。
' 主な入口  : PrepareComputationValues / RecordReplacements / ReplacementFileText
' 依存先    : modSeriesSettings（cfg）、modX13Options、modX13IO、modMain、modSheetNames
' 規則      :
'   変換がlogまたはauto（X-13がlogを選ぶ可能性がある）の場合
'     正の数 → そのまま
'     0      → 置換ONならepsilon。置換OFFなら実行前エラー
'     負の数 → 実行前エラー（epsilonへ変えない）
'   変換がnoneの場合
'     数値はすべてそのまま（0も負の数も保持）
'   どの変換でも、空欄・文字列（欠損記号を含む）・エラー値は実行前エラー（0として扱わない）。
'   DATへ書く固定小数の文字列にした結果、log/autoで0以下になる値があれば停止する
'   （例：DATの小数桁が5なら 0.000001 は "0.00000" となり消えるため停止）。
'   transform.constant（全観測への定数加算）やtrimzeroは使わない。

Public Const REASON_ZERO_TO_EPSILON As String = "zero_to_epsilon"

' 目的  : 1系列の計算用の値を作る。
' 入力  : cfg（解決済み設定）、rawValues（原データのセル値。1始まりの配列）、firstMonthNumber（月番号）、
'         datDecimals（DATの小数桁）
' 出力  : Dictionary
'           "transform"    : log/none/auto
'           "enabled"      : 置換ON/OFF（YES/NO）
'           "epsilon"      : 置換値（Double。OFFでも検査済みの値）
'           "original"     : 原データの数値（Double配列、1始まり）
'           "computed"     : 計算用の値（Double配列）
'           "datText"      : DATに書く文字列（String配列）
'           "replaced"     : Collection。要素は Array(月番号, 元値, 計算値, DAT文字列)
' 副作用: なし（シートは読まない。呼出側が渡した値だけを使う）
' エラー: 上の規則に反する値があれば、年月と理由を付けて停止する。
Public Function PrepareComputationValues(ByVal cfg As Object, ByVal rawValues As Variant, _
                                         ByVal firstMonthNumber As Long, ByVal datDecimals As Long) As Object
    Dim result As Object, transform As String, enabled As Boolean, epsilon As Double
    Dim count As Long, i As Long, original() As Double, computed() As Double, datText() As String
    Dim replaced As New Collection, monthLabel As String, rawValue As Variant

    transform = modX13Options.TransformFunction(cfg)
    enabled = modMain.ParseYesNo(modSeriesSettings.SettingText(cfg, "X13_ZERO_REPLACE"), "X13_ZERO_REPLACE")
    epsilon = CheckedEpsilon(modSeriesSettings.SettingText(cfg, "X13_ZERO_EPSILON"))
    count = UBound(rawValues) - LBound(rawValues) + 1
    ReDim original(1 To count): ReDim computed(1 To count): ReDim datText(1 To count)

    For i = 1 To count
        rawValue = rawValues(LBound(rawValues) + i - 1)
        monthLabel = modX13IO.MonthText(firstMonthNumber + i - 1)
        original(i) = RequireNumeric(rawValue, monthLabel)
        computed(i) = original(i)
        If NeedsPositive(transform) Then
            If original(i) < 0 Then
                Err.Raise vbObjectError + 2700, "PrepareComputationValues", monthLabel & "：負の値（" & _
                    modX13IO.NumberText(original(i)) & "）は対数変換できません。epsilonへは置き換えません。"
            ElseIf original(i) = 0 Then
                If Not enabled Then Err.Raise vbObjectError + 2701, "PrepareComputationValues", monthLabel & _
                    "：値が0です。変換が" & transform & "でゼロ置換がNOのため計算できません。"
                computed(i) = epsilon
            End If
        End If
        datText(i) = modX13IO.FixedDecimalText(computed(i), datDecimals)
        If NeedsPositive(transform) And modX13IO.NumberValue(datText(i)) <= 0 Then
            If original(i) = 0 Then
                Err.Raise vbObjectError + 2702, "PrepareComputationValues", monthLabel & "：epsilon（" & _
                    modX13IO.NumberText(epsilon) & "）がDATの小数桁" & datDecimals & "で0に丸められます。小数桁を増やしてください。"
            Else
                Err.Raise vbObjectError + 2703, "PrepareComputationValues", monthLabel & "：正の値（" & _
                    modX13IO.NumberText(original(i)) & "）がDATの小数桁" & datDecimals & "で0に丸められます。"
            End If
        End If
        If computed(i) <> original(i) Then replaced.Add Array(firstMonthNumber + i - 1, original(i), computed(i), datText(i))
    Next i

    Set result = CreateObject("Scripting.Dictionary")
    result.Add "transform", transform
    result.Add "enabled", IIf(enabled, "YES", "NO")
    result.Add "epsilon", epsilon
    result.Add "original", original
    result.Add "computed", computed
    result.Add "datText", datText
    result.Add "replaced", replaced
    Set PrepareComputationValues = result
End Function

' 対数を取る可能性がある変換か。autoはX-13がlogを選ぶ可能性があるため、logと同じ規則にする。
Private Function NeedsPositive(ByVal transform As String) As Boolean
    NeedsPositive = (transform = "log" Or transform = "auto")
End Function

' 目的: epsilonの設定値を検査する（0より大きい数値）。
Public Function CheckedEpsilon(ByVal text As String) As Double
    CheckedEpsilon = modX13Options.CheckedNumber(text, "X13_ZERO_EPSILON")
    If CheckedEpsilon <= 0 Then Err.Raise vbObjectError + 2704, "CheckedEpsilon", "X13_ZERO_EPSILON は0より大きい数値で指定してください：" & text
End Function

' 目的  : 原データのセル値が数値であることを確かめる。空欄・文字列・エラー値は0にせず停止する。
Private Function RequireNumeric(ByVal rawValue As Variant, ByVal monthLabel As String) As Double
    If IsError(rawValue) Then Err.Raise vbObjectError + 2705, "RequireNumeric", monthLabel & "：原指数がエラー値です"
    If IsEmpty(rawValue) Then Err.Raise vbObjectError + 2706, "RequireNumeric", monthLabel & "：原指数が空欄です（0として扱いません）"
    Select Case VarType(rawValue)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            RequireNumeric = CDbl(rawValue)
        Case Else
            Err.Raise vbObjectError + 2707, "RequireNumeric", monthLabel & "：原指数が数値ではありません（" & CStr(rawValue) & "）。欠損記号や文字は0として扱いません"
    End Select
End Function

' 目的  : 置換の記録を計算用_ゼロ置換記録シートへ追記する。置換が0件でも「0件」の行を書く。
' 入力  : runId、seriesId、settingId、prepared（PrepareComputationValuesの戻り値）
' 副作用: シートを作成・追記する。
Public Sub RecordReplacements(ByVal runId As String, ByVal seriesId As String, ByVal settingId As String, ByVal prepared As Object)
    Dim ws As Worksheet, rows As New Collection, item As Variant, startRow As Long, count As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.ZeroLogSheet())
    modX13IO.Headers ws, ZeroLogHeaders()
    For Each item In prepared("replaced")
        rows.Add Array(runId, seriesId, settingId, modX13IO.MonthText(item(0)), item(1), item(2), item(3), _
            "0を対数計算用にepsilonへ置換（" & REASON_ZERO_TO_EPSILON & "）", prepared("transform"), prepared("enabled"), prepared("epsilon"))
    Next item
    rows.Add Array(runId, seriesId, settingId, "(系列計)", "", "", "", _
        "置換" & prepared("replaced").Count & "件", prepared("transform"), prepared("enabled"), prepared("epsilon"))
    modX13IO.AppendRows ws, modX13IO.RowsToArray(rows, 11), startRow, count, Array(5, 6, 11)
End Sub

Public Function ZeroLogHeaders() As Variant
    ZeroLogHeaders = Array("実行ID", "系列ID", "設定ID", "年月", "元値_原データ", "計算値_DAT", "DAT書出文字列", _
        "理由", "変換", "ゼロ置換設定", "epsilon")
End Function

' 目的: 試行フォルダに保存する置換記録（タブ区切り）。置換0件でも見出しと件数行を書く。
Public Function ReplacementFileText(ByVal seriesId As String, ByVal settingId As String, ByVal prepared As Object) As String
    Dim item As Variant, body As String
    body = "series_id" & vbTab & "setting_id" & vbTab & "date" & vbTab & "original" & vbTab & "computed" & vbTab & _
        "dat_text" & vbTab & "reason" & vbCrLf
    For Each item In prepared("replaced")
        body = body & seriesId & vbTab & settingId & vbTab & modX13IO.MonthText(item(0)) & vbTab & _
            modX13IO.NumberText(item(1)) & vbTab & modX13IO.NumberText(item(2)) & vbTab & item(3) & vbTab & REASON_ZERO_TO_EPSILON & vbCrLf
    Next item
    body = body & "#replaced_count" & vbTab & prepared("replaced").Count & vbTab & "transform=" & prepared("transform") & _
        vbTab & "zero_replace=" & prepared("enabled") & vbTab & "epsilon=" & modX13IO.NumberText(prepared("epsilon")) & vbCrLf
    ReplacementFileText = body
End Function
