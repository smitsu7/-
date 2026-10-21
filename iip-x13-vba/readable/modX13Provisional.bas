Attribute VB_Name = "modX13Provisional"
Option Explicit

' 役割      : 最終決定した各系列（最終_保存履歴で「採用中」）について、予測月の暫定季節指数を2つの方法で作る。
'               方法1（前年同月）: 前年同月のD10（季節成分）と、当年のD18（カレンダー成分）を組み合わせる
'               方法2（予測）    : 当年のD10の予測値と、当年のD18を組み合わせる
'             組合せ方は季節調整の型で変える。
'               乗法型：暫定季節指数 = D10 × D18 × 100（指数表示。D10・D18は1前後の比率）
'               加法型：暫定季節指数 = D10 + D18（原指数と同じ単位の調整量）
'             カレンダー調整なしの採用では、D18はVBAが補った中立値（乗法1・加法0）で、その旨を出所の列に示す。
' 主な入口  : RefreshProvisional
' 依存先    : modX13Store（ValidateFinal・CurrentAdoptions）、modX13Results、modX13IO、modMain、modSheetNames、modPackagePaths、
'             modUI（系列1つごとの件数と経過時間の表示。長い処理の外では何もしない）
' 書く場所  : 最終_暫定季節指数（毎回作り直す）
' 計算例    : 乗法型、2026年1月：前年同月D10=0.9385、当年D10予測=0.9390、当年D18=1.0145 なら
'               方法1 = 0.9385×1.0145×100 = 95.21、方法2 = 0.9390×1.0145×100 = 95.26

Public Function ProvisionalHeaders() As Variant
    ProvisionalHeaders = Array("保存ID", "系列ID", "年月", "季節調整の型", "前年同月のD10", "当年のD10_予測", "当年のD18", _
        "方法1_前年同月D10とD18", "方法2_D10予測とD18", "差_方法2−方法1", "単位", "D18の出所", "出典ファイル")
End Function

' 目的  : 最終_暫定季節指数を作り直す。採用中の結果が無い系列は載せない。
' エラー: 最終保存の検査に失敗した系列は停止する（壊れた結果から暫定値を作らないため）。
'         乗法型・加法型以外は停止する。予測なし（予測0か月）の採用は行を作らない。
Public Sub RefreshProvisional()
    Dim ws As Worksheet, adoptions As Object, sid As Variant, meta As Object, base As String, values As Object
    Dim first As Long, actual As Long, forecast As Long, d10 As Variant, d18 As Variant, mode As String
    Dim rows As New Collection, k As Long, previous As Double, current As Double, calendar As Double
    Dim method1 As Double, method2 As Double, unitText As String, origin As String, startRow As Long, count As Long
    Dim done As Long
    Set ws = modMain.EnsureWorksheet(modSheetNames.ProvisionalSheet())
    modX13IO.ResetSheet ws, ProvisionalHeaders()
    Set adoptions = modX13Store.CurrentAdoptions()
    ' 採用中の系列ごとに最終保存を検査するため、系列が多いと時間がかかる。系列1つごとに件数と経過時間を出す。
    modUI.BeginStage "PROVISIONAL", "暫定季節指数の作り直し", adoptions.Count
    For Each sid In adoptions.Keys
        done = done + 1
        modUI.Progress done, CStr(sid)
        Set meta = modX13Store.ValidateFinal(CStr(adoptions(sid)))
        base = modMain.CombinePath(CStr(meta("FOLDER")), CStr(meta("STEM")))
        first = modX13IO.MonthIndexFromText(CStr(meta("FIRST")))
        actual = CLng(meta("ACTUAL_MONTHS"))
        forecast = CLng(meta("FORECAST_MONTHS"))
        If forecast > 0 Then
            d10 = modX13IO.ReadMonthlyTable(base & ".d10", first, actual + forecast)
            d18 = modX13IO.ReadMonthlyTable(base & ".d18", first, actual + forecast)
            Set values = modX13Results.ReadUdgValues(base & ".udg")
            mode = modX13Results.ReadValue(values, "samode")
            origin = IIf(CStr(meta("D18_ORIGIN")) = "VBA_NEUTRAL", "VBAの中立値（カレンダー調整なし）", "X-13の出力")
            For k = actual + 1 To actual + forecast
                previous = CDbl(d10(k - 12, 2))
                current = CDbl(d10(k, 2))
                calendar = CDbl(d18(k, 2))
                If InStr(mode, "multiplicative") > 0 Then
                    method1 = previous * calendar * 100
                    method2 = current * calendar * 100
                    unitText = "指数（季節成分×カレンダー成分×100）"
                ElseIf InStr(mode, "additive") > 0 Then
                    method1 = previous + calendar
                    method2 = current + calendar
                    unitText = "原指数と同じ単位（季節成分＋カレンダー成分）"
                Else
                    Err.Raise vbObjectError + 3600, "RefreshProvisional", "乗法型・加法型以外の暫定季節指数は作れません（samode=" & mode & "）：" & sid
                End If
                rows.Add Array(adoptions(sid), CStr(sid), modX13IO.MonthText(first + k - 1), mode, previous, current, calendar, _
                    method1, method2, method2 - method1, unitText, origin, _
                    modPackagePaths.RelativeToPackage(base & ".d10") & "; " & modPackagePaths.RelativeToPackage(base & ".d18"))
            Next k
        End If
    Next sid
    If rows.Count > 0 Then modX13IO.AppendRows ws, modX13IO.RowsToArray(rows, UBound(ProvisionalHeaders()) + 1), startRow, count, _
        Array(5, 6, 7, 8, 9, 10)
    ws.Columns.ColumnWidth = 16
    modUI.EndStage
End Sub
