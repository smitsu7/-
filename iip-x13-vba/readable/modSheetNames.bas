Attribute VB_Name = "modSheetNames"
Option Explicit

' 役割      : ブック内のシート名を一か所で定義する。名前の先頭で役割を区別する。
'             原データ_ = 取り込んだ原本の値と来歴（計算で書き換えない）
'             計算用_   = 原本から作った計算用の値と、その作り方の記録
'             設定_     = 利用者が編集する設定表
'             試行_     = 比較候補の計算結果（最終結果ではない）
'             最終_     = 最終モデルとして保存した結果
'             色は補助で、識別はこの名前と各表の列名で行う。
' 主な入口  : 各関数がシート名を返す。
' 依存先    : なし。
' 補足      : 日本語の名前をConstにしないのは、ASCII版の取込用ソース（import/ascii）で
'             文字列を関数呼出しに置き換えるため。Constには関数を書けない。

Public Function HomeSheet() As String
    HomeSheet = "季節調整"
End Function

Public Function ControlSheet() As String
    ControlSheet = "Control"
End Function

Public Function ImportConfigSheet() As String
    ImportConfigSheet = "IIP_取込設定"
End Function

Public Function ImportPreviewSheet() As String
    ImportPreviewSheet = "IIP_取込候補"
End Function

' 原データ：取り込んだ原指数。1行1観測値で、元ファイル・シート・セルを同じ行に持つ。
Public Function RawIndexSheet() As String
    RawIndexSheet = "原データ_原指数"
End Function

' 原データ：取り込んだ系列だけの一覧（名称・種類・期間内の月数・状態IMPORTED_・取込ID）。
' 取り込む前の候補（期間不足の系列を含む）は IIP_取込候補 に載り、ここには載らない。
Public Function SeriesListSheet() As String
    SeriesListSheet = "原データ_系列一覧"
End Function

' 原データ：取り込んだファイルごとのSHA-256・大きさ・出典。
Public Function ProvenanceSheet() As String
    ProvenanceSheet = "原データ_来歴"
End Function

' 内閣府CSVの日付一覧（HolidayCalendar相当）。
Public Function HolidayCalendarSheet() As String
    HolidayCalendarSheet = "原データ_祝日CSV"
End Function

' 年月・A・B・A-B・実績/予測の計算表（HolidayCalculation相当）。
Public Function HolidayCalculationSheet() As String
    HolidayCalculationSheet = "計算用_祝日変数"
End Function

Public Function ZeroLogSheet() As String
    ZeroLogSheet = "計算用_ゼロ置換記録"
End Function

Public Function SeriesSelectionSheet() As String
    SeriesSelectionSheet = "設定_対象系列"
End Function

Public Function SeriesSettingsSheet() As String
    SeriesSettingsSheet = "設定_系列別"
End Function

Public Function InterventionSheet() As String
    InterventionSheet = "設定_外れ値指定"
End Function

' 実行ごとにツールが作った比較候補の定義（カレンダー変数、次数、外れ値の扱い）。利用者は編集しない。
Public Function CandidateSheet() As String
    CandidateSheet = "試行_候補定義"
End Function

' X-13へ渡した各ジョブの状態と失敗理由。
Public Function JobSheet() As String
    JobSheet = "試行_ジョブ"
End Function

' 保存が完了した比較候補（試行）の登録簿。保存先は04_results/candidates。
Public Function TrialSheet() As String
    TrialSheet = "試行_一覧"
End Function

Public Function ComparisonSheet() As String
    ComparisonSheet = "試行_候補比較"
End Function

Public Function RegressionSheet() As String
    RegressionSheet = "試行_変数検定"
End Function

Public Function SearchHistorySheet() As String
    SearchHistorySheet = "試行_探索履歴"
End Function

Public Function ExecutionRecordSheet() As String
    ExecutionRecordSheet = "試行_実行記録"
End Function

Public Function DetectedOutlierSheet() As String
    DetectedOutlierSheet = "試行_検出外れ値"
End Function

Public Function DecisionSheet() As String
    DecisionSheet = "最終_決定"
End Function

Public Function AdoptionSheet() As String
    AdoptionSheet = "最終_保存履歴"
End Function

Public Function FinalMonthlySheet() As String
    FinalMonthlySheet = "最終_月次結果"
End Function

Public Function FinalSettingsSheet() As String
    FinalSettingsSheet = "最終_使用設定"
End Function

Public Function ProvisionalSheet() As String
    ProvisionalSheet = "最終_暫定季節指数"
End Function

Public Function LogSheet() As String
    LogSheet = "VBA_Log"
End Function
