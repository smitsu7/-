VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 役割      : ブックを開いた時と保存した後に、ボタンの参照先を現在のブック名へ付け替える。
'             X-13の実行・ネット取得・取込などの外部処理は、ここからは起動しない（利用者がボタンで始める）。
' 依存先    : modUI（RefreshActions）

Private Sub Workbook_Open()
    modUI.RefreshActions
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then modUI.RefreshActions
End Sub
