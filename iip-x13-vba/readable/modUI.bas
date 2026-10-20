Attribute VB_Name = "modUI"
Option Explicit

' 役割      : 画面の共通処理。シート上のボタンの作成、ボタンの参照先の更新、日本語メッセージの表示、
'             自動試験のための表示抑止を担当する。計算や保存は行わない。
' 主な入口  : AddButton / RefreshActions / ShowMessage / Confirm / SetSilent / IsSilent
' 依存先    : modX13Native（MessageBoxW）、modLog
' ボタン    : 名前は "btn" で始める。OnActionには「'ブック名'!モジュール.手続き」を入れる。
'             ブック名を変えて保存した場合は RefreshActions が新しいブック名へ付け替える。

' 自動試験中は確認・通知の画面を出さない（ログには残す）。配布ブックの通常操作では常にFalse。
Private mSilent As Boolean

' MessageBoxWの戻り値（はい）と、表示の種類の組合せに使う値。
Private Const ANSWER_YES As Long = 6

' 目的  : シートにボタン（四角形の図形にマクロを割り当てたもの）を置く。同名のボタンがあれば作り直す。
' 入力  : ws、buttonName（"btn"で始まる英数字）、caption（表示名）、macroName（"modX.Proc"）、anchor（左上のセル）
' 副作用: シートに図形を作る。r10で動作を確かめた方式（AddShape＋TextFrame2）と同じ。
Public Sub AddButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal caption As String, _
                     ByVal macroName As String, ByVal anchor As Range, Optional ByVal widthPoints As Double = 150)
    Dim shapeItem As Shape
    If Left$(buttonName, 3) <> "btn" Then Err.Raise vbObjectError + 3800, "AddButton", "ボタンの名前は btn で始めてください：" & buttonName
    For Each shapeItem In ws.Shapes
        If shapeItem.Name = buttonName Then
            shapeItem.Delete
            Exit For
        End If
    Next shapeItem
    Set shapeItem = ws.Shapes.AddShape(msoShapeRectangle, anchor.Left + 2, anchor.Top + 2, widthPoints, 22)
    shapeItem.Name = buttonName
    shapeItem.TextFrame2.TextRange.Text = caption
    shapeItem.OnAction = QualifiedMacro(macroName)
    shapeItem.Placement = xlFreeFloating
End Sub

' 目的  : 全シートのボタンの参照先を、現在のブック名へ付け替える（名前を変えて保存した後のため）。
' 副作用: ボタンのOnActionだけを書き換える。マクロは実行しない。
Public Sub RefreshActions()
    Dim ws As Worksheet, shapeItem As Shape, action As String, pos As Long
    For Each ws In ThisWorkbook.Worksheets
        For Each shapeItem In ws.Shapes
            If Left$(shapeItem.Name, 3) = "btn" Then
                action = shapeItem.OnAction
                pos = InStrRev(action, "!")
                If pos > 0 Then action = Mid$(action, pos + 1)
                If Len(action) > 0 Then shapeItem.OnAction = QualifiedMacro(action)
            End If
        Next shapeItem
    Next ws
End Sub

Private Function QualifiedMacro(ByVal macroName As String) As String
    QualifiedMacro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

' 目的  : 日本語のメッセージを表示する。自動試験中は表示せずログへ書く。
' 出力  : 押されたボタンの値（vbOK等）。自動試験中は vbOK。
Public Function ShowMessage(ByVal text As String, Optional ByVal flags As Long = vbInformation, _
                            Optional ByVal caption As String = "") As Long
    If Len(caption) = 0 Then caption = "IIP X-13 季節調整"
    If mSilent Then
        modLog.WriteLog "ShowMessage", "INFO", caption & "：" & text
        ShowMessage = vbOK
        Exit Function
    End If
    ShowMessage = modX13Native.UnicodeMessage(text, flags, caption)
End Function

' 目的  : はい／いいえの確認。自動試験中は「はい」とみなす（呼出側の試験コードが明示的に設定した場合だけ）。
Public Function Confirm(ByVal text As String, ByVal caption As String) As Boolean
    If mSilent Then
        modLog.WriteLog "Confirm", "INFO", caption & "：" & text & "（自動試験のため確認を省略）"
        Confirm = True
        Exit Function
    End If
    Confirm = (modX13Native.UnicodeMessage(text, vbYesNo + vbQuestion, caption) = ANSWER_YES)
End Function

' 目的: 自動試験用。画面の確認・通知を出さないようにする。試験の最後に必ずFalseへ戻す。
Public Sub SetSilent(ByVal value As Boolean)
    mSilent = value
End Sub

Public Function IsSilent() As Boolean
    IsSilent = mSilent
End Function

' 目的  : 文字を入力してもらう。自動試験中は使わない（呼出側が値を直接渡す入口を持つ）。
' 出力  : 入力文字列。取消しは空文字。
Public Function AskText(ByVal prompt As String, ByVal caption As String, Optional ByVal defaultText As String = "") As String
    Dim answer As Variant
    If mSilent Then Err.Raise vbObjectError + 3801, "AskText", "自動試験中は入力画面を使えません：" & caption
    answer = Application.InputBox(prompt, caption, defaultText, Type:=2)
    If VarType(answer) = vbBoolean Then Exit Function
    AskText = Trim$(CStr(answer))
End Function
