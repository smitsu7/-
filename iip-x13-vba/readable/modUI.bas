Attribute VB_Name = "modUI"
Option Explicit

' 役割      : 画面の共通処理。シート上のボタンの作成、ボタンの参照先の更新、日本語メッセージの表示、
'             自動試験のための表示抑止、長い処理の進捗表示・応答の維持・中断の受付・段階ごとの所要時間の記録を担当する。
'             計算や保存は行わない。
' 主な入口  : AddButton / RefreshActions / ShowMessage / Confirm / SetSilent / IsSilent
'             長い処理：BeginLongTask / EndLongTask / BeginStage / Progress / EndStage / Pulse
'                       CancelRequested / RaiseIfCancelled / AllowCancel / StartProgressFile / ProgressInfo
' 依存先    : modX13Native（MessageBoxW）、modLog、modX13IO（数値の文字列化）
' ボタン    : 名前は "btn" で始める。OnActionには「'ブック名'!モジュール.手続き」を入れる。
'             ブック名を変えて保存した場合は RefreshActions が新しいブック名へ付け替える。
' 長い処理  : 季節調整の実行・最終決定・系列別Excelの保存は、BeginLongTask から EndLongTask までを「長い処理」とする。
'             その間だけ、次のことを行う（開始していない時の Pulse・BeginStage 等は何もしない）。
'               1) ステータスバーに「処理名｜段階 件数｜この段階の秒数｜全体の経過時間」を出す
'               2) Pulse が呼ばれるたびに、前回から0.2秒以上たっていれば DoEvents で画面を更新する
'                  （Windowsは約5秒メッセージを処理しない画面を「応答なし」にして白く表示するため）
'               3) Application.Interactive=False にして、DoEvents の間にセルの編集・ボタン・ブックを閉じる操作が
'                  割り込まないようにする（二重実行と実行途中の入力変更を防ぐ）
'               4) Application.EnableCancelKey=xlDisabled にする。Ctrl+Break等でVBAが途中で止まると、
'                  起動したX-13・Excelの設定・実行中の印が残るため。中断はEscキーと確認画面で受け付ける（5）
'               5) Excelが手前にある時にEscキーが押されたら、確認のうえ中断を受け付ける。中断は、呼出側が
'                  区切りのよい所で RaiseIfCancelled を呼んで行う（保存の途中では止めない）
'               6) 段階ごとの件数と秒数を VBA_Log と、実行ごとの記録ファイル（StartProgressFile）へ書く
'             変えたExcelの設定（ステータスバー、その表示、カーソル、Interactive、EnableCancelKey）は、
'             EndLongTask で開始前の値へ戻す。成功・エラー・中断のどの場合も、呼出側が必ず EndLongTask を呼ぶ。
'             入れ子で呼んでもよい（最初の BeginLongTask で設定を記録し、対応する最後の EndLongTask で戻す）。

' 自動試験中は確認・通知の画面を出さない（ログには残す）。配布ブックの通常操作では常にFalse。
Private mSilent As Boolean
' 自動試験中の確認（Confirm）の答え。SetSilent の第2引数で決める（既定は「はい」）。
Private mSilentAnswer As Boolean

' MessageBoxWの戻り値（はい）と、表示の種類の組合せに使う値。
Private Const ANSWER_YES As Long = 6

' 長い処理：画面を更新する間隔（秒）と、記録ファイルへ途中経過を書く間隔（秒）。
Private Const PULSE_SECONDS As Double = 0.2
Private Const PROGRESS_LINE_SECONDS As Double = 30
' Escキーの仮想キーコード。
Private Const VK_ESCAPE As Long = &H1B

#If Mac Then
#Else
#If VBA7 Then
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal virtualKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
#End If
#End If

' 長い処理の状態。
Private mTaskDepth As Long              ' BeginLongTask の入れ子の深さ（0なら長い処理の外）
Private mTaskTitle As String            ' ステータスバーの先頭に出す処理名
Private mTaskStart As Double            ' 開始時の Timer の値（秒）
Private mLastPulse As Double            ' 最後に DoEvents した時の Timer の値
Private mLastFileLine As Double         ' 最後に記録ファイルへ書いた時の Timer の値
Private mPausedSeconds As Double        ' 長い処理の中で確認画面を開いていた秒数の合計（X-13の待ち時間の上限から除く）
Private mDialogStart As Double          ' 確認画面を開いた時の Timer の値
Private mStageId As String              ' 段階の記号（記録ファイル用。英数字）
Private mStageLabel As String           ' 段階の名前（画面用）
Private mStageStart As Double
Private mStageDone As Long
Private mStageTotal As Long
Private mStageDetail As String
Private mStageCounts As Object          ' 段階の記号 → これまでに始まった回数（試験用の中断の指定に使う）
Private mCancelRequested As Boolean     ' 利用者が中断を確認した（または試験で中断を指定した）
Private mCancelEnabled As Boolean       ' Falseの段階（最終保存・系列別Excel）は中断を受け付けない
Private mCancelRefusedShown As Boolean  ' 「この段階は中断できません」を表示したか
Private mCancelAtStage As String        ' 試験用：この段階が mCancelAtOccurrence 回目に始まったら（mCancelAtDone>0なら、
Private mCancelAtOccurrence As Long     '         その段階で mCancelAtDone 件が終わったら）中断を指定する
Private mCancelAtDone As Long
Private mProgressFile As String         ' 実行ごとの記録ファイル（空なら書かない）
Private mOldStatusBar As Variant
Private mOldDisplayStatusBar As Boolean
Private mOldCursor As Long
Private mOldInteractive As Boolean
Private mOldCancelKey As Long

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
' 補足  : 長い処理の途中で呼ばれた場合は、表示している間だけ操作を受け付ける（BeginDialog・EndDialog）。
Public Function ShowMessage(ByVal text As String, Optional ByVal flags As Long = vbInformation, _
                            Optional ByVal caption As String = "") As Long
    If Len(caption) = 0 Then caption = "IIP X-13 季節調整"
    If mSilent Then
        modLog.WriteLog "ShowMessage", "INFO", caption & "：" & text
        ShowMessage = vbOK
        Exit Function
    End If
    BeginDialog
    ShowMessage = modX13Native.UnicodeMessage(text, flags, caption)
    EndDialog
End Function

' 目的  : はい／いいえの確認。自動試験中は SetSilent で決めた答え（既定は「はい」）とみなす
'         （呼出側の試験コードが明示的に設定した場合だけ）。
Public Function Confirm(ByVal text As String, ByVal caption As String) As Boolean
    If mSilent Then
        modLog.WriteLog "Confirm", "INFO", caption & "：" & text & "（自動試験のため確認を省略。答え " & IIf(mSilentAnswer, "はい", "いいえ") & "）"
        Confirm = mSilentAnswer
        Exit Function
    End If
    BeginDialog
    Confirm = (modX13Native.UnicodeMessage(text, vbYesNo + vbQuestion, caption) = ANSWER_YES)
    EndDialog
End Function

' 目的: 自動試験用。画面の確認・通知を出さないようにする。試験の最後に必ずFalseへ戻す。
'       confirmAnswer は自動試験中の Confirm の答え（Falseは取消しの試験に使う）。省略すると「はい」。
Public Sub SetSilent(ByVal value As Boolean, Optional ByVal confirmAnswer As Boolean = True)
    mSilent = value
    mSilentAnswer = confirmAnswer
End Sub

Public Function IsSilent() As Boolean
    IsSilent = mSilent
End Function

' 目的  : 文字を入力してもらう。自動試験中は使わない（呼出側が値を直接渡す入口を持つ）。
' 出力  : 入力文字列。取消しは空文字。
Public Function AskText(ByVal prompt As String, ByVal caption As String, Optional ByVal defaultText As String = "") As String
    Dim answer As Variant
    If mSilent Then Err.Raise vbObjectError + 3801, "AskText", "自動試験中は入力画面を使えません：" & caption
    BeginDialog
    answer = Application.InputBox(prompt, caption, defaultText, Type:=2)
    EndDialog
    If VarType(answer) = vbBoolean Then Exit Function
    AskText = Trim$(CStr(answer))
End Function

' 長い処理の途中で画面を出す前後に呼ぶ。表示中だけ操作を受け付け、表示していた秒数を記録する。
Private Sub BeginDialog()
    If mTaskDepth = 0 Then Exit Sub
    mDialogStart = Timer
    Application.Interactive = True
    Application.Cursor = xlDefault
End Sub

Private Sub EndDialog()
    If mTaskDepth = 0 Then Exit Sub
    mPausedSeconds = mPausedSeconds + SecondsBetween(mDialogStart, Timer)
    Application.Cursor = xlWait
    Application.Interactive = False
End Sub

' ================= 長い処理の進捗・応答・中断 =================

' 目的  : 長い処理を始める。Excelの設定を記録してから、操作の受付・Ctrl+Break・カーソル・ステータスバーを切り替える。
' 入力  : title（ステータスバーの先頭に出す処理名。例 "季節調整（1系列）"）
' 副作用: 最初の呼出しだけ、Application.Interactive=False、EnableCancelKey=xlDisabled、Cursor=xlWait、
'         DisplayStatusBar=True にする。入れ子の呼出しは深さを数えるだけ。
' 注意  : 呼出側は、成功・エラー・中断のどの場合も EndLongTask を1回呼ぶこと。
Public Sub BeginLongTask(ByVal title As String)
    Dim ignored As Integer
    If mTaskDepth = 0 Then
#If Mac Then
#Else
        ' 開始より前に押されたEscを中断の操作と数えない（「前回の確認以後に押された」の印をここで消す）。
        ignored = GetAsyncKeyState(VK_ESCAPE)
#End If
        mOldStatusBar = Application.StatusBar
        mOldDisplayStatusBar = Application.DisplayStatusBar
        mOldCursor = Application.Cursor
        mOldInteractive = Application.Interactive
        mOldCancelKey = Application.EnableCancelKey
        mTaskTitle = title
        mTaskStart = Timer
        mLastPulse = mTaskStart
        mLastFileLine = mTaskStart
        mPausedSeconds = 0
        mStageId = ""
        mStageLabel = "開始"
        mStageStart = mTaskStart
        mStageDone = 0
        mStageTotal = 0
        mStageDetail = ""
        Set mStageCounts = CreateObject("Scripting.Dictionary")
        mCancelRequested = False
        mCancelEnabled = True
        mCancelRefusedShown = False
        mProgressFile = ""
        mTaskDepth = 1
        Application.DisplayStatusBar = True
        Application.Cursor = xlWait
        Application.EnableCancelKey = xlDisabled
        Application.Interactive = False
        Application.StatusBar = StatusText()
    Else
        mTaskDepth = mTaskDepth + 1
    End If
End Sub

' 目的  : 長い処理を終える。最後の呼出しで、BeginLongTask の前のExcelの設定へ戻し、中断の指定を消す。
' 入力  : outcome（記録ファイルの最後の行に書く結果。英数字。例 "OK"・"CANCELLED"・"ERROR 1234"）
'         "OK" のときは、終わっていない段階を終えたものとして記録する（エラー・中断のときは、止まった段階のまま残す）。
' 補足  : 設定を戻す途中で失敗しても、残りの設定は戻す。呼出側は、元のエラーを先に退避してから呼ぶ。
Public Sub EndLongTask(ByVal outcome As String)
    If mTaskDepth = 0 Then Exit Sub
    If mTaskDepth = 1 And outcome = "OK" Then EndStage
    mTaskDepth = mTaskDepth - 1
    If mTaskDepth > 0 Then Exit Sub
    WriteProgressLine "FINISH", outcome
    ' 後始末：Excelの設定を開始前の値へ戻す。1つ失敗しても残りを続ける。
    On Error Resume Next
    Application.Interactive = mOldInteractive
    Application.EnableCancelKey = mOldCancelKey
    Application.Cursor = mOldCursor
    Application.StatusBar = mOldStatusBar
    Application.DisplayStatusBar = mOldDisplayStatusBar
    On Error GoTo 0
    mCancelRequested = False
    mCancelAtStage = ""
    mCancelAtOccurrence = 0
    mCancelAtDone = 0
    mProgressFile = ""
    mStageId = ""
End Sub

' 目的: 長い処理の途中か（試験と、Pulse を呼ぶか迷う呼出側のため）。
Public Function IsLongTaskActive() As Boolean
    IsLongTaskActive = (mTaskDepth > 0)
End Function

' 目的  : 段階を始める。前の段階が終わっていなければ、先に終えたものとして記録する。
' 入力  : stageId（記録ファイルに書く英数字の記号。例 "X13_RUN"）、label（画面の名前）、total（件数。分からなければ0）
' 副作用: ステータスバーを更新し、記録ファイルに BEGIN の行を書く。
Public Sub BeginStage(ByVal stageId As String, ByVal label As String, ByVal total As Long)
    If mTaskDepth = 0 Then Exit Sub
    If Len(mStageId) > 0 Then EndStage
    mStageId = stageId
    mStageLabel = label
    mStageTotal = total
    mStageDone = 0
    mStageDetail = ""
    mStageStart = Timer
    mStageCounts(stageId) = mStageCounts(stageId) + 1
    WriteProgressLine "BEGIN", ""
    CheckTestCancel
    Application.StatusBar = StatusText()
    mLastPulse = 0
    Pulse
End Sub

' 目的  : 段階の中の進み具合（done件目まで終わった）を伝え、必要なら画面を更新する。
' 入力  : done（終わった件数）、detail（画面に添える短い文字。例 系列ID・候補ID）
Public Sub Progress(ByVal done As Long, Optional ByVal detail As String = "")
    If mTaskDepth = 0 Then Exit Sub
    mStageDone = done
    mStageDetail = detail
    CheckTestCancel
    Pulse
End Sub

' 試験用の中断の指定（RequestCancelAtStage）に当たれば、Escで中断を確認したのと同じ状態にする。
Private Sub CheckTestCancel()
    If Not mCancelEnabled Or mCancelRequested Or Len(mCancelAtStage) = 0 Then Exit Sub
    If mCancelAtStage <> mStageId Or mStageCounts(mStageId) <> mCancelAtOccurrence Then Exit Sub
    If mStageDone < mCancelAtDone Then Exit Sub
    mCancelRequested = True
    WriteProgressLine "CANCEL_REQUESTED", "test"
End Sub

' 目的  : 段階を終える。件数と秒数を、記録ファイルの END の行と VBA_Log に書く。
' 入力  : note（記録ファイルに添える英数字の短い説明。例 "exit=0"）
Public Sub EndStage(Optional ByVal note As String = "")
    If mTaskDepth = 0 Or Len(mStageId) = 0 Then Exit Sub
    WriteProgressLine "END", note
    modLog.WriteLog "Progress", "INFO", mTaskTitle & "：" & mStageLabel & " " & mStageDone & "／" & mStageTotal & "件、" & _
        DurationText(SecondsBetween(mStageStart, Timer)) & IIf(Len(note) > 0, "（" & note & "）", "")
    mStageId = ""
End Sub

' 目的  : 長い処理の途中で、前回から0.2秒以上たっていれば、ステータスバーを更新して DoEvents を呼ぶ。
'         Escキーが押されていれば中断を確認する。30秒ごとに記録ファイルへ途中経過を書く。
' 補足  : 長い処理の外では何もしない（DoEventsも呼ばない）。何度呼んでもよい（間隔は時計で決める）。
Public Sub Pulse()
    Dim nowValue As Double
    If mTaskDepth = 0 Then Exit Sub
    nowValue = Timer
    If SecondsBetween(mLastPulse, nowValue) < PULSE_SECONDS Then Exit Sub
    mLastPulse = nowValue
    Application.StatusBar = StatusText()
    DoEvents
    If EscapePressed() Then
        If mCancelEnabled Then
            If Not mCancelRequested Then AskCancel
        ElseIf Not mCancelRefusedShown Then
            mCancelRefusedShown = True
            ShowMessage "最終保存・系列別Excelの保存の段階は中断できません。終わるまでお待ちください。", vbInformation, "中断"
        End If
    End If
    If SecondsBetween(mLastFileLine, nowValue) >= PROGRESS_LINE_SECONDS Then WriteProgressLine "PROGRESS", ""
End Sub

' Escキーで中断を確認する。「はい」なら中断を指定する（実際に止めるのは呼出側の区切り）。
Private Sub AskCancel()
    If Confirm("計算を中断しますか？" & vbLf & _
        "このブックが起動したX-13だけを止め、区切りのよい所（候補の保存の途中ではない所）で終わります。" & vbLf & _
        "中断した実行は、候補比較・最終決定に使いません。", "中断の確認") Then
        mCancelRequested = True
        WriteProgressLine "CANCEL_REQUESTED", "user"
        Application.StatusBar = StatusText()
    End If
End Sub

' Excelが手前にあり、Escキーが押された（または前回の確認以後に押された）か。自動試験中は見ない。
Private Function EscapePressed() As Boolean
#If Mac Then
    EscapePressed = False
#Else
    If mSilent Then Exit Function
    If (GetAsyncKeyState(VK_ESCAPE) And &H8001) = 0 Then Exit Function
    EscapePressed = (GetForegroundWindow() = Application.hwnd)
#End If
End Function

' 目的: 中断が指定されているか。
Public Function CancelRequested() As Boolean
    CancelRequested = mCancelRequested
End Function

' 目的: 中断を表すエラー番号（呼出側が中断とそれ以外のエラーを見分けるため）。
Public Function CancelErrorNumber() As Long
    CancelErrorNumber = vbObjectError + 3810
End Function

' 目的  : 中断が指定されていれば、中断のエラーを送出する。呼出側は、保存の途中ではない区切りでだけ呼ぶ。
' エラー: CancelErrorNumber のエラー。
Public Sub RaiseIfCancelled()
    If mCancelRequested Then Err.Raise CancelErrorNumber(), "RaiseIfCancelled", "利用者の操作で中断しました。"
End Sub

' 目的  : 中断を受け付けるかを切り替える。最終保存・系列別Excelの段階は False にする（途中で止めると、
'         一部の系列だけが採用される・ファイルが欠けるため）。BeginLongTask の時点では True。
Public Sub AllowCancel(ByVal value As Boolean)
    If mTaskDepth = 0 Then Exit Sub
    mCancelEnabled = value
    Application.StatusBar = StatusText()
End Sub

' 目的: 試験用。stageId の段階が occurrence 回目に始まった時（afterDone>0なら、その段階で afterDone 件が終わった時）に、
'       Escで中断を確認したのと同じ状態にする。中断を受け付けない段階では何もしない。
'       次の EndLongTask で消える。配布ブックの通常操作では使わない（自己試験 modR11SelfTests だけが呼ぶ）。
Public Sub RequestCancelAtStage(ByVal stageId As String, ByVal occurrence As Long, Optional ByVal afterDone As Long = 0)
    mCancelAtStage = stageId
    mCancelAtOccurrence = occurrence
    mCancelAtDone = afterDone
End Sub

' 目的: 長い処理の中で確認画面を開いていた秒数の合計（X-13の待ち時間の上限に数えないため）。
Public Function PausedSeconds() As Double
    PausedSeconds = mPausedSeconds
End Function

' 目的: 今の段階の名前（エラーの記録に使う）。長い処理の外では空文字。
Public Function CurrentStageLabel() As String
    If mTaskDepth > 0 Then CurrentStageLabel = mStageLabel
End Function

' 目的: 秒数を「1時間02分」「3分05秒」「12.3秒」の形にする。
Public Function DurationText(ByVal seconds As Double) As String
    Dim whole As Long
    If seconds < 60 Then
        DurationText = modX13IO.NumberText(Round(seconds, 1)) & "秒"
        Exit Function
    End If
    whole = CLng(Int(seconds))
    If whole >= 3600 Then
        DurationText = (whole \ 3600) & "時間" & Format$((whole Mod 3600) \ 60, "00") & "分"
    Else
        DurationText = (whole \ 60) & "分" & Format$(whole Mod 60, "00") & "秒"
    End If
End Function

' ステータスバーの文字。例「季節調整（1系列）｜再推定の検査・保存 12／29件｜この段階 35.2秒｜経過 2分10秒｜Escで中断」
Private Function StatusText() As String
    Dim text As String
    text = mTaskTitle & "｜" & mStageLabel
    If mStageTotal > 0 Then text = text & " " & mStageDone & "／" & mStageTotal & "件"
    If Len(mStageDetail) > 0 Then text = text & "（" & mStageDetail & "）"
    text = text & "｜この段階 " & DurationText(SecondsBetween(mStageStart, Timer)) & "｜経過 " & DurationText(SecondsBetween(mTaskStart, Timer))
    If mCancelRequested Then
        text = text & "｜中断しています…"
    ElseIf mCancelEnabled Then
        text = text & "｜Escで中断"
    Else
        text = text & "｜この段階は中断できません"
    End If
    StatusText = text
End Function

' Timer（0時からの秒数）の差。0時をまたいだ場合は1日分を足す。
Private Function SecondsBetween(ByVal fromValue As Double, ByVal toValue As Double) As Double
    SecondsBetween = toValue - fromValue
    If SecondsBetween < 0 Then SecondsBetween = SecondsBetween + 86400#
End Function

' ---- 実行ごとの記録ファイル ----
' 形式  : 英数字だけのタブ区切り。「#」で始まる見出し、INFO の行（環境と設定）、段階の行
'         （経過秒、出来事 BEGIN／END／PROGRESS／CANCEL_REQUESTED／FAIL／FINISH、段階の記号、件数、全件数、段階の秒数、補足）。
'         原指数の値・利用者名・フォルダの場所は書かない（そのまま共有して止まった段階を説明できるように）。
'         1行ごとに開いて追記して閉じるため、Excelが応答しなくなった場合も、それまでの行はファイルに残る。

' 目的  : 記録ファイルを決めて見出しを書く。長い処理の外では何もしない。
' 入力  : path（英数字だけのパス。X13_WORK_DIR の中を想定）
' エラー: 書けない場合は、記録ファイルを使わずに続ける（VBA_Logに残す）。計算は止めない。
Public Sub StartProgressFile(ByVal path As String)
    If mTaskDepth = 0 Then Exit Sub
    mProgressFile = path
    AppendProgressText "# IIP X-13 r11 progress record. Stage names, counts and seconds only (no index values, user names or folders)."
    AppendProgressText "# columns: elapsed_s<TAB>event<TAB>stage<TAB>done<TAB>total<TAB>stage_s<TAB>detail"
    ProgressInfo "office_bits", OfficeBits()
    ProgressInfo "excel_version", Application.Version & " build " & Application.Build
    ProgressInfo "os", Application.OperatingSystem
End Sub

' 目的: 記録ファイルに INFO の行（キーと値）を書く。値の英数字以外は「?」にする。
Public Sub ProgressInfo(ByVal key As String, ByVal value As String)
    If mTaskDepth = 0 Then Exit Sub
    AppendProgressText "INFO" & vbTab & AsciiText(key) & vbTab & AsciiText(value)
End Sub

' 目的: 記録ファイルに FAIL の行を書く（エラーで止まった段階を残すため）。errorNumber だけを書き、文言は VBA_Log に残す。
Public Sub ProgressFailure(ByVal errorNumber As Long)
    If mTaskDepth = 0 Then Exit Sub
    WriteProgressLine "FAIL", "error " & errorNumber
End Sub

Private Sub WriteProgressLine(ByVal eventName As String, ByVal detail As String)
    If Len(mProgressFile) = 0 Then Exit Sub
    mLastFileLine = Timer
    AppendProgressText modX13IO.NumberText(Round(SecondsBetween(mTaskStart, Timer), 1)) & vbTab & eventName & vbTab & mStageId & vbTab & _
        mStageDone & vbTab & mStageTotal & vbTab & modX13IO.NumberText(Round(SecondsBetween(mStageStart, Timer), 1)) & vbTab & AsciiText(detail)
End Sub

' 記録ファイルへ1行を追記する。失敗したら記録ファイルの使用をやめ、VBA_Logに1回だけ残す。
Private Sub AppendProgressText(ByVal line As String)
    Dim h As Integer, errorText As String
    If Len(mProgressFile) = 0 Then Exit Sub
    On Error GoTo Failed
    h = FreeFile
    Open mProgressFile For Append As #h
    Print #h, line
    Close #h
    Exit Sub
Failed:
    errorText = Err.Description
    ' 後始末：開いたファイル番号だけを閉じる。
    On Error Resume Next
    If h > 0 Then Close #h
    On Error GoTo 0
    modLog.WriteLog "Progress", "WARN", "進捗の記録ファイルに書けないため、記録をやめます：" & mProgressFile & " " & errorText
    mProgressFile = ""
End Sub

' 英数字・記号（32～126）以外とタブを「?」・空白にする（記録ファイルをどの言語のWindowsでも同じに読めるように）。
Private Function AsciiText(ByVal text As String) As String
    Dim i As Long, code As Long
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code = 9 Then
            Mid$(text, i, 1) = " "
        ElseIf code < 32 Or code > 126 Then
            Mid$(text, i, 1) = "?"
        End If
    Next i
    AsciiText = text
End Function

' Officeの版（32／64ビット）。社用PCでの状況を記録ファイルから分かるようにするため。
Private Function OfficeBits() As String
#If Win64 Then
    OfficeBits = "64"
#Else
    OfficeBits = "32"
#End If
End Function
