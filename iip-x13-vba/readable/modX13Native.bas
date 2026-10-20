Attribute VB_Name = "modX13Native"
Option Explicit

' 役割      : Windows APIで公式X-13を1プロセスだけ起動し、終了コードと待ち時間の上限を確かめる。
'             PowerShellやコマンドプロンプトを経由しない。実行ポリシー・VBAの信頼設定は変更しない。
'             他のプロセスは終了しない（時間切れの場合も、自分が起動したX-13だけを止める）。
'             あわせて、日本語を正しく表示するメッセージボックス（MessageBoxW）を提供する。
' 主な入口  : RunBatch / UnicodeMessage
' 依存先    : modMain（実行許可）、modX13IO（パス検査）
' 起動方法  : 作業フォルダ（jobs.mta のあるフォルダ）を現在のフォルダとして
'             "<x13as_ascii.exe>" -m jobs -s -q を実行する。exeのパスは引用符で囲み、
'             日本語・空白を含む場所でも起動できる（CreateProcessWはUnicodeでパスを受け取る）。
' 記録      : runner.status.txt に状態（RUNNING／COMPLETED／TIMED_OUT／RUNNER_FAILED）と終了コードを書く。
' 出典      : r10の modX13Native をそのまま引き継いだ（構造体の大きさの検査を含む）。

#If Mac Then
#Else
#If VBA7 Then
Private Type STARTUPINFO_NATIVE
    cb As Long
#If Win64 Then
    pad1 As Long
#End If
    reserved As LongPtr
    desktop As LongPtr
    title As LongPtr
    x As Long
    y As Long
    xSize As Long
    ySize As Long
    xCount As Long
    yCount As Long
    fill As Long
    flags As Long
    showWindow As Integer
    reservedSize As Integer
#If Win64 Then
    pad2 As Long
#End If
    reservedBuffer As LongPtr
    stdInput As LongPtr
    stdOutput As LongPtr
    stdError As LongPtr
End Type
Private Type PROCESSINFO_NATIVE
    processHandle As LongPtr
    threadHandle As LongPtr
    processId As Long
    threadId As Long
End Type
Private Declare PtrSafe Function CreateProcessW Lib "kernel32" ( _
    ByVal applicationName As LongPtr, ByVal commandLine As LongPtr, _
    ByVal processSecurity As LongPtr, ByVal threadSecurity As LongPtr, _
    ByVal inheritHandles As Long, ByVal creationFlags As Long, _
    ByVal environment As LongPtr, ByVal directory As LongPtr, _
    ByRef startupInfo As STARTUPINFO_NATIVE, ByRef processInfo As PROCESSINFO_NATIVE) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal handle As LongPtr, ByVal milliseconds As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal handle As LongPtr, ByRef exitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal handle As LongPtr, ByVal exitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal text As LongPtr, _
    ByVal caption As LongPtr, ByVal flags As Long) As Long
#End If
#End If

' 待機の単位（ミリ秒）と、WaitForSingleObjectの戻り値。
Private Const WAIT_SLICE_MS As Long = 50
Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_TIMEOUT As Long = 258
' 時間切れで止めた場合の終了コード（X-13自身の終了コードと区別するため124を使う）。
Private Const EXIT_TIMED_OUT As Long = 124
Private Const EXIT_RUNNER_FAILED As Long = 125
' CREATE_NO_WINDOW。コンソール画面を出さない。
Private Const CREATE_NO_WINDOW As Long = &H8000000

' 目的  : 日本語を含むメッセージを表示する（VBAのMsgBoxは環境によって文字化けするため）。
Public Function UnicodeMessage(ByVal text As String, ByVal flags As Long, ByVal caption As String) As Long
#If Mac Then
    UnicodeMessage = VBA.MsgBox(text, flags, caption)
#Else
#If VBA7 Then
    UnicodeMessage = MessageBoxW(Application.hwnd, StrPtr(text), StrPtr(caption), flags)
#Else
    UnicodeMessage = VBA.MsgBox(text, flags, caption)
#End If
#End If
End Function

' 目的  : 作業フォルダの jobs.mta を、X-13の1プロセスで一括実行する。
' 入力  : executable（x13as_ascii.exeの絶対パス）、folder（jobs.mtaのあるバッチフォルダ）、timeoutSeconds（1～3600）
' 出力  : X-13の終了コード（時間切れは124）
' 副作用: folder に runner.status.txt を書く。X-13が各ジョブの出力ファイルと jobs.log を書く。
' エラー: 実行許可が無い、ファイルが無い、既に実行済みのフォルダ、起動失敗は停止する。
Public Function RunBatch(ByVal executable As String, ByVal folder As String, ByVal timeoutSeconds As Long) As Long
#If Mac Then
    Err.Raise vbObjectError + 1770, "RunBatch", "X-13の実行にはWindows版Excelが必要です。"
#Else
#If VBA7 Then
    Dim si As STARTUPINFO_NATIVE, pi As PROCESSINFO_NATIVE, fso As Object
    Dim command As String, statusPath As String, started As Double, elapsed As Double
    Dim waitResult As Long, exitCode As Long, errorNumber As Long, errorText As String
    Dim owned As Boolean, exited As Boolean, dllError As Long
    On Error GoTo Failed
    If Not modMain.GetControlBoolean("ALLOW_X13_EXECUTION", False) Then Err.Raise vbObjectError + 1784, , "ControlのALLOW_X13_EXECUTIONがYESではありません。"
    If timeoutSeconds < 1 Or timeoutSeconds > 3600 Then Err.Raise vbObjectError + 1771, , "待ち時間の上限は1～3600秒です。"
    Set fso = CreateObject("Scripting.FileSystemObject")
    modX13IO.RequireLocalPath executable
    modX13IO.RequireLocalPath folder
    If Not fso.FileExists(executable) Or LCase$(fso.GetExtensionName(executable)) <> "exe" Then _
        Err.Raise vbObjectError + 1772, , "X-13の実行ファイル（.exe）がありません：" & executable
    If Not fso.FolderExists(folder) Or Not fso.FileExists(fso.BuildPath(folder, "jobs.mta")) Then _
        Err.Raise vbObjectError + 1773, , "バッチの入力（jobs.mta）がありません：" & folder
    If InStr(executable, Chr$(34)) > 0 Or InStr(executable, vbCr) > 0 Or InStr(executable, vbLf) > 0 Or InStr(executable, vbNullChar) > 0 Then _
        Err.Raise vbObjectError + 1774, , "実行ファイルのパスに使えない文字があります。"
    statusPath = fso.BuildPath(folder, "runner.status.txt")
    If fso.FileExists(statusPath) Or fso.FileExists(fso.BuildPath(folder, "jobs.log")) Then _
        Err.Raise vbObjectError + 1775, , "このバッチフォルダは既に実行されています：" & folder
    si.cb = LenB(si)
#If Win64 Then
    If si.cb <> 104 Or LenB(pi) <> 24 Then Err.Raise vbObjectError + 1776, , "64ビット版の構造体の大きさが想定と違います。"
#Else
    If si.cb <> 68 Or LenB(pi) <> 16 Then Err.Raise vbObjectError + 1776, , "32ビット版の構造体の大きさが想定と違います。"
#End If
    command = Chr$(34) & executable & Chr$(34) & " -m jobs -s -q" & vbNullChar
    owned = True
    WriteStatus statusPath, "RUNNING", -1, 0
    If CreateProcessW(StrPtr(executable), StrPtr(command), 0, 0, 0, CREATE_NO_WINDOW, 0, StrPtr(folder), si, pi) = 0 Then
        dllError = Err.LastDllError
        Err.Raise vbObjectError + 1777, , "X-13を起動できません（CreateProcessW " & dllError & "）。"
    End If
    started = TickValue(GetTickCount())
    Do
        waitResult = WaitForSingleObject(pi.processHandle, WAIT_SLICE_MS)
        If waitResult = WAIT_OBJECT_0 Then Exit Do
        If waitResult <> WAIT_TIMEOUT Then Err.Raise vbObjectError + 1778, , "X-13の終了待ちに失敗しました。"
        elapsed = TickValue(GetTickCount()) - started
        If elapsed < 0 Then elapsed = elapsed + 4294967296#
        If elapsed >= CDbl(timeoutSeconds) * 1000# Then
            If TerminateProcess(pi.processHandle, EXIT_TIMED_OUT) = 0 Then Err.Raise vbObjectError + 1779, , "時間切れのX-13を止められません。"
            If WaitForSingleObject(pi.processHandle, 5000) <> WAIT_OBJECT_0 Then Err.Raise vbObjectError + 1780, , "時間切れのX-13が終了しません。"
            exited = True
            WriteStatus statusPath, "TIMED_OUT", EXIT_TIMED_OUT, pi.processId
            RunBatch = EXIT_TIMED_OUT
            GoTo Done
        End If
        DoEvents
    Loop
    exited = True
    If GetExitCodeProcess(pi.processHandle, exitCode) = 0 Then Err.Raise vbObjectError + 1781, , "X-13の終了コードを取得できません。"
    WriteStatus statusPath, "COMPLETED", exitCode, pi.processId
    RunBatch = exitCode
Done:
    If pi.threadHandle <> 0 Then CloseHandle pi.threadHandle
    If pi.processHandle <> 0 Then CloseHandle pi.processHandle
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    ' 後始末：自分が起動したX-13だけを止め、状態を記録し、ハンドルを閉じる。後始末の失敗で元のエラーを隠さない。
    On Error Resume Next
    If pi.processHandle <> 0 And Not exited Then
        TerminateProcess pi.processHandle, EXIT_RUNNER_FAILED
        WaitForSingleObject pi.processHandle, 5000
    End If
    If owned Then WriteStatus statusPath, "RUNNER_FAILED", EXIT_RUNNER_FAILED, pi.processId
    If pi.threadHandle <> 0 Then CloseHandle pi.threadHandle
    If pi.processHandle <> 0 Then CloseHandle pi.processHandle
    On Error GoTo 0
    Err.Raise errorNumber, "modX13Native.RunBatch", errorText
#Else
    Err.Raise vbObjectError + 1782, "RunBatch", "X-13の実行にはVBA7以降（Excel 2010以降）が必要です。"
#End If
#End If
End Function

' GetTickCountの符号付き値を、0～2^32-1の経過ミリ秒にする。
Private Function TickValue(ByVal value As Long) As Double
    If value < 0 Then TickValue = CDbl(value) + 4294967296# Else TickValue = CDbl(value)
End Function

' runner.status.txt を書く（実行の度に上書きする小さな状態ファイル）。
Private Sub WriteStatus(ByVal path As String, ByVal state As String, ByVal code As Long, ByVal pid As Long)
    Dim h As Integer, errorNumber As Long, errorText As String
    On Error GoTo Failed
    h = FreeFile
    Open path For Output As #h
    Print #h, state
    Print #h, "EXIT_CODE=" & CStr(code)
    Print #h, "PROCESS_ID=" & CStr(pid)
    Print #h, "BACKEND=NATIVE"
    Close #h
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    ' 後始末：開いたファイル番号だけを閉じる。
    On Error Resume Next
    If h > 0 Then Close #h
    On Error GoTo 0
    Err.Raise errorNumber, "modX13Native.WriteStatus", errorText
End Sub

' 目的: runner.status.txt の1行目（状態）を読む。ファイルが無ければ "MISSING"。
Public Function ReadStatus(ByVal folder As String) As String
    Dim path As String, text As String
    path = modMain.CombinePath(folder, "runner.status.txt")
    If Not modMain.FileExists(path) Then
        ReadStatus = "MISSING"
        Exit Function
    End If
    text = Replace(modX13IO.ReadText(path, "UTF8"), vbCr, "")
    ReadStatus = Trim$(Split(text, vbLf)(0))
End Function
