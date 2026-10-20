Attribute VB_Name = "modPackagePaths"
Option Explicit

' 役割      : パッケージ（このブックがあるフォルダ）を基準に、エンジン・原指数・祝日・結果の場所を決める。
'             Controlには相対パス（例 01_engine\x13as_ascii.exe）を書き、実行時にブックの保存先と結ぶ。
'             パッケージ全体を別の場所へ移しても、同じ設定のまま動く。
' 主な入口  : PackageRoot / ResolvePath / SettingPath / RequireFolder / RequireFile / RelativeToPackage
'             RequireAsciiPath（X-13の作業フォルダの検査）
' 依存先    : modMain（Control）、modX13Options（初期値）
' 規則      : 1) 相対パスはブックの保存先からの位置。".."（上の階層へ出る指定）は受け付けない。
'             2) 「C:\」のようなドライブ文字から始まる絶対パスは、利用者が明示した外部の場所として使う。
'             3) ネットワーク共有（\\server\...）とWeb上の場所（https://...）は使わない。
'             4) CurDir（Excelの現在のフォルダ）には依存しない。

' 目的  : パッケージの最上位フォルダ（このブックの保存先）を返す。
' 出力  : 例 "D:\work\IIP_X13_Package"
' エラー: ブックが未保存、またはOneDrive等のWeb上の場所で開いている場合は停止する。
'         相対パスの基準が決まらないまま計算・保存しないため。
Public Function PackageRoot() As String
    Dim root As String
    root = ThisWorkbook.Path
    If Len(root) = 0 Then Err.Raise vbObjectError + 3000, "PackageRoot", _
        "このブックはまだ保存されていません。IIP_X13_Packageフォルダの直下に.xlsmとして保存してから操作してください。"
    If LCase$(Left$(root, 4)) = "http" Then Err.Raise vbObjectError + 3001, "PackageRoot", _
        "このブックはWeb上の場所（" & root & "）で開かれています。パッケージをローカルのフォルダへ置き、そこから開いてください。"
    If Left$(root, 2) = "\\" Then Err.Raise vbObjectError + 3002, "PackageRoot", _
        "ネットワーク共有上のブックからは実行できません。ローカルのドライブへパッケージを置いてください：" & root
    PackageRoot = root
End Function

' 目的  : ドライブ文字から始まる絶対パスか（例 C:\ または C:/）。
Public Function IsAbsolutePath(ByVal pathText As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z]:[\\/]"
    IsAbsolutePath = re.Test(pathText)
End Function

' 目的  : 設定に書かれたパスを絶対パスにする。
' 入力  : configured = 相対パス（パッケージ基準）または絶対パス
' 出力  : 絶対パス（区切りは \）。configuredが空なら空文字。
' エラー: ".."を含む相対パス、\\で始まる共有、\だけで始まるパスは停止する。
Public Function ResolvePath(ByVal configured As String) As String
    Dim text As String, part As Variant
    text = Trim$(Replace(configured, "/", "\"))
    If Len(text) = 0 Then Exit Function
    If IsAbsolutePath(text) Then
        ResolvePath = text
        Exit Function
    End If
    If Left$(text, 1) = "\" Then Err.Raise vbObjectError + 3003, "ResolvePath", _
        "パスは「01_engine\x13as_ascii.exe」のような相対パスか、「C:\...」の絶対パスで指定してください：" & configured
    For Each part In Split(text, "\")
        If CStr(part) = ".." Then Err.Raise vbObjectError + 3004, "ResolvePath", _
            "相対パスに「..」は使えません。パッケージの外を使う場合は絶対パスで指定してください：" & configured
    Next part
    ResolvePath = modMain.CombinePath(PackageRoot(), text)
End Function

' 目的  : 絶対パスを、パッケージ内ならパッケージ基準の相対パスへ戻す（シートへ記録する用途）。
' 出力  : パッケージ内なら "04_results\candidates\..."、外なら絶対パスのまま。
Public Function RelativeToPackage(ByVal absolutePath As String) As String
    Dim root As String
    root = PackageRoot() & "\"
    If StrComp(Left$(absolutePath, Len(root)), root, vbTextCompare) = 0 Then
        RelativeToPackage = Mid$(absolutePath, Len(root) + 1)
    Else
        RelativeToPackage = absolutePath
    End If
End Function

' 目的  : Controlのパス設定を読み、絶対パスで返す。空欄は設定定義の初期値を使う。
Public Function SettingPath(ByVal key As String) As String
    SettingPath = ResolvePath(modMain.GetControlValue(key, CStr(modX13Options.DefinitionOf(key)(modX13Options.DEF_DEFAULT))))
End Function

' 目的  : フォルダ設定を読み、存在を確かめて絶対パスを返す。
' 入力  : key = Controlの設定キー、createInsidePackage = パッケージ内で無い場合に作るか
' 副作用: createInsidePackage=Trueで、パッケージ内のフォルダが無い場合は作る（上の階層も含めて）。
' エラー: 設定が空、外部のフォルダが無い場合は停止する。外部の場所は勝手に作らない。
Public Function RequireFolder(ByVal key As String, ByVal createInsidePackage As Boolean) As String
    Dim folder As String
    folder = SettingPath(key)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 3005, "RequireFolder", "Controlの" & key & "が空欄です。"
    If Not modMain.FolderExists(folder) Then
        If createInsidePackage And RelativeToPackage(folder) <> folder Then
            EnsureFolderTree folder
        Else
            Err.Raise vbObjectError + 3006, "RequireFolder", "フォルダがありません（Controlの" & key & "）：" & folder
        End If
    End If
    RequireFolder = folder
End Function

' 目的  : ファイル設定を読み、存在を確かめて絶対パスを返す。
' エラー: 設定が空、またはファイルが無い場合は停止する。
Public Function RequireFile(ByVal key As String) As String
    Dim path As String
    path = SettingPath(key)
    If Len(path) = 0 Then Err.Raise vbObjectError + 3007, "RequireFile", "Controlの" & key & "が空欄です。"
    If Not modMain.FileExists(path) Then Err.Raise vbObjectError + 3008, "RequireFile", _
        "ファイルがありません（Controlの" & key & "）：" & path
    RequireFile = path
End Function

' 目的  : フォルダを上の階層から順に作る。既にあれば何もしない。
' エラー: 作れない場合は元のエラーを送出する。
Public Sub EnsureFolderTree(ByVal folder As String)
    Dim fso As Object, parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folder) Then Exit Sub
    parent = fso.GetParentFolderName(folder)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then EnsureFolderTree parent
    fso.CreateFolder folder
End Sub

' 目的  : X-13の作業フォルダに使えるパスか確かめる。
' 規則  : 半角英数字・記号（スペース可）だけを許す。日本語等の文字と引用符・制御文字は不可。
'         X-13（Fortran製のASCII版）が非ASCIIのフォルダで正しく動くことは確認していないため、
'         実行前に止めて、英数字だけの場所（例 C:\IIP_X13_work）を指定してもらう。
' エラー: 最初に見つかった使えない文字と位置を示して停止する。
Public Sub RequireAsciiPath(ByVal pathText As String, ByVal purpose As String)
    Dim i As Long, code As Long
    For i = 1 To Len(pathText)
        code = AscW(Mid$(pathText, i, 1))
        If code < 32 Or code > 126 Or code = 34 Then
            Err.Raise vbObjectError + 3009, "RequireAsciiPath", purpose & "に半角英数字以外の文字「" & Mid$(pathText, i, 1) & _
                "」（" & i & "文字目）が含まれています：" & pathText & vbLf & _
                "ControlのX13_WORK_DIRに、英数字だけのフォルダ（例 C:\IIP_X13_work）を絶対パスで指定してください。" & _
                "計算後の保存先（04_results\candidates等）は日本語を含む場所でも使えます。"
        End If
    Next i
End Sub
