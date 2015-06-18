'''MsgBox Complement.GetCurrentWord()
' ------------------------------------------------------------------------------

' jediを呼ぶpythonスクリプトへのパス
'     e.g.) C:\Users\%USERNAME%\AppData\Roaming\sakura\plugins\jedi-sakura\jedi-sakura.py
Dim strJediBridgeScrpit
strJediBridgeScrpit = Plugin.GetPluginDir() & "\jedi-sakura.py"

' Pythonのベースディレクトリ
'     e.g.) C:\Python27, C:\Python34 etc
Dim strPythonEnv
strPythonEnv = Plugin.GetOption("Option", "PyEnv")
' python.exeへのパス
'     e.g.) C:\Python27\python.exe
Dim strPythonExe
strPythonExe = strPythonEnv & "\python.exe"

' 編集中ファイルの一時保存先パス
'     e.g.) C:\jedi-sakura-editing.txt
Dim strEditingTempPath
strEditingTempPath = Plugin.GetOption("Option", "EditingTempPath")

' 補完一時ファイルへのパス
'     e.g.) C:\jedi-sakura-cmpl.txt
Dim strCmplTempPath
strCmplTempPath = Plugin.GetOption("Option", "CmplTempPath")

' カーソル位置（行） start by 1
Dim nLine
nLine = Editor.ExpandParameter("$y")

' カーソル位置（列） start by 1
Dim nColumn
nColumn = Editor.ExpandParameter("$x") - 1

' 入力ファイルパス
Dim strEditingFilePath
strEditingFilePath = Editor.ExpandParameter("$F")

'''MsgBox strEditingTempPath
'''MsgBox strCmplTempPath
'''MsgBox nLine
'''MsgBox nColumn
' ------------------------------------------------------------------------------

Dim strBuffer

' 描画ストップ
Editor.SetDrawSwitch(0)
'   編集中バッファを取得
    Editor.MoveHistSet()
    Editor.SelectAll()
    strBuffer = Editor.GetSelectedString(0)
    Editor.MoveHistPrev()
Editor.SetDrawSwitch(1)
Editor.ReDraw(0)

'   一時保存先へ、編集中バッファを保存
Dim fso, MyFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile(strEditingTempPath, True)
MyFile.Write(strBuffer)
MyFile.Close


Dim WshShell, oExec
Set WshShell = CreateObject("WScript.Shell")

strCommand = "%comspec% /c " & strPythonExe _
           & " " & strJediBridgeScrpit _
           & " " & nLine _
           & " " & nColumn _
           & " " & strEditingFilePath _
           & " " & strEditingTempPath _
           & " 1> " & strCmplTempPath _
           & " 2>&1"
'''MsgBox strCommand
Dim nReturnCode
nReturnCode = WshShell.Run(strCommand, 0, True)



Dim objFSO      ' FileSystemObject
Dim objFile     ' ファイル読み込み用

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    Set objFile = objFSO.OpenTextFile(strCmplTempPath)
    If Err.Number = 0 Then
        Do While objFile.AtEndOfStream <> True


            word = objFile.ReadLine()
            If Complement.GetCurrentWord() = "." Then
                Complement.AddList "." & word
            Else
                Complement.AddList word
            End If


        Loop
        objFile.Close
    Else
        Echo "ファイルオープンエラー: " & Err.Description
    End If
Else
    Echo "エラー: " & Err.Description
End If

Set objFile = Nothing
Set objFSO = Nothing

