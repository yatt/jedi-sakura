'''MsgBox Complement.GetCurrentWord()
' ------------------------------------------------------------------------------

' jedi���Ă�python�X�N���v�g�ւ̃p�X
'     e.g.) C:\Users\%USERNAME%\AppData\Roaming\sakura\plugins\jedi-sakura\jedi-sakura.py
Dim strJediBridgeScrpit
strJediBridgeScrpit = Plugin.GetPluginDir() & "\jedi-sakura.py"

' Python�̃x�[�X�f�B���N�g��
'     e.g.) C:\Python27, C:\Python34 etc
Dim strPythonEnv
strPythonEnv = Plugin.GetOption("Option", "PyEnv")
' python.exe�ւ̃p�X
'     e.g.) C:\Python27\python.exe
Dim strPythonExe
strPythonExe = strPythonEnv & "\python.exe"

' �ҏW���t�@�C���̈ꎞ�ۑ���p�X
'     e.g.) C:\jedi-sakura-editing.txt
Dim strEditingTempPath
strEditingTempPath = Plugin.GetOption("Option", "EditingTempPath")

' �⊮�ꎞ�t�@�C���ւ̃p�X
'     e.g.) C:\jedi-sakura-cmpl.txt
Dim strCmplTempPath
strCmplTempPath = Plugin.GetOption("Option", "CmplTempPath")

' �J�[�\���ʒu�i�s�j start by 1
Dim nLine
nLine = Editor.ExpandParameter("$y")

' �J�[�\���ʒu�i��j start by 1
Dim nColumn
nColumn = Editor.ExpandParameter("$x") - 1

' ���̓t�@�C���p�X
Dim strEditingFilePath
strEditingFilePath = Editor.ExpandParameter("$F")

'''MsgBox strEditingTempPath
'''MsgBox strCmplTempPath
'''MsgBox nLine
'''MsgBox nColumn
' ------------------------------------------------------------------------------

Dim strBuffer

' �`��X�g�b�v
Editor.SetDrawSwitch(0)
'   �ҏW���o�b�t�@���擾
    Editor.MoveHistSet()
    Editor.SelectAll()
    strBuffer = Editor.GetSelectedString(0)
    Editor.MoveHistPrev()
Editor.SetDrawSwitch(1)
Editor.ReDraw(0)

'   �ꎞ�ۑ���ցA�ҏW���o�b�t�@��ۑ�
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
Dim objFile     ' �t�@�C���ǂݍ��ݗp

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
        Echo "�t�@�C���I�[�v���G���[: " & Err.Description
    End If
Else
    Echo "�G���[: " & Err.Description
End If

Set objFile = Nothing
Set objFSO = Nothing

