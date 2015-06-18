'''MsgBox Complement.GetCurrentWord()
' ------------------------------------------------------------------------------
' Constant
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

' python.exe�ɗ^����t���O
Dim strPythonExeFlags
strPythonExeFlags = "-O"

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
strEditingFilePath = """" & strEditingFilePath & """"

' ���O�o�͗L��
Dim blnLogEnabled
blnLogEnabled = Plugin.GetOption("Option", "LogEnabled")

' ���O�t�@�C���p�X
Dim strLogFilePath
strLogFilePath = Plugin.GetOption("Option", "LogFilePath")

' �C���^�v���^�G���[�o�͐�
Dim strStdErrPath
strStdErrPath = Plugin.GetOption("Option", "PyStdErrPath")

'''MsgBox strEditingTempPath
'''MsgBox strCmplTempPath
'''MsgBox nLine
'''MsgBox nColumn

Dim oFileSystem
Set oFileSystem = CreateObject("Scripting.FileSystemObject")
Dim oFileStream01
Dim oFileStream99 ' for logging

' ------------------------------------------------------------------------------
' Function/Procedure
' ------------------------------------------------------------------------------

' ���d�N���h�~
Const strLapFilePath = "C:\temp\jedi-sakura-save.txt"
Function JS_GetLastExecutionTimestamp() 
    Dim strTimestamp
    
    If Err.Number = 0 Then
        If oFileSystem.FileExists( strLapFilePath ) Then
            Set oFile = oFileSystem.OpenTextFile( strLapFilePath )
            
            
            If Err.Number = 0 Then
                Do While oFile.AtEndOfStream <> True
                    strTimestamp = oFile.ReadLine()

                    JS_GetLastExecutionTimestamp = strTimestamp
                Loop
                oFile.Close
            Else
                Echo "�t�@�C���I�[�v���G���[: " & Err.Description
            End If

        Else
            JS_GetLastExecutionTimestamp = "1970/01/01 00:00:00"
        End If
    Else
        Echo "�G���[: " & Err.Description
    End If

    Set oFile = Nothing
End Function

Sub JS_Touch()
    Set oFileStream01 = oFileSystem.CreateTextFile(strLapFilePath, 8, True) ' open for appending
    oFileStream01.Write(CStr(Now))
    oFileStream01.Close
End Sub

' ���s�u���b�N�i���d���s�h�~�j
Function JS_IsBlocked()
    Dim strCurrentTimestamp
    Dim strLastTimestamp
    
    strCurrentTimestamp = Now
    strLastTimestamp = JS_GetLastExecutionTimestamp()
    'MsgBox strCurrentTimestamp & vbNewLine & strLastTimestamp
    If StrComp(strCurrentTimestamp, strLastTimestamp, vbBinaryCompare) = 0 Then
        JS_IsBlocked = True
    Else
        JS_IsBlocked = False
    End If

End Function

' ���M���O�i�f�o�b�O�p�j
Sub JS_Logging(strMessage)
    If blnLogEnabled = True Then
        Set oFileStream99 = oFileSystem.OpenTextFile(strLogFilePath, ForAppending, True)
        
        oFileStream99.WriteLine(Now() & " " & strMessage)
        
        oFileStream99.Close
        Set oFileStream99 = Nothing
    End If
End Sub

' jedi-sakura�p���b�Z�[�W�{�b�N�X
Sub JS_MsgBox(strMessage)
    MsgBox strMessage, vbOKOnly + vbExclamation, "jedi-sakura"
End Sub

JS_Logging("start")

' ------------------------------------------------------------------------------
' Application
'   - Initialize
' ------------------------------------------------------------------------------
' �K�v�Ȓl�����邩
Function JS_IsRequirementsSatisfied()
    Dim blnSatisfied

    blnSatisfied = True
    ' Python�̃x�[�X�f�B���N�g��
    If strPythonEnv = "" Then
        JS_MsgBox "�ݒ�'Python�x�[�X�f�B���N�g��'������܂���B" & vbNewLine & "�v���O�C���ݒ���m�F���Ă��������B"
        blnSatisfied = False
    ' �ҏW���t�@�C���̈ꎞ�ۑ���p�X
    ElseIf strEditingTempPath = "" Then
        JS_MsgBox "�ݒ�'�ҏW���t�@�C���̈ꎞ�ۑ���p�X'������܂���B" & vbNewLine & "�v���O�C���ݒ���m�F���Ă��������B"
        blnSatisfied = False
    ' �⊮�ꎞ�t�@�C���ւ̃p�X
    ElseIf strCmplTempPath = "" Then
        JS_MsgBox "�ݒ�'�⊮�ꎞ�t�@�C���ւ̃p�X'������܂���B" & vbNewLine & "�v���O�C���ݒ���m�F���Ă��������B"
        blnSatisfied = False
    ' ���O�L���Ń��O�t�@�C���o��
    ElseIf blnLogEnabled And strLogFilePath = "" Then
        JS_MsgBox "�ݒ�'���O�t�@�C���p�X'������܂���B" & vbNewLine & "�v���O�C���ݒ���m�F���Ă��������B"
        blnSatisfied = False
    End If
    JS_IsRequirementsSatisfied = blnSatisfied
End Function
If JS_IsRequirementsSatisfied = True Then
' ------------------------------------------------------------------------------
'   - Main
' ------------------------------------------------------------------------------
If Not JS_IsBlocked Then

    Dim strBuffer

    ' �ҏW���o�b�t�@��ۑ�����
    '   �`��X�g�b�v
    Editor.SetDrawSwitch(0)
    '   �ҏW���o�b�t�@���擾
        Editor.MoveHistSet()
        Editor.SelectAll()
        strBuffer = Editor.GetSelectedString(0)
        Editor.MoveHistPrev()
    Editor.SetDrawSwitch(1)
    Editor.ReDraw(0)

    '   �ꎞ�ۑ���ցA�ҏW���o�b�t�@��ۑ�
    Set oFileStream01 = oFileSystem.CreateTextFile(strEditingTempPath, True)
    oFileStream01.Write(strBuffer)
    oFileStream01.Close

    ' �R�}���h���C�����s��������擾
    Dim WshShell, oExec
    Set WshShell = CreateObject("WScript.Shell")

    strCommand = "%comspec% /c " & strPythonExe _
               & " " & strPythonExeFlags _
               & " " & strJediBridgeScrpit _
               & " " & nLine _
               & " " & nColumn _
               & " " & strEditingFilePath _
               & " " & strEditingTempPath _
               & " 1> " & strCmplTempPath _
               & " 2> " & strStdErrPath
'�ݒ�t�@�C����""�͂��Ă�����������悢
'    strCommand = "%comspec% /c " & strPythonExe _
'               & " """ & strPythonExeFlags & """" _
'               & " """ & strJediBridgeScrpit & """" _
'               & " " & nLine _
'               & " " & nColumn _
'               & " """ & strEditingFilePath & """" _
'               & " """ & strEditingTempPath & """" _
'               & " 1> """ & strCmplTempPath & """" _
'               & " 2> """ & strStdErrPath & """"
    '''MsgBox strCommand

    Dim nReturnCode
    nReturnCode = WshShell.Run(strCommand, 0, True)

End If

' Python�C���^�v���^�ŃG���[���o���ꍇ�G���[���b�Z�[�W��\��
Dim strStdErrContent
If oFileSystem.FileExists( strStdErrPath ) Then
    Set oFileStream01 = oFileSystem.OpenTextFile( strStdErrPath )

    strStdErrContent = ""
    Do Until oFileStream01.AtEndofStream
        strStdErrContent = strStdErrContent & oFileStream01.ReadLine & vbNewLine
    Loop

    If strStdErrContent <> "" Then
        JS_MsgBox strStdErrContent
    End If

    oFileStream01.Close
End If

' �⊮�t�@�C����ǂݎ��
Set oFileStream01 = oFileSystem.OpenTextFile(strCmplTempPath)
Do While oFileStream01.AtEndOfStream <> True
    word = oFileStream01.ReadLine()
    ' �umath.�v�Ƃ��p
    If Complement.GetCurrentWord() = "." Then
        Complement.AddList "." & word
    ' �u^import �v�p .. �X�y�[�X���ƃG�f�B�^���ŕ⊮������Ȃ��̂ŕ⊮���悤������
    'ElseIf Complement.GetCurrentWord() = "" And Editor.S_GetLineStr(nLine) = "import " Then
    '    Complement.AddList word
    Else
        Complement.AddList word
    End If
Loop
oFileStream01.Close

' �������X�V
JS_Touch

End If

' ------------------------------------------------------------------------------
'   - Finalize
' ------------------------------------------------------------------------------
Set oFileSystem = Nothing
