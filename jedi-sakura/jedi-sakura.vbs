'
' jedi���g�������͕⊮
'
' ref: http://d.hatena.ne.jp/language_and_engineering/20081021/1224511689
' ref: http://www.geocities.jp/maru3128/SakuraMacro/
' ref: http://d.hatena.ne.jp/nanntekotta/20100623/1277300689

' TODO: �h�b�g�����̏ꍇ�� ( "requests." �݂�����) �⊮�ł���悤�ɂ���B��̓I�ɂ�1�O�̃g�[�N�����܂߂ĕ⊮�t�@�C�����č쐬����B

Dim strPythonEnvironment    ' Python���ւ̃p�X
Dim strPythonInterpreter    ' Python�C���^�v���^�ւ̃p�X
Dim strCompletionEnginePath ' �⊮�����t�@�C���ɏo�͂���Python�X�N���v�g�ւ̃p�X
Dim strInputPath            ' �⊮�Ώۂ̓��̓t�@�C���i�ҏW���ꎞ�I�ɏo�͂����t�@�C���ւ̃p�X�j
Dim strInputEditingPath     ' �⊮�Ώۂ̓��̓t�@�C���i�{���̃p�X�j
Dim strOutputPath           ' �⊮���t�@�C���̏o�̓p�X
Dim strCursorLineAt         ' �⊮����ʒu�i�s�j
Dim strCursorColumnAt       ' �⊮����ʒu�i��j
Dim strPythonCode           ' �R�[�h������i�t�@�C�����̃e�L�X�g�S�āj
Dim strCommand              ' �R�}���h


strPythonEnvironment = "C:\Python27"
strPythonInterpreter = strPythonEnvironment & "\python.exe"
strCompletionEnginePath = "C:\jedi_sakura\jedi_sakura.py"
strInputPath = Editor.ExpandParameter("$F")
strInputEditingPath = "C:\temp\temp.txt"
strOutputPath = strPythonEnvironment & "\completion_list.txt"
strCursorLineAt = Editor.ExpandParameter("$y")
strCursorColumnAt = Editor.ExpandParameter("$x") - 1

' 
' �t�@�C�����ꎞ�t�@�C���ɕۑ�
'   �S�e�L�X�g�擾
Dim strSelectedString

Editor.MoveHistSet()
Editor.SelectAll()
strSelectedString = Editor.GetSelectedString(0)
Editor.MoveHistPrev()

'   �t�@�C���ۑ�
Dim oFileSystem
Dim oFileStream
Set oFileSystem = CreateObject("Scripting.FileSystemObject")

' FileSystemObject.CreateTextFile(FileName, Overwrite, Unicode)
Set oFileStream = oFIleSystem.CreateTextFile(strInputEditingPath, True, False)
oFileStream.Write strSelectedString
oFileStream.Close

Set oFileStream = Nothing
Set oFileSystem = Nothing

'v2.1.1.1�̃o�O�HPutFile�ŃN���b�V������B��������������Ă��Ȃ��H�v�m�F
'MsgBox(strInputEditingPath)
'Editor.PutFile strInputEditingPath, -1, 0

' �ҏW���́i���ۑ��́j��Ԃ���������
strCommand = strPythonInterpreter _
           & " " & strCompletionEnginePath _
           & " " & strInputPath _
           & " " & strInputEditingPath _
           & " " & strOutputPath _
           & " " & strCursorLineAt _
           & " " & strCursorColumnAt

'MsgBox(strCommand)

Const SW_SHOWMINNOACTIVE = 7

' �⊮�����t�@�C���ɏ����o��
Set WshShell = CreateObject("WScript.Shell") 
Set oExec = WshShell.Exec( strCommand ) 
'https://msdn.microsoft.com/ja-jp/library/cc364421.aspx
'Set oExec = WshShell.Run( strCommand, SW_SHOWMINNOACTIVE, True ) 
'   �I���܂őҋ@
Do While oExec.Status = 0
Loop 

' �⊮��ʕ\��
Editor.Complete()
