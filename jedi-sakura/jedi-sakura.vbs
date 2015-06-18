'
' jediを使った入力補完
'
' ref: http://d.hatena.ne.jp/language_and_engineering/20081021/1224511689
' ref: http://www.geocities.jp/maru3128/SakuraMacro/
' ref: http://d.hatena.ne.jp/nanntekotta/20100623/1277300689

' TODO: ドットだけの場合も ( "requests." みたいな) 補完できるようにする。具体的には1つ前のトークンを含めて補完ファイルを再作成する。

Dim strPythonEnvironment    ' Python環境へのパス
Dim strPythonInterpreter    ' Pythonインタプリタへのパス
Dim strCompletionEnginePath ' 補完候補をファイルに出力するPythonスクリプトへのパス
Dim strInputPath            ' 補完対象の入力ファイル（編集を一時的に出力したファイルへのパス）
Dim strInputEditingPath     ' 補完対象の入力ファイル（本来のパス）
Dim strOutputPath           ' 補完候補ファイルの出力パス
Dim strCursorLineAt         ' 補完する位置（行）
Dim strCursorColumnAt       ' 補完する位置（列）
Dim strPythonCode           ' コード文字列（ファイル中のテキスト全て）
Dim strCommand              ' コマンド


strPythonEnvironment = "C:\Python27"
strPythonInterpreter = strPythonEnvironment & "\python.exe"
strCompletionEnginePath = "C:\jedi_sakura\jedi_sakura.py"
strInputPath = Editor.ExpandParameter("$F")
strInputEditingPath = "C:\temp\temp.txt"
strOutputPath = strPythonEnvironment & "\completion_list.txt"
strCursorLineAt = Editor.ExpandParameter("$y")
strCursorColumnAt = Editor.ExpandParameter("$x") - 1

' 
' ファイルを一時ファイルに保存
'   全テキスト取得
Dim strSelectedString

Editor.MoveHistSet()
Editor.SelectAll()
strSelectedString = Editor.GetSelectedString(0)
Editor.MoveHistPrev()

'   ファイル保存
Dim oFileSystem
Dim oFileStream
Set oFileSystem = CreateObject("Scripting.FileSystemObject")

' FileSystemObject.CreateTextFile(FileName, Overwrite, Unicode)
Set oFileStream = oFIleSystem.CreateTextFile(strInputEditingPath, True, False)
oFileStream.Write strSelectedString
oFileStream.Close

Set oFileStream = Nothing
Set oFileSystem = Nothing

'v2.1.1.1のバグ？PutFileでクラッシュする。処理が実装されていない？要確認
'MsgBox(strInputEditingPath)
'Editor.PutFile strInputEditingPath, -1, 0

' 編集中の（未保存の）状態を書きだす
strCommand = strPythonInterpreter _
           & " " & strCompletionEnginePath _
           & " " & strInputPath _
           & " " & strInputEditingPath _
           & " " & strOutputPath _
           & " " & strCursorLineAt _
           & " " & strCursorColumnAt

'MsgBox(strCommand)

Const SW_SHOWMINNOACTIVE = 7

' 補完候補をファイルに書き出し
Set WshShell = CreateObject("WScript.Shell") 
Set oExec = WshShell.Exec( strCommand ) 
'https://msdn.microsoft.com/ja-jp/library/cc364421.aspx
'Set oExec = WshShell.Run( strCommand, SW_SHOWMINNOACTIVE, True ) 
'   終了まで待機
Do While oExec.Status = 0
Loop 

' 補完画面表示
Editor.Complete()
