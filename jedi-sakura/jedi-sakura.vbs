'''MsgBox Complement.GetCurrentWord()
' ------------------------------------------------------------------------------
' Constant
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

' python.exeに与えるフラグ
Dim strPythonExeFlags
strPythonExeFlags = "-O"

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
strEditingFilePath = """" & strEditingFilePath & """"

' ログ出力有無
Dim blnLogEnabled
blnLogEnabled = Plugin.GetOption("Option", "LogEnabled")

' ログファイルパス
Dim strLogFilePath
strLogFilePath = Plugin.GetOption("Option", "LogFilePath")

' インタプリタエラー出力先
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

' 多重起動防止
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
                Echo "ファイルオープンエラー: " & Err.Description
            End If

        Else
            JS_GetLastExecutionTimestamp = "1970/01/01 00:00:00"
        End If
    Else
        Echo "エラー: " & Err.Description
    End If

    Set oFile = Nothing
End Function

Sub JS_Touch()
    Set oFileStream01 = oFileSystem.CreateTextFile(strLapFilePath, 8, True) ' open for appending
    oFileStream01.Write(CStr(Now))
    oFileStream01.Close
End Sub

' 実行ブロック（多重実行防止）
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

' ロギング（デバッグ用）
Sub JS_Logging(strMessage)
    If blnLogEnabled = True Then
        Set oFileStream99 = oFileSystem.OpenTextFile(strLogFilePath, ForAppending, True)
        
        oFileStream99.WriteLine(Now() & " " & strMessage)
        
        oFileStream99.Close
        Set oFileStream99 = Nothing
    End If
End Sub

' jedi-sakura用メッセージボックス
Sub JS_MsgBox(strMessage)
    MsgBox strMessage, vbOKOnly + vbExclamation, "jedi-sakura"
End Sub

JS_Logging("start")

' ------------------------------------------------------------------------------
' Application
'   - Initialize
' ------------------------------------------------------------------------------
' 必要な値があるか
Function JS_IsRequirementsSatisfied()
    Dim blnSatisfied

    blnSatisfied = True
    ' Pythonのベースディレクトリ
    If strPythonEnv = "" Then
        JS_MsgBox "設定'Pythonベースディレクトリ'がありません。" & vbNewLine & "プラグイン設定を確認してください。"
        blnSatisfied = False
    ' 編集中ファイルの一時保存先パス
    ElseIf strEditingTempPath = "" Then
        JS_MsgBox "設定'編集中ファイルの一時保存先パス'がありません。" & vbNewLine & "プラグイン設定を確認してください。"
        blnSatisfied = False
    ' 補完一時ファイルへのパス
    ElseIf strCmplTempPath = "" Then
        JS_MsgBox "設定'補完一時ファイルへのパス'がありません。" & vbNewLine & "プラグイン設定を確認してください。"
        blnSatisfied = False
    ' ログ有効でログファイル出力
    ElseIf blnLogEnabled And strLogFilePath = "" Then
        JS_MsgBox "設定'ログファイルパス'がありません。" & vbNewLine & "プラグイン設定を確認してください。"
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

    ' 編集中バッファを保存する
    '   描画ストップ
    Editor.SetDrawSwitch(0)
    '   編集中バッファを取得
        Editor.MoveHistSet()
        Editor.SelectAll()
        strBuffer = Editor.GetSelectedString(0)
        Editor.MoveHistPrev()
    Editor.SetDrawSwitch(1)
    Editor.ReDraw(0)

    '   一時保存先へ、編集中バッファを保存
    Set oFileStream01 = oFileSystem.CreateTextFile(strEditingTempPath, True)
    oFileStream01.Write(strBuffer)
    oFileStream01.Close

    ' コマンドライン実行文字列を取得
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
'設定ファイルで""囲ってもらった方がよい
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

' Pythonインタプリタでエラーが出た場合エラーメッセージを表示
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

' 補完ファイルを読み取る
Set oFileStream01 = oFileSystem.OpenTextFile(strCmplTempPath)
Do While oFileStream01.AtEndOfStream <> True
    word = oFileStream01.ReadLine()
    ' 「math.」とか用
    If Complement.GetCurrentWord() = "." Then
        Complement.AddList "." & word
    ' 「^import 」用 .. スペースだとエディタ側で補完が走らないので補完しようが無い
    'ElseIf Complement.GetCurrentWord() = "" And Editor.S_GetLineStr(nLine) = "import " Then
    '    Complement.AddList word
    Else
        Complement.AddList word
    End If
Loop
oFileStream01.Close

' 時刻を更新
JS_Touch

End If

' ------------------------------------------------------------------------------
'   - Finalize
' ------------------------------------------------------------------------------
Set oFileSystem = Nothing
