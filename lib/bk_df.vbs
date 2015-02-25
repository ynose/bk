' Df ----------------------------------------------------------------------------------------------
Sub ExecDf
    Dim objHEADRepoFolder
    Dim pathHEADRepo, pathWork
    Dim WshShell
    Const DFEXE = "C:\Program Files\DF\DF.exe"


    If objParm.Count < 2 Then
        EchoUsage "BK DF filename"
        WScript.Quit
    ELse
        ' 対象ファイルを追加する
        strCmdTarget = objParm(1)
    End If

    If objFileSys.FileExists(DFEXE) = False Then
        EchoError "Not exist " & DFEXE
        WScript.Quit
    End If


    Set objHEADRepoFolder = objFileSys.GetFolder(REPOSITORY_DIR & "\" & GetHEADRepository())
    pathHEADRepo = objHEADRepoFolder.Path & "\" & strCmdTarget
    pathWork = objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & strCmdTarget

    ' 前回のコミットログを表示
    WScript.echo vbCRLF & CreateCommitLog(objHEADRepoFolder.Name)
    WScript.echo pathWork

    Call ShellRun("""" & DFEXE & """" & " """ & pathWork & """ """ & pathHEADRepo & """", 5)
    
End Sub
