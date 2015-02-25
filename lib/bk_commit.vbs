' Commit ------------------------------------------------------------------------------------------
Sub Commit
    Dim toCopyDir
    Dim commitLog
    Dim intCommitFiles


    If objParm.Count < 2 Then
        EchoUsage "BK COMMIT ""commit-message"""
        WScript.Quit
    ELse
        strCmdTarget = objParm(1)
    End If

    If objFileSys.FolderExists(REPOSITORY_DIR) Then
        ' コミットフォルダを作成（yyyymmdd_hhmmss_CommitMessage）
        commitLog = FormatDateTime(Now, 2) & "_" & Right("0" &  FormatDateTime(Now, 3), 8) & "_" & strCmdTarget
        commitLog = Replace(commitLog, "/", "")  ' フォルダ名に使用できない文字を削除
        commitLog = Replace(commitLog, ":", "")
        commitLog = Replace(commitLog, ",", "_")
        toCopyDir = REPOSITORY_DIR & "\" & commitLog
        If CreateFolder(toCopyDir) Then

            ' コミットログを表示
            WScript.echo vbCRLF & CreateCommitLog(commitLog) & vbCRLF

            ' ファイルをコミット（カレントディレクトのファイルをコピー）
            ' ワーキングディレクトリ内のサブフォルダを探す
            intCommitFiles = TrackedFolders(WORK_DIR, toCopyDir, ECHO)

            ' ワーキングディレクトリ内のファイルを探す
            intCommitFiles = intCommitFiles + TrackedFiles(WORK_DIR, toCopyDir, ECHO)

            WScript.echo vbCRLF & intCommitFiles & " folders/files commited"
        End If

    End If

End Sub
