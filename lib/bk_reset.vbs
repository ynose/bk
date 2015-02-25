' Reset -------------------------------------------------------------------------------------------
Sub CommitReset
    Dim objHEADRepoFolder, objHEADRepoFile
    Dim objWorkFolder, objWorkFile
    

    If objFileSys.FolderExists(REPOSITORY_DIR) = False Then
        EchoError "Not a bk repository : " & REPOSITORY
        Exit Sub
    End If
    If GetHEADRepository() = "" Then
        EchoError "Not a commit"
        Exit Sub
    End If
    
    Set objHEADRepoFolder = objFileSys.GetFolder(REPOSITORY_DIR & "\" & GetHEADRepository())

    ' 前回のコミットログを表示
    WScript.echo vbCRLF & CreateCommitLog(objHEADRepoFolder.Name) & vbCRLF
    

    ' ワーキングディレクトリ内のファイルをUNDO用に退避する（カレントディレクトのファイルをコピー）
    Call DeleteFolder(RESETUNDO_DIR)
    WScript.sleep(500) ' DeleteFolderの後に少しスリープを入れないCreateFolderでエラーになる
    If CreateFolder(RESETUNDO_DIR) Then

        ' ワーキングディレクトリ内のサブフォルダを探す
        Call TrackedFolders(WORK_DIR, RESETUNDO_DIR, NOECHO)

        ' ワーキングディレクトリ内のファイルを探す
        Call TrackedFiles(WORK_DIR, RESETUNDO_DIR, NOECHO)

    End If

    
    ' ワーキングディレクトリにある管理対象ファイルを削除する
    Set objWorkFolder = objFileSys.GetFolder(WORK_DIR)
    For Each objWorkFile In objFileSys.GetFolder(objWorkFolder.Path).Files
        If IsTracked(objWorkFile.Name) Then

' 削除は怖いので一時保留
'            objFileSys.DeleteFile objWorkFile.path

        End If
    Next
    
 
    ' リポジトリの最終コミット(HEAD)のファイルを
    ' ワーキングディレクトリに復元（上書きコピー）する
    For Each objHEADRepoFile In objFileSys.GetFolder(objHEADRepoFolder.Path).Files
    
        objHEADRepoFile.Copy WORK_DIR
    
    Next
    
    
    WScript.echo "reset working folder"

End Sub

' 管理対象かを判定する
Function IsTracked(filename)
    Dim pathTrackFile
    Dim tracked
    Dim objFile
    Dim objRE
    Dim strPattern


    pathTrackFile = REPOSITORY_DIR & "\" & TRACKFILE

    ' 正規表現で対象ファイルであるかを判定する
    Set objRE = CreateObject("VBScript.RegExp") ' 正規表現
    objRE.IgnoreCase = True ' 大文字小文字を区別しない

    ' リポジトリ用フォルダは対象外とする
    strPattern = objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & REPOSITORY
    strPattern = Replace(strPattern, "\", "\\")   ' エスケープ
    strPattern = Replace(strPattern, ".", "\.")   ' エスケープ
    objRE.Pattern = "^" & strPattern & ".*$"
'    Call DebugPrint(objRE.Pattern)

'    Call DebugPrint(filename)

    If objRE.Test(filename) Then
        tracked = false ' リポジトリ用フォルダは対象外とする
        
    ELseIf objFileSys.FileExists(pathTrackFile) Then
    
        tracked = false
        Set objFile = objFileSys.OpenTextFile(pathTrackFile, ForReading, true)
        Do Until objFile.AtEndOfStream
            strPattern = objFile.ReadLine
            strPattern = objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & strPattern
            strPattern = Replace(strPattern, "\", "\\")   ' エスケープ
            strPattern = Replace(strPattern, ".", "\.")   ' エスケープ
            strPattern = Replace(strPattern, "*", ".+")
            strPattern = Replace(strPattern, "?", ".")
            objRE.Pattern = "^" & strPattern & "$"
            
            'Call DebugPrint(objRE.Pattern)
            
            If objRE.Test(filename & "\") Then
                Call DebugPrint(objRE.Pattern & ">" & filename)
                tracked = true
                Exit Do
            End If
        Loop

        objFile.Close
    Else
        ' Trackファイルが無い場合は全ファイルを対象とする
        tracked = true
    End If

    IsTracked = tracked

End Function
