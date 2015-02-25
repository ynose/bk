' Init   ------------------------------------------------------------------------------------------
Sub Initialize

    ' リポジトリ用フォルダを作成する
    If objFileSys.FolderExists(REPOSITORY_DIR) = False Then
        If CreateFolder(REPOSITORY_DIR) Then
            EchoSuccess "Initialized empty BK repository in " & objFileSys.GetAbsolutePathName(REPOSITORY_DIR)
        End If
    Else
        ' すでに初期化されている
        EchoSuccess "Existing BK repository in " & objFileSys.GetAbsolutePathName(REPOSITORY_DIR)
    End IF
    
End Sub
