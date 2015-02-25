' Log ---------------------------------------------------------------------------------------------
Sub CommitLog
    Dim objRepositories, strRepository


    Set objRepositories = GetRepositories()
    If objRepositories.Count > 0 Then
        ' フォルダ名(日時順)でソートする
        objRepositories.Sort()     ' 昇順ソート

        ' 日付順にコミットログを出力する
        For Each strRepository In objRepositories
            WScript.Echo CreateCommitLog(strRepository)
        Next
    Else
        EchoError "Not a bk repository : " & REPOSITORY
    End If

    Set objRepositories = Nothing

End Sub
