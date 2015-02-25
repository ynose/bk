' Add    ------------------------------------------------------------------------------------------
Sub addTrackfile
    Dim path
    Dim objFile

    If objParm.Count < 2 Then
        EchoUsage "BK ADD filename"
    ELse
        ' 対象ファイルを追加する
        strCmdTarget = objParm(1)

        WScript.echo "BK ADD"

        Set objFile = objFileSys.OpenTextFile(REPOSITORY_DIR & "\" & TRACKFILE, ForAppending, true)
        objFile.WriteLine strCmdTarget
        objFile.Close
        Set objFile = Nothing
    End If

    ' 対象ファイルの内容を列挙する
    path = REPOSITORY_DIR & "\" & TRACKFILE
    If objFileSys.FileExists(path) Then
    
        Set objFile = objFileSys.OpenTextFile(path, ForReading, true)

        Do Until objFile.AtEndOfStream      ' 入力ファイルの終端まで繰り返し
            WScript.echo objFile.ReadLine
        Loop

        objFile.Close
        Set objFile = Nothing
    Else
        EchoError "Not exist Track file"
    End If

End Sub
