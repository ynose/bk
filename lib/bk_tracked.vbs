' Tracked --------------------------------------------------------------------------------------
Sub Tracked
    Dim intTrackedFiles


    If objFileSys.FolderExists(REPOSITORY_DIR) Then

        ' ワーキングディレクトリ内のサブフォルダを探す
        intTrackedFiles = TrackedFolders(WORK_DIR, "", ECHO)

        ' ワーキングディレクトリ内のファイルを探す
        intTrackedFiles = intTrackedFiles + TrackedFiles(WORK_DIR, "", ECHO)
        
        WScript.echo vbCRLF & "Tracked " & intTrackedFiles & " File(s)"

    End If

End Sub

Function TrackedFolders(path, copyDir, forceEcho)
    Dim objTempFolder, subFolder
    Dim pathTrackFile
    Dim relativePath
    Dim intTrackedFiles
    Dim objFile
    Dim strPattern
    

    Set objTempFolder = objFileSys.GetFolder(path)

    pathTrackFile = REPOSITORY_DIR & "\" & TRACKFILE

    If objFileSys.FileExists(pathTrackFile) Then
        ' Trackファイルがある場合はTrackファイルに登録されているパスのフォルダーのみコピーする
        Set objFile = objFileSys.OpenTextFile(pathTrackFile, ForReading, true)
        Do Until objFile.AtEndOfStream
            strPattern = objFile.ReadLine

            If objFileSys.FolderExists(strPattern) = True Then
                If forceEcho = ECHO Then WScript.echo "  " & strPattern

                ' コピーする
                If copyDir <> "" Then
                    objFileSys.CopyFolder objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & strPattern, copyDir & "\" & relativePath
                End If

                intTrackedFiles = intTrackedFiles + 1
            End If
        Loop        
        objFile.Close
    Else
        ' Trackファイルがない場合は、すべてのフォルダーをコピーする
        For Each subFolder In objTempFolder.SubFolders
            relativePath = Replace(subFolder, objFileSys.GetAbsolutePathName(WORK_DIR) & "\", "")
            If subFolder.Name <> REPOSITORY Then
                If forceEcho = ECHO Then WScript.echo "  " & relativePath

                ' コピーする
                If copyDir <> "" Then
                    objFileSys.CopyFolder subFolder.Path, copyDir & "\" & relativePath
                End If
                
                intTrackedFiles = intTrackedFiles + 1
            End If
        Next
    End If

    ' Return
    TrackedFolders = intTrackedFiles

End Function

Function TrackedFiles(path, copyDir, forceEcho)
    Dim objFile
    Dim pathTrackFile
    Dim relativePath
    Dim intTrackedFiles
    Dim strPattern


    pathTrackFile = REPOSITORY_DIR & "\" & TRACKFILE

    If objFileSys.FileExists(pathTrackFile) Then
        ' Trackファイルがある場合は、Trackファイルに登録されているパスのファイルのみコピーする
        Set objFile = objFileSys.OpenTextFile(pathTrackFile, ForReading, true)
        Do Until objFile.AtEndOfStream
            strPattern = objFile.ReadLine

            If objFileSys.FileExists(strPattern) = True Then
                If forceEcho = ECHO Then WScript.echo "  " & strPattern
WScript.echo objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & strPattern & " -> " & copyDir & "\" & relativePath & strPattern
                ' コピーする
                If copyDir <> "" Then
                    objFileSys.CopyFile objFileSys.GetAbsolutePathName(WORK_DIR) & "\" & strPattern, copyDir & "\" & relativePath & strPattern
                End If

                intTrackedFiles = intTrackedFiles + 1
            End If
        Loop        
        objFile.Close
    Else
        ' Trackファイルがない場合は、すべてのファイルをコピーする
        For Each objFile In objFileSys.GetFolder(path).Files
            relativePath = Replace(objFile.path, objFileSys.GetAbsolutePathName(WORK_DIR) & "\", "")
            If forceEcho = ECHO Then WScript.echo "  " & relativePath

            ' コピーする
            If copyDir <> "" Then
                objFile.Copy copyDir & "\" & relativePath
            End If

            intTrackedFiles = intTrackedFiles + 1
        Next
    End If
    
    ' Return
    TrackedFiles = intTrackedFiles

End Function
