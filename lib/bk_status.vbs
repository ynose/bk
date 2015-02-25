' Status ------------------------------------------------------------------------------------------
Sub Status
    Const WORKING_DIR  = "working_dir.txt"
    Const WORKING_TIME = "working_time.txt"
    Const WORKING_SORT = "working_sort.txt"
    Const HEAD_DIR     = "HEADRep_dir.txt"
    Const HEAD_TIME    = "HEADRep_time.txt"
    Const HEAD_SORT    = "HEADRep_sort.txt"
    Const FCOUT        = "fc_out.txt"
    Const FCOUT_WORKING = 1
    Const FCOUT_HEAD    = 2
    Const SLEEP = 200


    If objParm.Count < 2 Then
        EchoUsage "BK STATUS filename"
        WScript.Quit
    ELse
        ' 対象ファイルを追加する
        strCmdTarget = objParm(1)
    End If

    If objFileSys.FolderExists(REPOSITORY_DIR) = False Then
        EchoError "Not a bk repository : " & REPOSITORY
        Exit Sub
    End If
    If GetHEADRepository() = "" Then
        EchoError "Not a commit"
        Exit Sub
    End If


    Dim WorkingOutPath_Dir:   WorkingOutPath_Dir   = objFileSys.BuildPath(strScriptPath, WORKING_DIR)
    Dim WorkingOutPath_Time:  WorkingOutPath_Time  = objFileSys.BuildPath(strScriptPath, WORKING_TIME)
    Dim WorkingOutPath_Sort:  WorkingOutPath_Sort  = objFileSys.BuildPath(strScriptPath, WORKING_SORT)
    Dim HEADRepoOutPath_Dir:  HEADRepoOutPath_Dir  = objFileSys.BuildPath(strScriptPath, HEAD_DIR)
    Dim HEADRepoOutPath_Time: HEADRepoOutPath_Time = objFileSys.BuildPath(strScriptPath, HEAD_TIME)
    Dim HEADRepoOutPath_Sort: HEADRepoOutPath_Sort = objFileSys.BuildPath(strScriptPath, HEAD_SORT)
    Dim fcOutPath:            FcOutPath            = objFileSys.BuildPath(strScriptPath, FCOUT)

    Dim strWorkingPath:       strWorkingPath       = objFileSys.BuildPath(WORK_DIR, strCmdTarget & "\")
    Dim strHEADRepoPath:      strHEADRepoPath      = objFileSys.BuildPath(REPOSITORY_DIR & "\" & GetHEADRepository(), strCmdTarget & "\")

    ' 無視するファイルパターンを読み込んで正規表現オブジェクトを作成する
    Dim objRE: Set objRE = Nothing
    Dim ignorePatterns: Set ignorePatterns = New FileReadArrayList
    ignorePatterns.ReadFile(REPOSITORY_DIR & "\" & ".ignore")
    If ignorePatterns.GetInstance.Count > 0 Then
        Set objRE = New RegExpForFilePath
        objRE.FilePaths = ignorePatterns.GetInstance
    End If

    Dim objReadFile, objWriteFile
    Dim strFullPath, strRelativePath
    Dim objFile


    ' ワーキングディレクトリ内の対象ファイルの一覧を作成する
    Call ShellRun("cmd /c del " & WorkingOutPath_Dir & " " & WorkingOutPath_Sort & " " & WorkingOutPath_Time, 0)

    ' dir -> sort
    Call ShellRun("cmd /c dir /b /s " & strWorkingPath & "*.aspx " & strWorkingPath & "*.vb " & strWorkingPath & "*.aspx.resx " & strWorkingPath & "*.ascx " & strWorkingPath & "*.html " & strWorkingPath & "*.js " & strWorkingPath & "*.css " & strWorkingPath & "*.gif " & strWorkingPath & "*.jpg " & strWorkingPath & "*.png " & strWorkingPath & "*.xls > " & WorkingOutPath_Dir, 0)
    WScript.sleep(SLEEP) ' 少しスリープを入れないと実行結果が反映されない
    Call ShellRun("cmd /c sort "  & WorkingOutPath_Dir & " /O " & WorkingOutPath_Sort, 0)
    WScript.sleep(SLEEP) ' 少しスリープを入れないと実行結果が反映されない

    ' sort -> タイムスタンプ埋め込み
    Set objReadFile = objFileSys.OpenTextFile(WorkingOutPath_Sort, ForReading, true)
    Set objWriteFile= objFileSys.OpenTextFile(WorkingOutPath_Time, ForWriting, true)

    Do Until objReadFile.AtEndOfStream
        strFullPath = objReadFile.ReadLine
        strRelativePath = Replace(strFullPath, strWorkingPath, "")

        If Not objRE Is Nothing Then
            If objRE.GetInstance.Test(strRelativePath) = False Then ' 無視するディレクトリ以外の場合に出力する
                Set objFile = objFileSys.GetFile(strFullPath)
                objWriteFile.WriteLine FormatDateTime(objFile.DateLastModified, 2) & " " & Right("0" &  FormatDateTime(objFile.DateLastModified, 3), 8) & " " & objFile.Name
            End If
        Else
                Set objFile = objFileSys.GetFile(strFullPath)
                objWriteFile.WriteLine FormatDateTime(objFile.DateLastModified, 2) & " " & Right("0" &  FormatDateTime(objFile.DateLastModified, 3), 8) & " " & objFile.Name
        End If
    Loop

    objReadFile.Close
    objWriteFile.Close


    ' HEADリポジトリディレクトリ内の対象ファイルの一覧を作成する
    Call ShellRun("cmd /c del " & HEADRepoOutPath_Dir & " " & HEADRepoOutPath_Sort & " " & HEADRepoOutPath_Time, 0)

    ' dir -> sort
    Call ShellRun("cmd /c dir /b /s " & strHEADRepoPath & "*.aspx " & strHEADRepoPath & "*.vb " & strHEADRepoPath & "*.aspx.resx " & strHEADRepoPath & "*.ascx " & strHEADRepoPath & "*.html " & strHEADRepoPath & "*.js " & strHEADRepoPath & "*.css " & strHEADRepoPath & "*.gif " & strHEADRepoPath & "*.jpg " & strHEADRepoPath & "*.png " & strHEADRepoPath & "*.xls > " & HEADRepoOutPath_Dir, 0)
    WScript.sleep(SLEEP) ' 少しスリープを入れないと実行結果が反映されない
    Call ShellRun("cmd /c sort "  & HEADRepoOutPath_Dir & " /O " & HEADRepoOutPath_Sort, 0)
    WScript.sleep(SLEEP) ' 少しスリープを入れないと実行結果が反映されない

    ' sort -> タイムスタンプ埋め込み
    Set objReadFile = objFileSys.OpenTextFile(HEADRepoOutPath_Sort, ForReading, true)
    Set objWriteFile= objFileSys.OpenTextFile(HEADRepoOutPath_Time, ForWriting, true)

    Do Until objReadFile.AtEndOfStream
        strFullPath = objReadFile.ReadLine
        strRelativePath = Replace(strFullPath, strHEADRepoPath, "")

        If Not objRE Is Nothing Then
            If objRE.GetInstance.Test(strRelativePath) = False Then ' 無視するディレクトリ以外の場合に出力する
                Set objFile = objFileSys.GetFile(strFullPath)
                objWriteFile.WriteLine FormatDateTime(objFile.DateLastModified, 2) & " " & Right("0" &  FormatDateTime(objFile.DateLastModified, 3), 8) & " " & objFile.Name
            End If
        Else
            Set objFile = objFileSys.GetFile(strFullPath)
            objWriteFile.WriteLine FormatDateTime(objFile.DateLastModified, 2) & " " & Right("0" &  FormatDateTime(objFile.DateLastModified, 3), 8) & " " & objFile.Name
        End If
    Loop

    objReadFile.Close
    objWriteFile.Close


    ' ワーキングディレクトリとHEADリポジトリディレクトリの内容を比較する
    Call ShellRun("cmd /c fc /LB100 " & WorkingOutPath_Time & " " & HEADRepoOutPath_Time & " > " & FcOutPath, 0)
    WScript.sleep(SLEEP) ' 少しスリープを入れないと実行結果が反映されない



    ' 前回のコミットログを表示
    Dim objHEADFolder: Set objHEADFolder = objFileSys.GetFolder(REPOSITORY_DIR & "\" & GetHEADRepository())
    WScript.echo vbCRLF & CreateCommitLog(objHEADFolder.Name) & vbCRLF

    ' 比較結果からファイル名をワーキングディレクトリとリポジトリディレクトリに振り分ける
    Set objReadFile = objFileSys.OpenTextFile(FcOutPath, ForReading, true)
    Dim FcOutWorking:  Set FcOutWorking = CreateObject("System.Collections.ArrayList")
    Dim FcOutHEADRepo: Set FcOutHEADRepo = CreateObject("System.Collections.ArrayList")
    Dim FcOutSwitch
    Dim count: count = 0

    Do Until objReadFile.AtEndOfStream
        Dim FcOutLine: FcOutLine = objReadFile.ReadLine
        
        If Ucase(Replace(FcOutLine, "***** ", "")) = Ucase(WorkingOutPath_Time) Then
            FcOutSwitch = FCOUT_WORKING
        ElseIf Ucase(Replace(FcOutLine, "***** ", "")) = Ucase(HEADRepoOutPath_Time) Then
            FcOutSwitch = FCOUT_HEAD
        ElseIf FcOutLine = "*****" Then
            Dim FcOutWorkingSplit, FcOutHEADRepoSplit
            Dim i, elem
            
            ' ワーキングディレクトリのファイルが、
            ' HEADリポジトリに存在して更新日時が違うものをmodifiedとして出力する
            ' HEADリポジトリに存在しない場合はnewfileとして出力する
            For i = 0 To FcOutWorking.Count - 1
                Dim modified: modified = False
                Dim newfile:  newfile = True
                FcOutWorkingSplit = Split(FcOutWorking(i), " ")
                For Each elem In FcOutHEADRepo
                    FcOutHEADRepoSplit = Split(elem, " ")
                    If FcOutWorkingSplit(2) = FcOutHEADRepoSplit(2) Then
                        newfile = False
                        If (FcOutWorkingSplit(0) <> FcOutHEADRepoSplit(0) Or FcOutWorkingSplit(1) <> FcOutHEADRepoSplit(1)) Then
                            modified = True
                            Exit For
                        End If
                    End If
                Next
                If modified = True Then
                    EchoChangesToBeCommited (count = 0), "  modified: " & FcOutWorking(i)
                    count = count + 1
                End If
                If newfile = True Then
                    EchoChangesToBeCommited (count = 0), " + newfile: " & FcOutWorking(i)
                    count = count + 1
                End If
            Next
            
            ' HEADリポジトリのファイルが
            ' ワーキングディレクトリに存在しない場合はdeletedとして出力する
            For i = 0 To FcOutHEADRepo.Count - 1
                Dim deleted: deleted = True
                FcOutHEADRepoSplit = Split(FcOutHEADRepo(i), " ")
                For Each elem In FcOutWorking
                    FcOutWorkingSplit = Split(elem, " ")
                    If FcOutHEADRepoSplit(2) = FcOutWorkingSplit(2) Then
                        deleted = False
                    End If
                Next
                If deleted = True Then
                    EchoChangesToBeCommited (count = 0), " - deleted: " & FcOutHEADRepo(i)
                    count = count + 1
                End If
            Next
            
            FcOutWorking.Clear
            FcOutHEADRepo.Clear
        ElseIf Trim(FcOutLine) <> "" Then
            ' 比較元(ワーキングディレクトリ)と比較先(HEADリポジトリディレクトリ)に振り分ける
            If FcOutSwitch = FCOUT_WORKING Then
                FcOutWorking.Add FcOutLine
            ElseIf FcOutSwitch = FCOUT_HEAD Then
                FcOutHEADRepo.Add FcOutLine
            End If
        End If
    Loop
    
    objReadFile.Close


    If count = 0 Then
        WScript.echo "nothing to commit"
    End If

End Sub

Function EchoChangesToBeCommited(blnChange, output)
    If blnChange = True Then
        WScript.echo "Changes to be commited:" & vbCRLF
        blnChange = False
    End If
    WScript.echo output
    
    ' Return
    EchoChangesToBeCommited = blnChange
End Function
