Option Explicit

Const DEBUGMODE = false
Const WORK_DIR = ".\"
Const REPOSITORY = ".bk"
Const REPOSITORY_DIR = ".\.bk"  ' WORK_DIR & REPOSITORY
Const TRACKFILE = ".track"

' OpenTextFile
Const ForReading = 1    ' 読み取りモード
Const ForAppending = 8  ' 追記モード

' コマンド
Dim objParm, strCmd, strCmdOpt, strCmdTarget
' ファイルシステム
Dim objFileSys


Call Main

' Main   ------------------------------------------------------------------------------------------
Sub Main
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    Set objParm = Wscript.Arguments
    If objParm.Count > 0 Then
        strCmd = LCase(objParm(0))
    Else
        Call CommandDocument
        Exit Sub
    End If

    ' コマンドの振り分け
    Select Case strCmd
        Case "init"
            Call Initialize
        Case "add"
            Call addTrackfile
        Case "tracked"
            Call Tracked
        Case "status"
            Call Status
        Case "commit"
            Call Commit
        Case "log"
            Call CommitLog
        Case Else
            EchoError "Unknown"
    End Select

    Set objFileSys = Nothing
End Sub


' Command Document  -------------------------------------------------------------------------------
Sub CommandDocument
    WScript.echo "BK Command"
    WScript.echo "  init    - リポジトリの初期化"
    WScript.echo "  add     - 管理対象ファイルを追加"
    WScript.echo "  tracked - 管理対象ファイルの確認"
    WScript.echo "  status  - 最終コミットからの変更を表示"
    WScript.echo "  commit  - 管理対象ファイルをコミット"
    WScript.echo "  log     - コミットログの表示"
End Sub


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


' Add    ------------------------------------------------------------------------------------------
Sub addTrackfile
    Dim path
    Dim objFile

    If objParm.Count < 2 Then
        EchoUsage "BK Add ""File-name"""
    ELse
        ' 対象ファイルを追加する
        strCmdTarget = objParm(1)

        WScript.echo "BK Add"

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


' Tracked --------------------------------------------------------------------------------------
Sub Tracked
    Dim intTrackedFiles


    If objFileSys.FolderExists(REPOSITORY_DIR) Then

        ' 作業フォルダ内のサブフォルダを探す
        intTrackedFiles = TrackedFolders(WORK_DIR, "")

        ' 作業フォルダ内のファイルを探す
        intTrackedFiles = intTrackedFiles + TrackedFiles(WORK_DIR, "")
        
        WScript.echo vbCRLF & "Tracked " & intTrackedFiles & " File(s)"

    End If

End Sub


Function TrackedFolders(path, copyDir)
    Dim objTempFolder, subFolder
    Dim relativePath
    Dim intTrackedFiles


    Set objTempFolder = objFileSys.GetFolder(path)

    For Each subFolder In objTempFolder.SubFolders
        If IsTracked(subFolder) Then
            relativePath = Replace(subFolder, objFileSys.GetAbsolutePathName(WORK_DIR) & "\", "")
            WScript.echo "  " & relativePath

            ' コピーする
            If copyDir <> "" Then
                Call CreateFolder(copyDir & "\" & relativePath)
            End If
            
            intTrackedFiles = intTrackedFiles + TrackedFiles(subFolder, copyDir)
            
            ' 再帰呼び出し
            intTrackedFiles = intTrackedFiles + TrackedFolders(subFolder, copyDir)
        End If
    Next

    Set objTempFolder = Nothing

    ' Return
    TrackedFolders = intTrackedFiles

End Function

Function TrackedFiles(path, copyDir)
    Dim objFile
    Dim relativePath
    Dim intTrackedFiles


    For Each objFile In objFileSys.GetFolder(path).Files
        If IsTracked(objFile.Name) Then
            relativePath = Replace(objFile.path, objFileSys.GetAbsolutePathName(WORK_DIR) & "\", "")
            WScript.echo "  " & relativePath

            ' コピーする
            If copyDir <> "" Then
                objFile.Copy copyDir & "\" & relativePath
            End If

            intTrackedFiles = intTrackedFiles + 1
        End If
    Next

    ' Return
    TrackedFiles = intTrackedFiles

End Function


' Status ------------------------------------------------------------------------------------------
Sub Status
    Dim objWorkFolder, objWorkFile
    Dim objHeadFolder, objHeadFile
    Dim intAddedFiles, intModifiedFiles, intDeletedFiles
    Dim blnChanges
    

    Set objWorkFolder = objFileSys.GetFolder(WORK_DIR)
    Set objHeadFolder = objFileSys.GetFolder(REPOSITORY_DIR & "\" & GetRepositoryHead())

    blnChanges = False


    ' 前回のコミットログを表示
    WScript.echo vbCRLF & CreateCommitLog(objHeadFolder.Name) & vbCRLF

    ' リポジトリにもある既存ファイル
    intModifiedFiles = 0
    For Each objWorkFile In objFileSys.GetFolder(objWorkFolder.Path).Files
        If IsTracked(objWorkFile.Name) Then
        
            If objFileSys.FileExists(objHeadFolder.Path & "\" & objWorkFile.Name) Then
                Set objHeadFile = objFileSys.GetFIle(objHeadFolder.Path & "\" & objWorkFile.Name)
                If objWorkFile.DateLastModified <> objHeadFile.DateLastModified Then
                    blnChanges = EchoChangesToBeCommited(blnChanges)
                    intModifiedFiles = intModifiedFiles + 1

                    WScript.echo "  modified: " & LogFormatDateTime(objWorkFile.DateLastModified) & " - " & objWorkFile.Name
                End If
            End If

        End If
    Next

    ' リポジトリにない新規ファイル
    intAddedFiles = 0
    For Each objWorkFile In objFileSys.GetFolder(objWorkFolder.Path).Files
        If IsTracked(objWorkFile.Name) Then
        
            If objFileSys.FileExists(objHeadFolder.Path & "\" & objWorkFile.Name) = False Then
                blnChanges = EchoChangesToBeCommited(blnChanges)
                intAddedFiles = intAddedFiles + 1

                WScript.echo "  new file: " & LogFormatDateTime(objWorkFile.DateLastModified) & " - " & objWorkFile.Name
            End If

        End If
    Next
    
    ' リポジトリにしかない削除されたファイル
    intDeletedFiles = 0
    For Each objHeadFile In objFileSys.GetFolder(objHeadFolder.Path).Files
'        If IsTracked(objHeadFile.Name) Then
        
            If objFileSys.FileExists(objWorkFolder.Path & "\" & objHeadFile.Name) = False Then
                blnChanges = EchoChangesToBeCommited(blnChanges)
                intDeletedFiles = intDeletedFiles + 1

                WScript.echo "  deleted:  " & LogFormatDateTime(objHeadFile.DateLastModified) & " - " & objHeadFile.Name
            End If

'        End If
    Next

    If intModifiedFiles + intAddedFiles + intDeletedFiles > 0 Then
'        WScript.echo vbCRLF & "  " &  "modified: " & intModifiedFiles & " file(s), " & _
'                                      "new file: " & intAddedFiles & " file(s), " & _
'                                      "deleted:  " & intDeletedFiles & " file(s)"
    Else
        WScript.echo "nothing to commit"
    End If

End Sub

Function EchoChangesToBeCommited(blnChange)
    If blnChange = False Then
        WScript.echo "Changes to be commited:" & vbCRLF
        blnChange = True
    End If
    
    ' Return
    EchoChangesToBeCommited = blnChange
End Function


' Commit ------------------------------------------------------------------------------------------
Sub Commit
    Dim dir
    Dim commitLog
    Dim intCommitFiles


    strCmdOpt = objParm(1)

    Select Case LCase(strCmdOpt)
    Case "-m"
        If objParm.Count < 3 Then
            EchoUsage "bk commit -m ""commit-message"""
            WScript.Quit
        ELse
            strCmdOpt = objParm(1)
            strCmdTarget = objParm(2)
        End If

        If objFileSys.FolderExists(REPOSITORY_DIR) Then
            ' コミットフォルダを作成（yyyymmdd_hhmmss_CommitMessage）
            commitLog = FormatDateTime(Now, 2) & "_" & FormatDateTime(Now, 3) & "_" & strCmdTarget
            commitLog = Replace(commitLog, "/", "")  ' フォルダ名に使用できない文字を削除
            commitLog = Replace(commitLog, ":", "")
            dir = REPOSITORY_DIR & "\" & commitLog
            If CreateFolder(dir) Then

                ' コミットログを表示
                WScript.echo vbCRLF & CreateCommitLog(commitLog) & vbCRLF

	            ' ファイルをコミット（カレントディレクトのファイルをコピー）
                ' 作業フォルダ内のサブフォルダを探す
                intCommitFiles = TrackedFolders(WORK_DIR, dir)

                ' 作業フォルダ内のファイルを探す
                intCommitFiles = intCommitFiles + TrackedFiles(WORK_DIR, dir)

	            WScript.echo vbCRLF & intCommitFiles & " files changed"
            End If

        End If
    End Select

End Sub


' Log ---------------------------------------------------------------------------------------------
Sub CommitLog
    Dim objRepositories, strRepository


    Set objRepositories = GetRepositories()
    If objRepositories.Count > 0 Then
        ' ディレクトリ名(日時順)でソートする
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


' リポジトリを取得
Function GetRepositories
    Dim objFolder
    Dim objArrayList, objItem


    If objFileSys.FolderExists(REPOSITORY_DIR) Then
        ' フォルダ名をArrayListに格納する
        Set objFolder = objFileSys.GetFolder(REPOSITORY_DIR)
        Set objArrayList = CreateObject("System.Collections.ArrayList")
        For Each objItem In objFolder.SubFolders
            objArrayList.Add objItem.Name
        Next
        
        ' Return
        Set GetRepositories = objArrayList
    End If
    
End Function


' リポジトリから最終コミットを取得
Function GetRepositoryHead
    Dim objRepositories


    Set objRepositories = GetRepositories()
    If objRepositories.Count > 0 Then
        objRepositories.Reverse()   ' 降順ソート
        
        ' Return
        GetRepositoryHead = objRepositories(0)
    End If

End Function


' リポジトリのフォルダ名を分解してコミットログを生成する
Function CreateCommitLog(strRepository)
    Dim logs, logDate, logTime, logMessage
    Dim i


    logs = Split(strRepository, "_")
    logDate = DateSerial(Mid(logs(0), 1, 4), Mid(logs(0), 5, 2), Mid(logs(0), 7, 2))
    logTime = TimeSerial(Mid(logs(1), 1, 2), Mid(logs(1), 3, 2), Mid(logs(1), 5, 2))
    logMessage = logs(2)
    ' 分解されてしまったコミットメッセージを復元する
    For i = 3 To UBound(logs)
        logMessage = logMessage & "_" & logs(i)
    Next

    ' Return
    CreateCommitLog = LogFormatDateTime(logDate + logTime) & " - " & _
                      logMessage

End Function


' 日付をコミットログ用の書式にする
Function LogFormatDateTime(d)
    ' Return
    ' yyyy/mm/dd(曜) hh:mm:ss
    LogFormatDateTime = FormatDateTime(d, 2) & "(" & WeekdayChar(Weekday(d)) & ") " & FormatDateTime(d, 3)
    
End function


' コミットログ用の曜日
Function WeekdayChar(w)
    Dim weekdays(7)

    
    weekdays(1) = "日"
    weekdays(2) = "月"
    weekdays(3) = "火"
    weekdays(4) = "水"
    weekdays(5) = "木"
    weekdays(6) = "金"
    weekdays(0) = "土"
    
    ' Return
    WeekdayChar = weekdays(w)
    
End Function


' フォルダを作成する
Function CreateFolder(dir)
    Dim objFolder


    ' フォルダが無かったら作成する
    If objFileSys.FolderExists(dir) = False Then
        objFileSys.CreateFolder(dir)
	    If objFileSys.FolderExists(dir) Then
'	        Set objFolder = objFileSys.GetFolder(dir)
'	        WScript.echo "Create Folder: " & objFolder.Path
            ' Return
	        CreateFolder = True
	    Else
	        ' Return
	        CreateFolder = False
	    End If
	Else
'	        Set objFolder = objFileSys.GetFolder(dir)
'	        WScript.echo "Exists Folder: " & objFolder.Path
            ' Return
	        CreateFolder = True
    End If

    Set objFolder = Nothing

End Function


' 管理対象かを判定する
Function IsTracked(filename)
    Dim pathTrackFile
    Dim tracked
    Dim objFile
    Dim objRE
    Dim strLine


    pathTrackFile = REPOSITORY_DIR & "\" & TRACKFILE
    
    Call DebugPrint(objFileSys.GetFileName(filename))
    
    If objFileSys.GetFileName(filename) = REPOSITORY Then
        tracked = false ' コミット用ディレクトリは対象外とする
    ELseIf objFileSys.FolderExists(filename) Then
        tracked = true
    ELseIf objFileSys.FileExists(pathTrackFile) Then
        ' 正規表現で対象ファイルであるかを判定する
        Set objFile = objFileSys.OpenTextFile(pathTrackFile, ForReading, true)
        Set objRE = CreateObject("VBScript.RegExp") ' 正規表現

        tracked = false
        Do Until objFile.AtEndOfStream
            strLine = objFile.ReadLine
            strLine = Replace(strLine, "*", ".*")
            strLine = Replace(strLine, "?", ".")
            objRE.Pattern = "^" & strLine & "$"
            If objRE.Test(filename) Then
                tracked = true
                Exit Do
            End If
        Loop

        objFile.Close
        Set objRE = Nothing
        Set objFile = Nothing
    Else
        ' Trackファイルが無い場合は全ファイルを対象とする
        tracked = true
    End If

    IsTracked = tracked

End Function


' コンソール出力用
Sub EchoSuccess(str)
    WScript.echo str
End Sub

Sub EchoUsage(str)
    WScript.echo "Usage: " & str
End Sub

Sub EchoError(str)
    WScript.echo "Error: " & str
End Sub


' デバッグ用
Sub DebugPrint(str)
    If DEBUGMODE Then
        WScript.echo "Degbug:" & str
    End If
End Sub
