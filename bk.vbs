Option Explicit

Const DEBUGMODE = false
Const WORK_DIR = ".\"
Const REPOSITORY = ".bk"
Const REPOSITORY_DIR = ".\.bk"  ' WORK_DIR & REPOSITORY
Const TRACKFILE = ".track"

' OpenTextFile
Const ForReading = 1    ' �ǂݎ�胂�[�h
Const ForAppending = 8  ' �ǋL���[�h

' �R�}���h
Dim objParm, strCmd, strCmdOpt, strCmdTarget
' �t�@�C���V�X�e��
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

    ' �R�}���h�̐U�蕪��
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
    WScript.echo "  init    - ���|�W�g���̏�����"
    WScript.echo "  add     - �Ǘ��Ώۃt�@�C����ǉ�"
    WScript.echo "  tracked - �Ǘ��Ώۃt�@�C���̊m�F"
    WScript.echo "  status  - �ŏI�R�~�b�g����̕ύX��\��"
    WScript.echo "  commit  - �Ǘ��Ώۃt�@�C�����R�~�b�g"
    WScript.echo "  log     - �R�~�b�g���O�̕\��"
End Sub


' Init   ------------------------------------------------------------------------------------------
Sub Initialize

    ' ���|�W�g���p�t�H���_���쐬����
    If objFileSys.FolderExists(REPOSITORY_DIR) = False Then
        If CreateFolder(REPOSITORY_DIR) Then
            EchoSuccess "Initialized empty BK repository in " & objFileSys.GetAbsolutePathName(REPOSITORY_DIR)
        End If
    Else
        ' ���łɏ���������Ă���
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
        ' �Ώۃt�@�C����ǉ�����
        strCmdTarget = objParm(1)

        WScript.echo "BK Add"

        Set objFile = objFileSys.OpenTextFile(REPOSITORY_DIR & "\" & TRACKFILE, ForAppending, true)
        objFile.WriteLine strCmdTarget
        objFile.Close
        Set objFile = Nothing
    End If

    ' �Ώۃt�@�C���̓��e��񋓂���
    path = REPOSITORY_DIR & "\" & TRACKFILE
    If objFileSys.FileExists(path) Then
    
        Set objFile = objFileSys.OpenTextFile(path, ForReading, true)

        Do Until objFile.AtEndOfStream      ' ���̓t�@�C���̏I�[�܂ŌJ��Ԃ�
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

        ' ��ƃt�H���_���̃T�u�t�H���_��T��
        intTrackedFiles = TrackedFolders(WORK_DIR, "")

        ' ��ƃt�H���_���̃t�@�C����T��
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

            ' �R�s�[����
            If copyDir <> "" Then
                Call CreateFolder(copyDir & "\" & relativePath)
            End If
            
            intTrackedFiles = intTrackedFiles + TrackedFiles(subFolder, copyDir)
            
            ' �ċA�Ăяo��
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

            ' �R�s�[����
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


    ' �O��̃R�~�b�g���O��\��
    WScript.echo vbCRLF & CreateCommitLog(objHeadFolder.Name) & vbCRLF

    ' ���|�W�g���ɂ���������t�@�C��
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

    ' ���|�W�g���ɂȂ��V�K�t�@�C��
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
    
    ' ���|�W�g���ɂ����Ȃ��폜���ꂽ�t�@�C��
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
            ' �R�~�b�g�t�H���_���쐬�iyyyymmdd_hhmmss_CommitMessage�j
            commitLog = FormatDateTime(Now, 2) & "_" & FormatDateTime(Now, 3) & "_" & strCmdTarget
            commitLog = Replace(commitLog, "/", "")  ' �t�H���_���Ɏg�p�ł��Ȃ��������폜
            commitLog = Replace(commitLog, ":", "")
            dir = REPOSITORY_DIR & "\" & commitLog
            If CreateFolder(dir) Then

                ' �R�~�b�g���O��\��
                WScript.echo vbCRLF & CreateCommitLog(commitLog) & vbCRLF

	            ' �t�@�C�����R�~�b�g�i�J�����g�f�B���N�g�̃t�@�C�����R�s�[�j
                ' ��ƃt�H���_���̃T�u�t�H���_��T��
                intCommitFiles = TrackedFolders(WORK_DIR, dir)

                ' ��ƃt�H���_���̃t�@�C����T��
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
        ' �f�B���N�g����(������)�Ń\�[�g����
        objRepositories.Sort()     ' �����\�[�g

        ' ���t���ɃR�~�b�g���O���o�͂���
        For Each strRepository In objRepositories
            WScript.Echo CreateCommitLog(strRepository)
        Next
    Else
        EchoError "Not a bk repository : " & REPOSITORY
    End If

    Set objRepositories = Nothing

End Sub


' ���|�W�g�����擾
Function GetRepositories
    Dim objFolder
    Dim objArrayList, objItem


    If objFileSys.FolderExists(REPOSITORY_DIR) Then
        ' �t�H���_����ArrayList�Ɋi�[����
        Set objFolder = objFileSys.GetFolder(REPOSITORY_DIR)
        Set objArrayList = CreateObject("System.Collections.ArrayList")
        For Each objItem In objFolder.SubFolders
            objArrayList.Add objItem.Name
        Next
        
        ' Return
        Set GetRepositories = objArrayList
    End If
    
End Function


' ���|�W�g������ŏI�R�~�b�g���擾
Function GetRepositoryHead
    Dim objRepositories


    Set objRepositories = GetRepositories()
    If objRepositories.Count > 0 Then
        objRepositories.Reverse()   ' �~���\�[�g
        
        ' Return
        GetRepositoryHead = objRepositories(0)
    End If

End Function


' ���|�W�g���̃t�H���_���𕪉����ăR�~�b�g���O�𐶐�����
Function CreateCommitLog(strRepository)
    Dim logs, logDate, logTime, logMessage
    Dim i


    logs = Split(strRepository, "_")
    logDate = DateSerial(Mid(logs(0), 1, 4), Mid(logs(0), 5, 2), Mid(logs(0), 7, 2))
    logTime = TimeSerial(Mid(logs(1), 1, 2), Mid(logs(1), 3, 2), Mid(logs(1), 5, 2))
    logMessage = logs(2)
    ' ��������Ă��܂����R�~�b�g���b�Z�[�W�𕜌�����
    For i = 3 To UBound(logs)
        logMessage = logMessage & "_" & logs(i)
    Next

    ' Return
    CreateCommitLog = LogFormatDateTime(logDate + logTime) & " - " & _
                      logMessage

End Function


' ���t���R�~�b�g���O�p�̏����ɂ���
Function LogFormatDateTime(d)
    ' Return
    ' yyyy/mm/dd(�j) hh:mm:ss
    LogFormatDateTime = FormatDateTime(d, 2) & "(" & WeekdayChar(Weekday(d)) & ") " & FormatDateTime(d, 3)
    
End function


' �R�~�b�g���O�p�̗j��
Function WeekdayChar(w)
    Dim weekdays(7)

    
    weekdays(1) = "��"
    weekdays(2) = "��"
    weekdays(3) = "��"
    weekdays(4) = "��"
    weekdays(5) = "��"
    weekdays(6) = "��"
    weekdays(0) = "�y"
    
    ' Return
    WeekdayChar = weekdays(w)
    
End Function


' �t�H���_���쐬����
Function CreateFolder(dir)
    Dim objFolder


    ' �t�H���_������������쐬����
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


' �Ǘ��Ώۂ��𔻒肷��
Function IsTracked(filename)
    Dim pathTrackFile
    Dim tracked
    Dim objFile
    Dim objRE
    Dim strLine


    pathTrackFile = REPOSITORY_DIR & "\" & TRACKFILE
    
    Call DebugPrint(objFileSys.GetFileName(filename))
    
    If objFileSys.GetFileName(filename) = REPOSITORY Then
        tracked = false ' �R�~�b�g�p�f�B���N�g���͑ΏۊO�Ƃ���
    ELseIf objFileSys.FolderExists(filename) Then
        tracked = true
    ELseIf objFileSys.FileExists(pathTrackFile) Then
        ' ���K�\���őΏۃt�@�C���ł��邩�𔻒肷��
        Set objFile = objFileSys.OpenTextFile(pathTrackFile, ForReading, true)
        Set objRE = CreateObject("VBScript.RegExp") ' ���K�\��

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
        ' Track�t�@�C���������ꍇ�͑S�t�@�C����ΏۂƂ���
        tracked = true
    End If

    IsTracked = tracked

End Function


' �R���\�[���o�͗p
Sub EchoSuccess(str)
    WScript.echo str
End Sub

Sub EchoUsage(str)
    WScript.echo "Usage: " & str
End Sub

Sub EchoError(str)
    WScript.echo "Error: " & str
End Sub


' �f�o�b�O�p
Sub DebugPrint(str)
    If DEBUGMODE Then
        WScript.echo "Degbug:" & str
    End If
End Sub
