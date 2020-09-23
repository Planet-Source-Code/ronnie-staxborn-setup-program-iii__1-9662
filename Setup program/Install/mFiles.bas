Attribute VB_Name = "mFiles"
'
'     I've found this module on the net and found
'     it very good so I decided to use it
'     in this code.
'
'
'     Ronnie Staxborn
'
'
'     PS: Read below who maked this module.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                                                        ''
''   MFILES.BAS WAS DESIGNED AND WRITTEN BY TONY WILSON   ''
''   ==================================================   ''
''                                                        ''
'' This file was written with the intention of helping    ''
'' both experienced VB programmers, and begginers.  It    ''
'' provides easy file operations which are organised into ''
'' an easy to use index system with excellent examples.   ''
''                                                        ''
'' Tony Wilson: tonyscomp@europe.com                      ''
'' For more information execute mAbout() in Form_Load     ''
''                                                        ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' WORKING WITH TEXT FILES             EXAMPLES AS FOLLOWS
' =======================             ===================
' 1)  Open file for input             mOpenFile"C:\Testfile.txt", 1
' 2)  Write text                      mWriteText "This is a test"
' 3)  Write encrypted text      'still to be added
' 4)  Append text                     mAppendText "This is a test"
' 5)  Read text                       MyString = mReadLine()
' 6)  Read encrypted text       'still to be added
' 7)  Read all text                   MyString = mReadAll()
' 8)  Read encrypted all text   'still to be added
' 9)  Close                           mCloseFile()
' 10) File exists                     MyBoolean mFileExists("C:\windows\calc.exe")
' 11) Open file (shell)         'still to be added
' 12) Copy file                       mCopyFile "C:\Windows\calc.exe", "C:\Windows\Desktop\Calculator.exe", False
' 13) Move file                       mMoveFile "C:\windows\calc.exe", "C:\windows\desktop\temp\calc.exe"
' 14) Rename file                     nRenameFile "C:\Test.exe", "Tested.exe"
' 15) Unencrypt file            'still to be added
' 16) Kill file                       mKillFile("C:\testing.exe")
' 17) Get Directory of file           mGetDir "C:\windows\calc.exe"


' WORKING WITH FOLDERS
' ====================
' 1)  Create folder                   mCreateFolder ("C:\Temp\Testing\TempFolder1\")
' 2)  Open folder (shell)             mOpenFolder "C:\Windows\"
' 3)  Copy folder                     mCopyFolder"C:\Program Files\", "C:\Backup\Program Files\", True
' 4)  Copy folder + all sub folders   mCopySubFolders "C:\Program Files\", "C:\Backup\Program Files\", True
' 5)  Empty folder                    mEmptyFolder ("C:\Windows\Temp\")
' 6)  Delete folder                   mDeleteFolder ("C:\Windows\Temp\")
' 7)  Folder exists                   MsgBox mFolderExists ("C:\Windows\")
' 8)  Check Path                      MyFileName = mCheckPath(FilePath)


' WORKING WITH OTHER FILES
' ========================
' 1)  Open a program (shell)          mShellFile ("C:\Windows\Calc.exe", 3)
' 2)  Email someone                   mEmail "tonyscomp@europe.com"
' 3)  Open web page                   mWeb "www.geocities.com/helpingthepoor"

Dim fs, f
Dim mFileName As String
Dim mDestination
Dim mSource
Dim mOverwrite As Boolean
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Const mcstrExamplePath As String = "C:\VBSBSamp\"


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub mAbout()
    'Example: mAbout()
    MsgBox "Thank you for downloading my module, I do hope this will help you with your programming." + vbCrLf + vbCrLf + "I have downloaded a lot from this site, and I think its about time I gaves something " + vbCrLf + "back in return, so I set about writing this.  I wrote it the best I could, I tried to " + vbCrLf + "eliminate all possible errors and make it easy to use and understand." + vbCrLf + vbCrLf + "If you find any problems with this module please let me know, If you like this module, " + vbCrLf + "or find it helpful, please also let me know and reply by using the sites forward " + vbCrLf + "emailing service (or whatever you call it)." + vbCrLf + vbCrLf + "If you have any suggestions, or improvements to this file, please email me these so that " + vbCrLf + "I can update the uploaded version." + vbCrLf + vbCrLf + "I am currently looking for a good library of modules or code, do you know of any?" + vbCrLf + vbCrLf + "Thanks for reading, Tony Wilson: tonyscomp@europe.com", vbInformation + vbOKOnly
End Sub

Public Sub mOpenFile(mFileName, mReadWriteAppend As Integer)
    'Example: mOpenFile"C:\Testfile.txt", 1
    Dim mToDo
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Select Case mReadWriteAppend
        Case 0 'default
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.OpenTextFile(mFileName, ForReading, TristateFalse)
        Case 1 'for reading
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.OpenTextFile(mFileName, ForReading, TristateFalse)
        Case 2 'for writing
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.OpenTextFile(mFileName, ForWriting, True, TristateFalse)
        Case 3 'for appending
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.OpenTextFile(mFileName, ForAppending, TristateFalse)
    End Select
End Sub

Public Sub mWriteText(mTextWrite As String)
    'Example: mWriteText "This is a test"
    If err_filenotopen Then
        MsgBox "You tried to write text to a file which is not open." + vbCrLf + "Please use mOpenFile before attempting to write to a file", vbCritical + vbOKOnly, "Design Error"
    End If
    f.Write mTextWrite + vbCrLf
End Sub

Public Sub mAppendText(mTextAppend As String)
    'Example: mAppendText "This is a test"
    If err_filenotopen Then
        MsgBox "You tried to write text to a file which is not open." + vbCrLf + "Please use mOpenFile before attempting to write to a file", vbCritical + vbOKOnly, "Design Error"
    End If
    f.Write mTextAppend + vbCrLf
End Sub

Function mReadLine()
    'Example: MyString = mReadLine()
    If err_filenotopen Then
        MsgBox "You tried to write text to a file which is not open." + vbCrLf + "Please use mOpenFile before attempting to write to a file", vbCritical + vbOKOnly, "Design Error"
    End If
    mReadLine = f.readline
End Function

Function mReadAll()
    'Example: MyString = mReadAll()
    If err_filenotopen Then
        MsgBox "You tried to write text to a file which is not open." + vbCrLf + "Please use mOpenFile before attempting to write to a file", vbCritical + vbOKOnly, "Design Error"
    End If
    mReadAll = f.readall
End Function

Function mCloseFile()
    'Example: mCloseFile()
    If err_filenotopen Then
        MsgBox "You tried to write text to a file which is not open." + vbCrLf + "Please use mOpenFile before attempting to write to a file", vbCritical + vbOKOnly, "Design Error"
    End If
    f.Close
End Function

Function mFileExists(mFileName)
    'Example: MsgBox mFileExists("C:\windows\calc.exe")
    Set fs = CreateObject("Scripting.FileSystemObject")
    mFileExists = fs.FileExists(mFileName)
End Function

Public Sub mCopyFile(mFileName, mDestination, mOverwrite)
    'Example: mCopyFile "C:\Windows\calc.exe", "C:\Windows\Desktop\Calculator.exe", False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim mCopyFileExists
    mCheckPathFolders (mDestination)
    If mOverwrite = False Then
        mCopyFileExists = fs.FileExists(mDestination)
        If mCopyFileExists = True Then
            Exit Sub
        Else
            fs.copyfile mFileName, mDestination
        End If
    Else
        fs.copyfile mFileName, mDestination, True
    End If
End Sub
Function mCheckPathFolders(mPathToCheck As String)
    'This function will check that the folders in the path exist,
    'it will create any folders that do not exist in the path.
    Dim mCopyFileExistsFolder
    Dim mCopyFileLen
    Dim mCopyFileChar
    Set fs = CreateObject("Scripting.FileSystemObject")
    mCopyFileLen = 1
    Do Until mCopyFileLen = Len(mPathToCheck) + 1
        mCopyFileChar = Mid(mPathToCheck, mCopyFileLen, 1)
        If mCopyFileChar = "\" Then
            If fs.folderexists(Left(mPathToCheck, mCopyFileLen - 1)) = False Then
                fs.CreateFolder (Left(mPathToCheck, mCopyFileLen - 1))
            End If
        End If
        mCopyFileLen = mCopyFileLen + 1
    Loop
End Function

Public Sub mMoveFile(mFileName, mDestination)
    'Example: mMoveFile "C:\windows\calc.exe", "C:\windows\desktop\temp\calc.exe"
    Set fs = CreateObject("Scripting.FileSystemObject")
    mCheckPathFolders (mDestination)
    fs.MoveFile mFileName, mDestination
End Sub

Public Sub mRenameFile(mFileName, mNewName As String)
    'Example: nRenameFile "C:\Test.exe", "Tested.exe"
    Dim mRenameFileLen
    Dim mRenameFileChar
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(mFileName) = False Then
        Exit Sub
    End If
    mRenameFileLen = Len(mFileName) - 1
    Do Until mRenameFileChar = "\"
        mRenameFileChar = Mid(mFileName, mRenameFileLen, 1)
        mRenameFileLen = mRenameFileLen - 1
    Loop
    mNewName = Left(mFileName, mRenameFileLen - 1) + "\" + mNewName
    Name mFileName As mNewName   ' Move and rename file.
End Sub

Public Sub mKillFile(mFileName)
    'Example: mKillFile("C:\testing.exe")
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(mFileName) = False Then
        Exit Sub
    End If
    Kill mFileName
End Sub

Public Function mGetDir(mFileName)
    'Example: mGetDir "C:\Windows\calc.exe"
    Dim c As String
    Dim i As Integer
    i = Len(mFileName)
    Do Until c = "\"
        i = i - 1
        If i <= 0 Then
            mGetDir = "Not Known"
            Exit Function
        End If
        c = Mid(mFileName, i, 1)
        If c = "\" Then
            mGetDir = Left(mFileName, i)
        End If
    Loop
End Function

Public Sub mCreateFolder(mDestination)
    'Example: mCreateFolder "C:\Temp\Testing\TempFolder1\"
    mCheckPathFolders (mDestination)
End Sub

Public Sub mOpenFolder(mDestination)
    'Example: mOpenFolder "C:\Windows\"
    Call Shell("explorer " & mDestination, vbNormalFocus)
End Sub

Public Sub mCopyFolder(mSource, mDestination, mOverwrite)
    'Example mCopyFolder"C:\Program Files\", "C:\Backup\Program Files\", True
    Set fs = CreateObject("Scripting.FileSystemObject")
    mCheckPathFolders (mDestination)
    fs.CopyFolder mSource, mDestination, mOverwrite
End Sub

Public Sub mCopySubFolders(mSource, mDestination, mOverwrite)
    'Example mCopySubFolders "C:\Program Files\", "C:\Backup\Program Files\", True
    Set fs = CreateObject("Scripting.FileSystemObject")
    mCheckPathFolders (mDestination)
    mSource = mSource + "*"
    fs.CopyFolder mSource, mDestination, mOverwrite
End Sub

Public Sub mEmptyFolder(mDestination)
    'Example: mEmptyFolder ("C:\Windows\Temp\")
    Set fs = CreateObject("Scripting.FileSystemObject")
    mDestination = mDestination + "*"
    fs.DeleteFolder mDestination, True
    fs.deletefile mDestination, True
End Sub

Public Sub mDeleteFolder(mDestination)
    'Example: mDeleteFolder ("C:\Windows\Temp\")
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.DeleteFolder mDestination, True
End Sub

Function mFolderExists(mDestination)
    'Example: MsgBox mFolderExists ("C:\Windows\")
    Set fs = CreateObject("Scripting.FileSystemObject")
    mFolderExists = fs.folderexists(mDestination)
End Function

Public Function mCheckPath(mDestination)
    'Example: MyFileName = mCheckPath(FilePath)
    If Right(mDestination, 1) = "\" Then
        mCheckPath = mDestination
    Else
        mCheckPath = mDestination & "\"
    End If
End Function



Function mGetIniValue(mSearchFor As String, mDestination)
    'Example: MsgBox mGetIniValue ("Caption= ", "app.path & "\MyFile.ini")
    Dim mTempLine As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(mDestination) = False Then
        MsgBox mDestination & " file does not exist!", vbCritical + vbOKOnly, "File Access Error"
        Exit Function
    End If
    Open mDestination For Input As #1
        Do While Not EOF(1)
            Line Input #1, mTempLine
            If Left(mTempLine, Len(mSearchFor)) = mSearchFor Then
                mGetIniValue = Mid(mTempLine, Len(mSearchFor))
                Exit Function
            End If
        Loop
    Close #1
End Function

Public Sub mShellFile(mFileName, mMaxNormalMin As Integer)
    'Example: mShellFile ("C:\Windows\Calc.exe", 3)
    Dim RetVal

    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(mFileName) = False Then
        Exit Sub
    End If
    Select Case mMaxNormalMin
        Case 0
            mMaxNormalMin = 1
        Case 1
            mMaxNormalMin = 3
        Case 2
            mMaxNormalMin = 1
        Case 3
            mMaxNormalMin = 0
    End Select
    
    RetVal = Shell(mFileName, mMaxNormalMin)
End Sub

Public Sub mEmail(mEmailAddress As String)
    'Example: mEmail "tonyscomp@europe.com"
    mEmailAddress = "mailto:" & mEmailAddress
    ShellExecute hwnd, "open", mEmailAddress, vbNullString, vbNullString, SW_SHOW
End Sub

Public Sub mWeb(mWebAddress As String)
    'Example: mWeb "www.geocities.com/helpingthepoor"
    ShellExecute hwnd, "open", mWebAddress, vbNullString, vbNullString, SW_SHOW
End Sub

Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)


    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If
End Function


Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetSystemDirectory(t, Len(t))
    SysDir = Left(t, i)


    If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
        SysDir = SysDir & "\"
    ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
        SysDir = Left(SysDir, Len(SysDir) - 1)
    End If
End Function
