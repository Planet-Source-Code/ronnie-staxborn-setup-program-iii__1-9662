Attribute VB_Name = "modSelfExtract"
Public iFilez As Integer
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Sub SelfExtract()

On Error Resume Next
Dim test
Dim test2
Dim Size As String
Dim iFreeFile As Integer
Dim iName As String
Dim rPath As String
Dim TheFile As String
Dim rWelcome As String
Dim rAbout As String



iFreeFile = FreeFile
rPath = App.Path
If Mid(rPath, Len(rPath)) <> "\" Then rPath = rPath & "\"

curPOS = 0
i = 0
Do
i = i + 1
    Open rPath & App.EXEName & ".exe" For Binary As iFreeFile
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 41 - 10 + curPOS
    iName = String(40, Chr(0))
    Get iFreeFile, , iName
    
    DoEvents
    iName = Replace$(iName, vbCr, "")
    frmPreinstall.lblfiles.Caption = "Extracting " & iName & "..."
    frmPreinstall.lblfiles.Refresh
        
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 11 + curPOS
    Size = String(10, Chr(0))
    Get iFreeFile, , Size
    DoEvents
    Size = CCur(Size)
   DoEvents
    Seek #iFreeFile, LOF(iFreeFile) - 51 - Size - (256 * 2) - 5 + curPOS
    TheFile = String(Size, Chr(0))
    Get iFreeFile, , TheFile
    DoEvents
    Close iFreeFile
    FFile = FreeFile
    
    test = GetTempPathName & "\staxborn\"
    test2 = test & iName
    Open test2 For Binary Access Write As #FFile
        Put #FFile, , TheFile
    DoEvents
    Close #FFile
    DoEvents
    curPOS = curPOS - Size - 50
DoEvents
Loop Until i >= iFilez

Exit Sub

Err:

Result = MsgBox("An error occured. Header may be damaged." _
    & vbCrLf & "Do you want to abort/retry?", _
    vbAbortRetryIgnore + vbExclamation, "Error")

If Result = vbRetry Then
    Resume
ElseIf Result = vbIgnore Then
    Resume Next
ElseIf Result = vbAbort Then
    End
End If

End Sub

Sub Unzipsetup()

On Error GoTo vbErrorHandler


'
' Unzip the ZIPTEST.ZIP file to the Windows Temp Directory
'
    Dim oUnZip As CGUnzipFiles
    
    Set oUnZip = New CGUnzipFiles
    
    With oUnZip
'
' What Zip File ?
'
        .ZipFileName = GetTempPathName & "staxborn" & "\DummyIns.ZIP"
'
' Where are we zipping to ?
'
        .ExtractDir = frmInstaller.txtdirectory.Text
'
' Keep Directory Structure of Zip ?
'
        .HonorDirectories = False
'
' Unzip and Display any errors as required
'
        If .Unzip <> 0 Then
            MsgBox .GetLastMessage
        End If
    End With
    frmInstaller.lblExtract.Caption = "Installing..."
    Set oUnZip = Nothing
        
    Kill GetTempPathName & "staxborn" & "\DummyIns.ZIP"
    
    'MsgBox "The installation is completed", vbInformation, "All went just fine..."
    'CloseAll
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Extracting fail..." & " " & Err.Description
End
End Sub


Sub CloseAll()

Kill GetTempPathName & "\Unzip32.dll"

End Sub

Public Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lREt As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lREt = GetTempPath(255, sBuffer)
    
    If lREt > 0 Then
        sBuffer = Left$(sBuffer, lREt)
    End If
    GetTempPathName = sBuffer
    
End Function

