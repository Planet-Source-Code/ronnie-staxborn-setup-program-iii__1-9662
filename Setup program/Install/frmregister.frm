VERSION 5.00
Begin VB.Form frmregister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registering the dependencies:"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   3840
   End
   Begin VB.TextBox Status 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   6255
   End
   Begin VB.ListBox lstregister 
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   2295
      Visible         =   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3120
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
      Visible         =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All files have been registered to you computer :)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   4815
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If there are any dependencies in the installer, It will now be registered in your system folder. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' *******************************************************
' *          INSTALLER PROGRAM by Ronnie Staxborn       *
' *                                                     *
' *    Thanx to Vasilis Sagonas and Chris Eastwood      *
' *    for helping me with the code.                    *
' *    If you like to program plz vote and if you want  *
' *    to contact me plz write to rompa@hem.passagen.se *
' *                                                     *
' *******************************************************
'


Option Explicit

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpdata As Any) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
  (ByVal lpLibFileName As String) As Long
  
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
   
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, _
   ByVal dwExitCode As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
   
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim mCompanyName As String
Dim mProductVersion As String
Dim RegFlag As Boolean
Dim UnregFlag As Boolean
Dim mresult
Dim gcdg As Object
Dim x
Dim entry
Dim entry2



Private Sub Command1_Click()

frmInstaller.lblExtract.Caption = "Installation is completed..."
frmInstaller.cmdCancel.Caption = "CLOSE"

Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo Tjohoo

' Use the lstregisterBox to select only the ocx, dll and the ole systemfiles.
File1.Path = (GetTempPathName & "staxborn\")
File1.Pattern = "*.ocx;*.dll;*.ole"

'Here I move to contents in the File1 Box to a listbox include to System directory for easy register.
For x = 0 To File1.ListCount - 1
entry = SysDir(True) & File1.List(x)
lstregister.AddItem entry
Next

Show
timedPause (2)
For x = 0 To lstregister.ListCount - 1
entry2 = lstregister.List(x)
DispProdVersion (entry2)
RegUnReg entry2
Next
 
Label3.Visible = True
timedPause (2)
Command1.Enabled = True
Exit Sub

Tjohoo:
MsgBox "something is wrong, check the path"

End Sub

Private Sub DispProdVersion(inFile As String)
    If Not GetFileInfo(inFile) Then
       
        Status.Text = "(No Product Version available for this file)"
    Else
      
       ' Status.Text = "Company Name:  " & mCompanyName & "Product Version:  " & mProductVersion
    End If
End Sub

Private Function GetFileInfo(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim lInfoSize As Long
    Dim lpHandle As Long
    Dim strFileInfoString As String
    Dim i As Integer
    
    GetFileInfo = False                                ' Assume
    
     ' GetFileVersionInfoSize determines if system can obtain version info
     ' about the specified file.  If yes, it returns its size in bytes and
     ' a handle to the data.
    lpHandle = 0
    lInfoSize = GetFileVersionInfoSize(inFileSpec, lpHandle)
    If lInfoSize = 0 Then
        Exit Function
    End If

     ' We pass the file name, size(ignored), size of buffer and the buffer of
     ' version info to GetFileVersionInfo, which will fill the buffer with
     ' version info about the file. (Modified here).
    strFileInfoString = String(lInfoSize, 0)
    mresult = GetFileVersionInfo(ByVal inFileSpec, 0&, ByVal lInfoSize, _
          ByVal strFileInfoString)
    If mresult = 0 Then
        Exit Function
    End If

     ' We now have a block of version data, in an unreadable format though. If you
     ' wish, you may check the existence of "StringFileInfo" with InStr function.
     ' Normally we must call VerQueryValue to read selected pieces of data of the
     ' above, with arguments such as "\VarFileInfo\Translation" or "\StringFileInfo
     ' \lang-codepage\string-name" where lang-codepage is a code which has yet to be
     ' obtained from first 2 words(high-low) returned by "\VarFileInfo\Translation"
     ' from the strFileInfoString (and padded to fixed 8-digit), and string-name is
     ' one of predefined string names such as "CompanyName" & "FileDescription", etc.
     ' However, the following simple alternative is OK for our purpose.

     mCompanyName = ""
     mProductVersion = ""
     i = InStr(strFileInfoString, "CompanyName")
     If i > 0 Then
         i = i + 12
         mCompanyName = Mid$(strFileInfoString, i, 21)
     End If
     i = InStr(strFileInfoString, "FileDescription")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "FileVersion")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "InternalName")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "LegalCopyright")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "OriginalFilename")
     If i > 0 Then
         i = i + 20
     End If
     i = InStr(strFileInfoString, "ProductName")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "ProductVersion")
     If i > 0 Then
         i = i + 16
         mProductVersion = Mid$(strFileInfoString, i)
     End If

     If Trim(mProductVersion) <> "" Then
         GetFileInfo = True
     End If
End Function
    
Private Sub RegUnReg(ByVal inFileSpec As String, Optional inHandle As String = "")
    On Error Resume Next
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread
    
      ' Load the control DLL, i. e. map the specified DLL file into the
      ' address space of the calling process
    lLib = LoadLibrary(inFileSpec)
    If lLib = 0 Then
         ' e.g. file not exists or not a valid DLL file
        Status.Text = Status.Text & "Failure loading control DLL" & vbCrLf
        Exit Sub
    End If
    
      ' Find and store the DLL entry point, i.e. obtain the address of the
      ' “DllRegisterServer” or "DllUnregisterServer" function (to register
      ' or deregister the server’s components in the registry).
      '
    If inHandle = "" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf inHandle = "U" Or inHandle = "u" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
       Status.Text = Status.Text & "Unknown command handle" & vbCrLf
        Exit Sub
    End If
    If lpDLLEntryPoint = vbNull Then
        GoTo earlyExit1
    End If
    
    Screen.MousePointer = vbHourglass
    
      ' Create a thread to execute within the virtual address space of the calling process
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
      ' Use WaitForSingleObject to check the return state (i) when the specified object
      ' is in the signaled state or (ii) when the time-out interval elapses.  This
      ' function can be used to test Process and Thread.
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
      ' We don't call the dangerous TerminateThread(); after the last handle
      ' to an object is closed, the object is removed from the system.
    CloseHandle mThread
    FreeLibrary lLib
    
    Screen.MousePointer = vbDefault
    Status.Text = Status.Text & vbCrLf & "Registration completed: " & entry2 & " from: " & mCompanyName ' " v." & mProductVersion & " completed"
    'Resume Next
    Exit Sub
    
    
earlyExit1:
    Screen.MousePointer = vbDefault
    Status.Text = Status.Text & vbCrLf & "Process failed in obtaining entry point or creating thread for: " & entry & " from: " & mCompanyName
     ' Decrements the reference count of loaded DLL module before leaving
       
    lstregister.RemoveItem (entry2)
    lstregister.Refresh
    FreeLibrary lLib
    'Resume Next
    Exit Sub
    
earlyExit2:
    Screen.MousePointer = vbDefault
    Status.Text = Status.Text & "Process failed in signaled state or time-out." & vbCrLf
    FreeLibrary lLib
     ' Terminate the thread to free up resources that are used by the thread
     ' NB Calling ExitThread for an application's primary thread will cause
     ' the application to terminate
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub

