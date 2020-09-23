VERSION 5.00
Begin VB.Form frmInstaller 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Installer"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelfExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdirectory 
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Text            =   "C:\Program Files\MyFolder"
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Install"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome. The Installer will guide you to install :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblversion 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxversion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   120
      Width           =   2055
      Visible         =   0   'False
   End
   Begin VB.Label lblcompany 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxcompany"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "label3"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   600
      Width           =   2895
      Visible         =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Install to directory:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Installer Made by Ronnie Staxborn Copyright (c) 2000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label lblExtract 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxextract"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "xxfiles"
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Visible         =   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   480
      Picture         =   "frmSelfExtract.frx":1272
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Info "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   5040
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   120
      Top             =   2040
      Width           =   6015
   End
End
Attribute VB_Name = "frmInstaller"
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

Private Sub cmdAbout_Click()

    MsgBox "Installer by Ronnie Staxborn." & vbCr & vbCr & "Thanx to Vasilis Sagonas & Chris Eastwood", vbInformation, "About..."

End Sub

Private Sub cmdCancel_Click()

If mFolderExists(GetTempPathName & "staxborn") = True Then
mEmptyFolder (GetTempPathName & "staxborn\")
End If

MsgBox "All tempfiles have now been removed from your computer. Please VOTE for me"
End
End Sub

Private Sub cmdExtract_Click()
mCheckPathFolders (txtdirectory.Text)
cmdExtract.Enabled = False
'SelfExtract
'mCopyFile GetTempPathName & "\staxborn" & "\Unzip32.dll", "C:\windows\system\Unzip32.dll", False
Call CopyFiles(CStr(GetTempPathName & "staxborn"), CStr("C:\windows\system"), CInt(0), CStr("*.dll"))
Call CopyFiles(CStr(GetTempPathName & "staxborn"), CStr("C:\windows\system"), CInt(0), CStr("*.ocx"))
Call CopyFiles(CStr(GetTempPathName & "staxborn"), CStr("C:\windows\system"), CInt(0), CStr("*.ole"))

Unzipsetup


On Error GoTo EH
Dim strProgramPath   As String   ' The path of the executable file
Dim strGroup         As String
Dim strProgramIconTitle As String
Dim strProgramArgs   As String
Dim sParent          As String

   strProgramPath = txtdirectory & "\" & GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "exename")
   strGroup = ".."
   strProgramIconTitle = GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "pname")
   strProgramArgs = ""
   
   sParent = "$(Programs)"
   
   CreateShellLink strProgramPath, strGroup, strProgramArgs, strProgramIconTitle, True, sParent
   Kill GetTempPathName & "staxborn" & "\setupinfo.ini"
   frmregister.Show
   
   
   Exit Sub
EH:
   MsgBox Err.Description
   Exit Sub


End Sub

Private Sub cmdView_Click()
frmFiles.Show 1
End Sub

Private Sub Form_Load()
On Error GoTo Err
lblExtract.Caption = ""

If mFolderExists(GetTempPathName & "staxborn") = False Then
mCreateFolder (GetTempPathName & "staxborn\")
End If


Dim x
Dim y
Dim z
Dim zz
x = GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "dir")
frmInstaller.txtdirectory.Text = x
y = GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "company")
frmInstaller.lblcompany.Caption = "Copyright (c) " & y
z = GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "version")
'frmInstaller.lblversion.Caption = "Version " & z
zz = GetIni(GetTempPathName & "\staxborn\setupinfo.ini", "data", "pname")
frmInstaller.lblWelcome.Caption = zz & " " & z


'Show
'Refresh
Exit Sub
Err:
MsgBox "This file is damaged or it doesn't include any files.", vbCritical, "Error"
End
End Sub
