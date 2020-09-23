VERSION 5.00
Begin VB.Form frmPreinstall 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pre Install"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
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
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a preinstall. When you cancel all temp files will be removed from you computer."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblfiles 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait until the Installer popups."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The Installer will now extract your files for easy installing. "
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Top             =   1560
      Width           =   4815
   End
End
Attribute VB_Name = "frmPreinstall"
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

Private Sub Form_Load()
On Error GoTo Err
Dim rWelcome As String
Dim rAbout As String
Dim iFiles As String
Dim iName As String
Dim Size As String

' see if the folder "staxborn" exists, if not
' it will be created if it does exsist it will
' empty the folder.
If mFolderExists(GetTempPathName & "staxborn") = False Then
mCreateFolder (GetTempPathName & "staxborn\")
Else
mEmptyFolder (GetTempPathName & "staxborn\")
End If



iFreeFile = FreeFile
curPOS = 0
i = 0
Open rPath & App.EXEName & ".exe" For Input As iFreeFile
Close iFreeFile
iFreeFile = FreeFile
Open rPath & App.EXEName & ".exe" For Binary As iFreeFile
    
    Seek #iFreeFile, LOF(iFreeFile) - 6 - (256 * 2)
    iFiles = String(5, Chr(0))
    Get iFreeFile, , iFiles

    iFiles = Replace$(iFiles, vbCr, "")
    iFilez = Val(iFiles)
    'lblfiles.Caption = "This file contains " & iFilez & " files."
    
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2)
    rWelcome = String(256, Chr(0))
    Get iFreeFile, , rWelcome
    
    rWelcome = Replace(rWelcome, vbTab, "")
   ' If rWelcome <> "" Then lblWelcome.Caption = rWelcome

    Seek #iFreeFile, LOF(iFreeFile) - 256
    rAbout = String(256, Chr(0))
    Get iFreeFile, , rAbout

    rAbout = Replace(rAbout, vbTab, "")
    'If rAbout <> "" Then txtdirectory.Text = rAbout
    
Close iFreeFile
Do
i = i + 1
    Open rPath & App.EXEName & ".exe" For Binary As iFreeFile

    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 41 - 10 + curPOS
    iName = String(40, Chr(0))
    Get iFreeFile, , iName
    
    
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 11 + curPOS
    Size = String(10, Chr(0))
    Get iFreeFile, , Size
        
    Size = CCur(Size)
    
    Close iFreeFile
    FFile = FreeFile
    iName = Replace$(iName, vbCr, "")
    
    curPOS = curPOS - Size - 50

Loop Until i >= iFilez

Show
'Refresh

SelfExtract

timedPause (2)
frmInstaller.Show
Unload frmPreinstall

Exit Sub
Err:
MsgBox "This file is damaged or it doesn't include any files.", vbCritical, "Error"
End


End Sub
