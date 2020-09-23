VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMakeInstaller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Installer"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmMakeSelExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Top             =   5760
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.Frame Frame5 
      Caption         =   "Done"
      Height          =   3735
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   5655
      Visible         =   0   'False
      Begin VB.Label Label10 
         Caption         =   "The SETUP is now done."
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "If anyone want something more in this Installer, please let me know.                      rompa@hem.passagen.se"
         Height          =   735
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Welcome"
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5655
      Begin VB.Label Label7 
         Caption         =   $"frmMakeSelExtract.frx":0442
         Height          =   735
         Left            =   360
         TabIndex        =   49
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakeSelExtract.frx":04DB
         Height          =   615
         Left            =   360
         TabIndex        =   32
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   "Version 1.3"
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   $"frmMakeSelExtract.frx":0599
         Height          =   1095
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "Welcome to this simple but yet effective setup program. "
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Make the Installer Info"
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   5655
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   15
         Top             =   1680
         Width           =   315
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "The template in this program is ""Installer.exe"" (you can configure this yourself int the folder Install)."
         Height          =   735
         Left            =   240
         TabIndex        =   35
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3960
         Picture         =   "frmMakeSelExtract.frx":06AF
         Top             =   2400
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose where the new file will be writen."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose the Installer Template"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5955
      TabIndex        =   36
      Top             =   0
      Width           =   6015
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5040
         Picture         =   "frmMakeSelExtract.frx":1921
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "You will now make your setup file."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Make Installer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "< Back"
      Height          =   495
      Left            =   840
      TabIndex        =   34
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   495
      Left            =   1680
      TabIndex        =   33
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtWelcome 
      Height          =   285
      Left            =   -120
      TabIndex        =   28
      Top             =   5520
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Picture         =   "frmMakeSelExtract.frx":2B93
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      Picture         =   "frmMakeSelExtract.frx":2FD5
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add your files/dependencies"
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   5655
      Begin VB.CommandButton Command9 
         Caption         =   "Create shortcut"
         Height          =   375
         Left            =   3720
         TabIndex        =   47
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add"
         Height          =   435
         Left            =   3720
         TabIndex        =   42
         ToolTipText     =   "Add dependecies file"
         Top             =   3120
         Width           =   795
      End
      Begin VB.ListBox lstdep 
         Height          =   1230
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add dir"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   960
         Width           =   795
         Visible         =   0   'False
      End
      Begin VB.ListBox lstFiles 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         ItemData        =   "frmMakeSelExtract.frx":3417
         Left            =   120
         List            =   "frmMakeSelExtract.frx":3419
         TabIndex        =   21
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3720
         TabIndex        =   20
         ToolTipText     =   "Add your program files"
         Top             =   1440
         Width           =   795
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4560
         TabIndex        =   19
         ToolTipText     =   "Remove selected file"
         Top             =   1440
         Width           =   795
      End
      Begin VB.ListBox lstzip 
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   2040
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label Label18 
         Caption         =   "Mark the selected file to make a shortcut to."
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Dependencies:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Program files:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to choose what files are going to be included at the Installer."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Program Information"
      Height          =   3735
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox txtexename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   46
         Top             =   3240
         Width           =   4335
      End
      Begin VB.TextBox txtProgramName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtcompany 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtVersion 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtAbout 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "C:\windows\desktop\"
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label label17 
         Caption         =   "The EXE name of the program (e.g. myapp.exe)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label lblProgramname 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Company Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Program path:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMakeInstaller"
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

'Option Explicit
Dim i
'Dim izip
Dim iName
Dim rTMP
Dim rTMP0
Dim p As Integer
Dim X
Private WithEvents m_oWiz As cWizardEngine
Attribute m_oWiz.VB_VarHelpID = -1
Function OnlyFileName(file) As String
If InStr(file, "\") = 0 Then OnlyFileName = file: Exit Function
rTMP = 1
Do
    rTMP0 = rTMP
    rTMP = InStr(rTMP + 1, file, "\")
Loop Until rTMP = 0
OnlyFileName = Right(file, Len(file) - Len(Left(file, rTMP0)))
End Function

Private Sub Command1_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
Text1 = CommonDialog1.FileName
UsrCancel:
End Sub

Private Sub Command2_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "All Files|*.*|"
CommonDialog1.flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
For i = 0 To lstFiles.ListCount - 1
    If LCase$(OnlyFileName(CommonDialog1.FileName)) = LCase$(OnlyFileName(lstFiles.List(i))) Then MsgBox "A file with the same name exists!", vbExclamation, "Oops!": Exit Sub
Next i
lstFiles.AddItem CommonDialog1.FileName

p = InStr(1, CommonDialog1.FileName, ".", vbTextCompare)
X = Mid(CommonDialog1.FileName, p + 1, Len(CommonDialog1.FileName))
If X = "exe" Then
txtexename.Text = OnlyFileName(CommonDialog1.FileName)
Else
'txtexename.Text = ""
End If

UsrCancel:

End Sub

Private Sub Command3_Click()
'check something first...
Dim ind
Dim entry
If Text1.Text = "" Then MsgBox "You must choose the X-tractor!", vbExclamation, "Oops!": Text1.SetFocus: Exit Sub
If Text3.Text = "" Then MsgBox "You must choose the output filename!", vbExclamation, "Oops!": Text3.SetFocus: Exit Sub
If lstFiles.ListCount = 0 Then MsgBox "You must add files!", vbExclamation, "Oops!": Exit Sub
If txtProgramName.Text = "" Then MsgBox "You must enter your Program Name", vbExclamation, "Oops!": txtProgramName.SetFocus: Exit Sub
If txtVersion.Text = "" Then MsgBox "You must enter the Version of your program", vbExclamation, "Oops!": txtVersion.SetFocus: Exit Sub
'if everything is ok continue...
 
 lstzip.Clear
    
    For ind = 0 To lstdep.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = lstdep.List(ind)
            lstzip.AddItem entry
            Next
            
'txtWelcome.Text = txtProgramName.Text & "  " & " v" & txtVersion.Text & vbCr & "Copyright (c) " & txtcompany.Text & " 2000"

Call PutIni(App.Path & "\setupinfo.ini", "data", "dir", txtAbout.Text)
Call PutIni(App.Path & "\setupinfo.ini", "data", "company", txtcompany.Text)
Call PutIni(App.Path & "\setupinfo.ini", "data", "version", txtVersion.Text)
Call PutIni(App.Path & "\setupinfo.ini", "data", "pname", txtProgramName.Text)
Call PutIni(App.Path & "\setupinfo.ini", "data", "exename", txtexename.Text)

Dim oZip As CGZipFiles

On Error GoTo vbErrorHandler

   
    Set oZip = New CGZipFiles
    
    With oZip
'
' Give Zip File a Name / Path
'
        .ZipFileName = App.Path & "\DummyIns.ZIP"
'
' Are we updating a Zip File ?
' - This doesn't seem to work - check InfoZip
' homepage for more info.
'
        .UpdatingZip = False ' ensures a new zip is created
'
' Add in the files to the zip - in this case, we
' want all the ones in the current directory

        For izip = 0 To lstFiles.ListCount - 1
        iName = frmMakeInstaller.OnlyFileName(lstFiles.List(izip))
        .AddFile frmMakeInstaller.OnlyFileName(lstFiles.List(izip))
        Next izip
'
' Make the zip file & display any errors
'
        If .MakeZipFile <> 0 Then
            MsgBox .GetLastMessage ' any errors
        End If
    End With
    
    Set oZip = Nothing
       
    lstzip.AddItem App.Path & "\setupinfo.ini"
    lstzip.AddItem App.Path & "\Unzip32.dll"
    lstzip.AddItem App.Path & "\Vb6stkit.dll"
    lstzip.AddItem App.Path & "\DummyIns.ZIP"
        
    If AddToSelfExtract(Text1, frmMakeInstaller.lstzip, Text3) = True Then
    Kill App.Path & "\DummyIns.ZIP"
    Kill App.Path & "\setupinfo.ini"
    Me.Frame5.Visible = True
    Me.Caption = "Make Installer"
    MsgBox "Done!", vbInformation, "Done!"
    End If
    
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Make Installer::Zipping..." & " " & Err.Description



End Sub

Private Sub Command4_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub
Text3 = CommonDialog1.FileName
UsrCancel:

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
On Error Resume Next
lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub Command7_Click()

On Error GoTo Er:

    Dim ReturnValue As String 'Keeps up the return
    Dim WithFiles As Long 'Just for this project, to add browsing with files or not
    Dim SelectedFolder
   
    ReturnValue = BrowseForFolder(Me.hWnd, "Choose a directory to add:", WithFiles, RecycleBin)
    If ReturnValue <> "" Then
      SelectedFolder = ReturnValue: GoTo 123
    Else
Exit Sub
    End If
    
123:
    
    If Right(SelectedFolder, 1) = "\" Then
    File1.Path = SelectedFolder

    Else
    File1.Path = SelectedFolder & "\"
    End If
    
File1.ListIndex = -1

Dim i

For i = 1 To File1.ListCount

File1.ListIndex = File1.ListIndex + 1
Dim mother
mother = File1.Path & "\" & File1.FileName

lstFiles.AddItem mother

DoEvents
Next i
Exit Sub
Er:
MsgBox Err.Description, vbCritical, "File/Path Error"
End Sub

Private Sub Command8_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "All Files|*.*|"
CommonDialog1.flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
For i = 0 To lstdep.ListCount - 1
    If LCase$(OnlyFileName(CommonDialog1.FileName)) = LCase$(OnlyFileName(lstdep.List(i))) Then MsgBox "A file with the same name exists!", vbExclamation, "Oops!": Exit Sub
Next i
lstdep.AddItem CommonDialog1.FileName
UsrCancel:
End Sub

Private Sub Command9_Click()
Dim short
For short = 0 To lstFiles.ListCount - 1
txtexename.Text = OnlyFileName(lstFiles.List(short))
Next
End Sub

Private Sub Form_Load()
'Show
'Text1.SetFocus
'Text1.Text = App.Path & "Intaller.exe"
 Set m_oWiz = New cWizardEngine

 '-- Add the panels in the order we wan
  '   t them displayed.
 m_oWiz.AddPanel Me.Frame1
 m_oWiz.AddPanel Me.Frame2
 m_oWiz.AddPanel Me.Frame3
 m_oWiz.AddPanel Me.Frame4
 'm_oWiz.AddPanel Me.Frame5

' '-- Add the buttons.
 Set m_oWiz.CancelButton = Me.Command5
'
 Set m_oWiz.FinishButton = Me.Command3
'
 Set m_oWiz.NextButton = Me.cmdNext
 Set m_oWiz.PrevButton = Me.cmdPrev
'
' '-- Only allow the finish button on th
'     e last panel.
 m_oWiz.FinishEnabledOnAllPanels = False
'
' '-- Start the wizard.
 m_oWiz.StartWizard

Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub txtAbout_Change()
If Len(txtAbout.Text) > 256 Then txtAbout.Text = Left(txtAbout.Text, 256)
End Sub


