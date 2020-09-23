Attribute VB_Name = "General1"
Global Deck As String



Public gListViewTotalSelected As Long
Public gListViewSelected() As Long
Public gListViewItemToInsertBefore As Long

Const S_OK = &H0                ' Success
Const S_FALSE = &H1             ' The Folder is valid, but does not exist
Const E_INVALIDARG = &H80070057 ' Invalid CSIDL Value

Const CSIDL_LOCAL_APPDATA = &H1C&
Const CSIDL_FLAG_CREATE = &H8000&

Const SHGFP_TYPE_CURRENT = 0
Const SHGFP_TYPE_DEFAULT = 1
Const MAX_PATH = 260

Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Function SHGetFolderLocation Lib "shell32" (hWnd As Long, nFolder As Long, hToken As Long, dwReserved As Long, ppidl As Long) As Long
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
Declare Function GetActiveWindow Lib "User32" () As Integer


Enum Folders
 Desktop = &H0
 Internet = &H1
 Programs = &H2
 ControlsFolder = &H3
 Printers = &H4
 Personal = &H5
 Favorites = &H6
 StartUp = &H7
 Recent = &H8
 SendTo = &H9
 RecycleBin = &HA
 StartMenu = &HB
 DesktopDirectory = &H10
 Drives = &H11
 Network = &H12
 Nethood = &H13
 Fonts = &H14
 Templates = &H15
 Common_StartMenu = &H16
 Common_Programs = &H17
 Common_StartUp = &H18
 Common_DesktopDirectory = &H19
 ApplicationData = &H1A
 PrintHood = &H1B
 AltStartUp = &H1D
 Common_AltStartUp = &H1E
 Common_Favorites = &H1F
 InternetCache = &H20
 Cookies = &H21
 History = &H22
End Enum

Enum BrowseForFolderFlags
    ReturnFileSystemFoldersOnly = &H1
    DontGoBelowDomain = &H2
    IncludeStatusText = &H4
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    IncludeTextBox = &H10
    ReturnFileSystemAncestors = &H8
End Enum

Type BrowseInfo
     hwndOwner As Long
     pidlroot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type





Dim pidlroot As Long

Sub ClearId()

End Sub


Function CheckFolderID(Folder As Folders) As String
Dim sPath As String
Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH, 0)

RetVal = SHGetFolderPath(0, Folder Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)

Select Case RetVal
    Case S_OK
        ' We retrieved the folder successfully
        
        ' All C strings are null terminated
        ' So we need to return the string upto the first null character
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        CheckFolderID = sPath
    Case S_FALSE
        ' The CSIDL in nFolder is valid, but the folder does not exist.
        ' Use CSIDL_FLAG_CREATE to have it created automatically
        CheckFolderID = ""
    Case E_INVALIDARG
        ' nFolder is invalid
        CheckFolderID = ""
End Select
End Function

Public Function BrowseForFolder(hWnd As Long, Optional Title As String, Optional flags As BrowseForFolderFlags, Optional StartUpSpecialFolder As Folders) As String

    'Variables for use:
     Dim iNull As Integer
     Dim IDList As Long
     Dim Result As Long
     Dim Path As String
     Dim bi As BrowseInfo
     Dim ret As String
     If flags = 0 Then flags = BIF_RETURNONLYFSDIRS
     
    'Type Settings
     With bi
        ret = CheckFolderID(StartUpSpecialFolder) 'Check if the special folder exists
        If ret <> "" Then .pidlroot = StartUpSpecialFolder 'If there is any valid ID use it
        .hwndOwner = hwndOwner 'Set Owner Window
        .ulFlags = flags 'Set flags (if any)
        .lpszTitle = lstrcat(Title, Chr(0)) 'Append title string to a long value
     End With

    'Execute the BrowseForFolder shell API and display the dialog
     IDList = SHBrowseForFolder(bi) 'Return ID List (selected path in a long value)
     
    'Get the info out of the dialog
     If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path) 'Convert ID list to a string
        iNull = InStr(Path, vbNullChar) 'Get the position of the null character
        If iNull Then Path = Left$(Path, iNull - 1) 'Remove the null character
     End If

    'If Cancel button was clicked, error occured or Non File System Folder was selected then Path = ""
     BrowseForFolder = Path
End Function

