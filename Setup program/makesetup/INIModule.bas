Attribute VB_Name = "INIModule"
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
 
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
                        
Function GetIni(ByVal INIFile As String, ByVal Section As String, ByVal Key As String)
 
    Dim RetVal As String, Worked As Integer
    RetVal = String$(255, 0)
    Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), INIFile)

    If Worked = 0 Then
        GetIni = ""
    Else
        GetIni = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
    End If
    
End Function

Function PutIni(ByVal INIFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
    WritePrivateProfileString Section, Key, Value, INIFile
End Function

