Attribute VB_Name = "ReadINIFile"
Option Explicit

' Retrieves a string from the specified section in an initialization file.
' See here: https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getprivateprofilestring?redirectedfrom=MSDN
' lpApplicationName     The name of the section containing the key name.
' lpKeyName             The name of the key whose associated string is to be retrieved.
' lpDefault             A default string.
' lpReturnedString      A pointer to the buffer that receives the retrieved string.
' nSize                 The size of the buffer pointed to by the lpReturnedString parameter, in characters.
' lpFileName            The name of the initialization file.

Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

' Given a key it returns the associated value from an INI file.
' iniFileName should specify a valid path.
' jesus.aneiros@gmail.com
' 2020.06.11
Public Function readIniFileString(ByVal iniFileName As String, ByVal sectionName As String, ByVal keyName As String) As String

    Dim lResult As Long
    Dim retString As String * 255
    Dim retStringSize As Long

    ' The buffer
    retString = Space(255)
        
    ' Returns the number of caracters copied to the buffer retString
    lResult = GetPrivateProfileString(sectionName, keyName, "", retString, Len(retString), iniFileName)
        
    If (lResult) Then
        readIniFileString = Left$(retString, lResult)
    Else
        readIniFileString = ""
    End If

End Function

Public Sub test()
    Debug.Print readIniFileString(ThisWorkbook.Path & "\model_c.ini", "DB", "host")
    Debug.Print readIniFileString(ThisWorkbook.Path & "\model_c.ini", "EMAIL", "host")
End Sub
