Attribute VB_Name = "FileHelper"
Option Explicit

' Returns a string of the first filename that matches a pattern inside
' a directory. If no file matches the pattern then it returns an empty string.
' Arguments are optional, default is current directory and *.* pattern.
' File pattern could use standard Windows wildcards * and ?
' Uses VBA function Dir
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dir-function
Public Function FindFile(Optional pathName As String = ".", _
    Optional filePattern As String = "*.*") As String
    
    Dim fullPath As String
    
    fullPath = pathName & "\" & filePattern
            
    FindFile = Dir(fullPath, vbNormal)
            
End Function

Public Sub testFindFile()
    Debug.Print FindFile(".", "*.xlsx")
    Debug.Print FindFile()
    Debug.Print FindFile(pathName:="c:\users\excel\Downloads")
    Debug.Print FindFile(filePattern:="*.txt")
    Debug.Print FindFile(filePattern:="FileHelper.*")
End Sub

