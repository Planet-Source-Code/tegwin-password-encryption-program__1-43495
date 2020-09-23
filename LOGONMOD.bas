Attribute VB_Name = "LOGINMOD"

Global DB As Database
Global RS As Recordset
Global Login As tLogin

'Settings for the INI file
Global Disclaimer As String
Global doini As String * 5
Global noDisclaimer As String * 5
Global Remember As String * 5
Global User As String


'Login...
Type tLogin
    FullName As String
    IsAdmin As Boolean
    LoginDateTime As String
    LoginName As String
End Type


'Declare statements for the ini file

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


