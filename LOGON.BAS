Attribute VB_Name = "Module1"

Global DB As Database
Global RS As Recordset
Global SQL As String
Global Login As tLogin
Global QuickRef As tQuickRef
Global NameExists As Integer


'Quick Reference Array...
Type tQuickRef
    CallingForm As String
    ContactName As String
    ContactID As Long
    DBFileName As String
    DBPassWord As String
    DBTimeOut As Long
    INIFileName As String
    LargeMenuHeight As Long
    LargeMenuWidth As Long
    MediumMenuHeight As Long
    MediumMenuWidth As Long
    NotesHaveChanged As Boolean
    PassNotes As Boolean
    ReLoggingIn As Boolean
    UpdateColors As Boolean
    UpdateInternetSites As Boolean
    UpdateNotes As Boolean
End Type

'Login...
Type tLogin
    FullName As String
    IsAdmin As Boolean
    LoginDateTime As String
    LoginName As String
End Type







Sub Main()

On Local Error Resume Next

'DB Password...
QuickRef.DBPassWord = "jkddjkdd"

'Window Heights and Widths...
QuickRef.MediumMenuHeight = 4320
QuickRef.MediumMenuWidth = 7020
QuickRef.LargeMenuHeight = 5705
QuickRef.LargeMenuWidth = 9665

    QuickRef.DBFileName = "c:test\Contactor.Mdb"
QuickRef.DBTimeOut = 500


frmLogin.Show


End Sub
