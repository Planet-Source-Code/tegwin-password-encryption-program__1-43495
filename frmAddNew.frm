VERSION 5.00
Begin VB.Form frmAddNew 
   Caption         =   "User Maintainance"
   ClientHeight    =   5835
   ClientLeft      =   4065
   ClientTop       =   2190
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete User"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit User"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display Users"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
      Begin VB.ListBox lstUsers 
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      Begin VB.CheckBox Check2 
         Caption         =   "Change at Next LOGON"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Administrator"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Full Name"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UpdateList()

cmdAddNew.Enabled = False


Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
   
End If

'Load all users...
lstUsers.Clear
Do
    lstUsers.AddItem RS("LoginName")
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close

End Sub
Private Sub Command1_Click()





Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
   
End If

'Load all users...
lstUsers.Clear
Do
    lstUsers.AddItem RS("LoginName")
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close






End Sub

Private Sub cmdAddNew_Click()

'If Text1.Text = "" Then
 '   MsgBox "Cannot Have blank user name ", vbOKOnly + vbCritical, "Error"
  '  Exit Sub
'End If




'Open the database and table
Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("Login")

' Check if the user name exists
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & Text1.Text & "'", dbOpenSnapshot)

'If the Username exists then Exit the Sub
If RS.RecordCount > 0 Then
    MsgBox "This user name already exists. Please select a unique user name.", vbInformation, "Account Already Exists..."
   frmAddNew.Text1.Text = ""
   frmAddNew.Text2.Text = ""
   frmAddNew.Text3.Text = ""
   frmAddNew.Check1.Value = Unchecked
   frmAddNew.cmdAddNew.Enabled = False
   
   Exit Sub
End If

' The database needs to be reopened here in order for
' the information to added into it.

Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("Login")
RS.MoveFirst
RS.AddNew
RS!LoginName = frmAddNew.Text1.Text
RS!Password = crypt(frmAddNew.Text2.Text, frmAddNew.Text2.Text)
 
'Crypt it with itself so there is no easy way of de-crypting the pass

TheKey = frmAddNew.Text2.Text

RS!FullName = frmAddNew.Text3.Text
RS!Administrator = Check1.Value
RS.Update

lstUsers.Clear
Do
    lstUsers.AddItem RS("LoginName")
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close



' Once all the user details have been entered
' pop up a message saying he has been added

MsgBox "The User " & Text3.Text & " has been added to the system", vbOKOnly + vbInformation, "New User Added"


' CLear all the values

frmAddNew.Text1.Text = ""
frmAddNew.Text2.Text = ""
frmAddNew.Text3.Text = ""
frmAddNew.Check1.Value = Unchecked

cmdAddNew.Enabled = False






End Sub

Private Sub cmdDelete_Click()


' Check to see if a users name has been clicked
' before deleting

If lstUsers.ListIndex < 0 Then
    MsgBox "Please select a user from a list at the left before you continue.", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If




Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & lstUsers & "'", dbOpenDynaset)

'No info found for contact ???...

    If RS.RecordCount > 0 Then
    RS.Delete
    RS.Close
    DB.Close
   
 End If



'Call the routine to update the list

UpdateList




MsgBox "The Details for " & Text3.Text & " have been deleted ", vbOKOnly + vbInformation, "Details Deleted"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = Unchecked
End Sub

Private Sub cmdEdit_Click()
Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & lstUsers & "'", dbOpenDynaset)

If lstUsers.ListIndex < 0 Then
    MsgBox "Please select a user from a list at the left before you continue.", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If


Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & lstUsers & "'", dbOpenDynaset)

'No info found for contact ???...
If RS.RecordCount > 0 Then
    RS.Edit
Else
    'rs.AddNew
End If

'Update the system accounts...
RS!LoginName = Text1.Text
RS!FullName = Text3.Text
RS!Password = crypt(frmAddNew.Text2.Text, frmAddNew.Text2.Text)
RS!Administrator = Check1.Value

 
RS.Update
RS.Close
DB.Close

' Calls the form load to update the list

UpdateList


MsgBox "The Details for " & Text3.Text & " have been changed", vbOKOnly + vbInformation, "Details Changed"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = Unchecked
cmdAddNew.Enabled = False


End Sub

Private Sub cmdExit_Click()

Unload Me
'MDIFrmMain.Show


End Sub

Private Sub Form_Load()

cmdAddNew.Enabled = False


Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
   
End If

'Load all users...
lstUsers.Clear
Do
    lstUsers.AddItem RS("LoginName")
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close

End Sub

Private Sub lstUsers_Click()
'Query the database and see if user exists...

Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & lstUsers & "'", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Sub
End If

'Populate fields...
Text1.Text = RS!LoginName
Text3.Text = RS!FullName
Text2.Text = RS!Password
'Check1.Value = RS!Administrator

If RS!Administrator = True Then
    Check1.Value = 1
Else
   Check1.Value = 0
End If

If lstUsers.List(lstUsers.ListIndex) = "admin" Then
    Check1.Enabled = False
    cmdDelete.Enabled = False
    Text1.Enabled = False
    
Else
    Check1.Enabled = True
    cmdDelete.Enabled = True
    Text1.Enabled = True
    
    
End If





RS.Close
DB.Close
End Sub

Private Sub Text1_Change()


cmdAddNew.Enabled = True

End Sub

Private Sub Text1_LostFocus()

'Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
'Set RS = DB.OpenRecordset("Login")


'Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & Text1.Text & "'", dbOpenSnapshot)

'Create a new account...
'If RS.RecordCount > 0 Then
    'MsgBox "This user name already exists. Please select a unique user name.", vbInformation, "Account Already Exists..."
    'Exit Sub
 '   NameExists = 1
  '  MsgBox NameExists
'End If


End Sub
