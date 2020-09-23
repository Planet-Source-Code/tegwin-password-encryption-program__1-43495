VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Compass Car Hire Login Screen"
   ClientHeight    =   4020
   ClientLeft      =   3165
   ClientTop       =   2655
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4290
   Begin VB.CheckBox Check1 
      Caption         =   "Change Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3855
      Begin VB.Frame Frame3 
         Caption         =   "Password "
         Height          =   855
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   3015
         Begin VB.TextBox txtPassWord 
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "UserName"
         Height          =   855
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         Begin VB.TextBox txtLoginName 
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   2295
         End
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function LogUserIn() As Boolean

'Query the database and see if user exists...

Set DB = OpenDatabase(App.Path & "\contactor.mdb", False, False, ";pwd=jkddjkdd")
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin WHERE LoginName = '" & txtLoginName & "' AND PassWord = '" & txtPassWord & "'", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    MsgBox "Your login name or password is incorrect. If you can not log in, contact your system administrator and have them set up an account for you on the system.", vbInformation, "Login..."
    RS.Close
    DB.Close
    txtPassWord.SelStart = 0
    txtPassWord.SelLength = Len(txtPassWord)
    txtLoginName.SetFocus
    txtLoginName = ""
    txtPassWord = ""
    Exit Function
    frmLogin.Show
Else
    Me.Hide
    Form2.Show
End If

'Account Name...
If Not IsNull(RS!LoginName) Then
    Login.LoginName = RS!LoginName
End If
'Fullname...
If Not IsNull(RS!FullName) Then
    Login.FullName = RS!FullName
End If
'Is Admin...
Login.IsAdmin = RS!Administrator

'If RS!FirstLogin = "1" Then
'sgBox "Password needs changing@"
'frmChgPwd.Show

'End If

'Main Menu Panel...
Form2.lblLoginTime.Caption = Format$(Now, "H:MM AMPM DD/MM/YYYY")
Form2.lblUser.Caption = Login.FullName

If Login.IsAdmin = True Then
    Form2.Label1.Caption = "Administrator"
    Form2.Command1.Enabled = True
Else
    Form2.Label1.Caption = "User"
        Form2.Command1.Enabled = False
    
End If

Set RS = DB.OpenRecordset("tblSYSMSG")
RS.AddNew
RS!LoginName = txtLoginName
RS!FullName = Login.FullName
RS!DATEOFLogon = Format$(Now, "H:MM AMPM DD/MM/YYYY")
RS.Update






'Close the database...
RS.Close
DB.Close



LogUserIn = True
Exit Function







End Function


Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdLogin_Click()


If txtLoginName = "" Then
    MsgBox "You must login in order to continue. " & Chr$(10) + Chr$(13) & "If you do not have a login name " & Chr$(10) + Chr$(13) & "Please contact your Supervisor.", vbOKOnly + vbExclamation, "Login Required"
    Exit Sub
Else
    Call LogUserIn
End If




End Sub

