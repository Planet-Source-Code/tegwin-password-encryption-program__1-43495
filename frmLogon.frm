VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compass Car Hire Login Screen"
   ClientHeight    =   7050
   ClientLeft      =   2490
   ClientTop       =   555
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   4440
   Begin VB.CheckBox chkDoNotDisplay 
      Caption         =   "I accept the above conditions, do not show again. "
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Frame fraDisclaimer 
      Caption         =   "Disclaimer"
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   3855
      Begin VB.TextBox txtDisclaimer 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
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
      Begin VB.CheckBox chkRememberMe 
         Caption         =   "Remember Me"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
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
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "By Pressing OK, You are agreeing to, and understand  the terms and conditions in the disclaimer. "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Please Press Cancel if you do not agree. You must cease the use of this software if you do not agree to the disclaimer above. "
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   4215
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function UserLogin() As Boolean

'Query the database and see if user exists...

Set DB = OpenDatabase(App.Path & "\password.mdb", False, False, ";pwd=passw0rd")
Set RS = DB.OpenRecordset("SELECT * FROM Login WHERE LoginName = '" & txtLoginName & "' AND PassWord = '" & crypt(txtPassWord.Text, txtPassWord.Text) & "'", dbOpenSnapshot)



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
   frmTest.Show
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
'Main Menu Panel...



frmTest.lblLoginTime.Caption = Format$(Now, "H:MM AMPM DD/MM/YYYY")
frmTest.lblUser.Caption = Login.FullName



If Login.IsAdmin = True Then
    frmTest.Label1.Caption = "Administrator"
    frmTest.Command1.Enabled = True
Else
    frmTest.Label1.Caption = "User"
        frmTest.Command1.Enabled = False
    
End If

'Close the database...
RS.Close
DB.Close

UserLogin = True
Exit Function

End Function



Private Sub chkDoNotDisplay_Click()


Dim noDisclaimer As String


    If chkDoNotDisplay.Value = Checked Then
        fraDisclaimer.Visible = False
        frmLogon.Height = 4005
    Else
        fraDisclaimer.Visible = True
    End If
    
  'Set noDisclaimer to No, this will not show it again
  
  noDisclaimer = "No"
    
  'If this button was checked then the user does not want to display the
  'Disclaimer make the field no
  
  doini = WritePrivateProfileString("Disclaimer", "Show Disclaimer", noDisclaimer, App.Path & "\compass.ini")
  
    

End Sub

Private Sub chkRememberMe_Click()
Dim RememberMe As String

   
  'Set RememberMe to Yes
  
  RememberMe = "Yes"
    
  'Write Settings into Ini File
  doini = WritePrivateProfileString("RememberMe", "Remember", RememberMe, App.Path & "\compass.ini")
  
End Sub

Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdLogin_Click()


If txtLoginName = "" Then
    MsgBox "You must login in order to continue. " & Chr$(10) + Chr$(13) & "If you do not have a login name " & Chr$(10) + Chr$(13) & "Please contact your Supervisor.", vbOKOnly + vbExclamation, "Login Required"
    Exit Sub
Else
    Call UserLogin
End If




End Sub

Private Sub Form_Load()

'Set the string for the disclaimer
        Disclaimer = "To the maximum extent permitted by applicable law, in no event shall the company or it's representatives be liable for any special, incidental, indirect, or consequential damages whatsoever, (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or any other loss) arising out of the use of or inability to use this SOFTWARE PRODUCT or the provision or failure to provide support services, even if the company has been informed of  the possibility of such damages."


'Put the disclaimer string into the text box

txtDisclaimer.Text = Disclaimer

'Check the INI File to see if the disclaimer must be set to on

doini = GetPrivateProfileString("Disclaimer", " Show Disclaimer", "Yes", noDisclaimer, 5, App.Path & "\compass.ini")

'Check to see if The password should be remembered
doini = GetPrivateProfileString("RemeberMe", "Remember", "Yes", Remember, 5, App.Path & "\compass.ini")


If Left(noDisclaimer, 2) = "No" Then
    Me.chkDoNotDisplay.Value = Checked
End If


    If chkDoNotDisplay.Value = Checked Then
        Me.fraDisclaimer.Visible = False
        frmLogon.Height = 4005
    Else
        Me.fraDisclaimer.Visible = True
    End If
End Sub
