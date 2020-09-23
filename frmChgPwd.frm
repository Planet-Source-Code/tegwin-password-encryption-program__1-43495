VERSION 5.00
Begin VB.Form frmChgPwd 
   Caption         =   "Change Password "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmChgPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set DB = OpenDatabase("C:\test\contactor.mdb", False, False, ";pwd=jkddjkdd")
Set RS = DB.OpenRecordset("tblLogin")
RS.MoveFirst
RS.Edit
'RS!LoginName = Form1.Text1.Text
RS!Password = frmChgPwd.Text1.Text
'RS!FullName = Form1.Text3.Text
'RS!Administrator = Check1.Value
'RS.Update

End Sub
