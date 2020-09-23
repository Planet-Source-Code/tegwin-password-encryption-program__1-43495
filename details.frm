VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test "
   ClientHeight    =   2055
   ClientLeft      =   945
   ClientTop       =   5190
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   ScaleHeight     =   2055
   ScaleWidth      =   7335
   Begin VB.CommandButton Command1 
      Caption         =   "User Maintenance"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblLoginTime 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblUser 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Unload Me
frmAddNew.Show

End Sub


