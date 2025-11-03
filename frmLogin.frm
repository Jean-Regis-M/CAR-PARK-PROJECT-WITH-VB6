VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAR PARKING SYSTEM"
   ClientHeight    =   5160
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3048.698
   ScaleMode       =   0  'User
   ScaleWidth      =   7056.179
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox admin 
      Height          =   345
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub admin_Change()

End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
If txtUserName.Text = "admin" Then
     If txtPassword.Text = "12345" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
         MsgBox "welcome!", , "Login"
      Form1.Show LoginSucceeded = True
        Me.Hide
    Else
    MsgBox "Invalid Password or User name, try again!", , "Login"
      LoginSucceeded = False
    txtPassword.Text = ""
    End If
    End If
End Sub



