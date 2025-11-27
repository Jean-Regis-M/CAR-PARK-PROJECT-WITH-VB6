VERSION 5.00

Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows Default

   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End

   Begin VB.TextBox Text1 
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Text            =   "Car plate"
      Top             =   1920
      Width           =   2175
   End

   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   3
      Top             =   4920
      Width           =   2415
   End

   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   1080
      ScaleHeight     =   1155
      ScaleWidth      =   13515
      TabIndex        =   1
      Top             =   360
      Width           =   13575
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "WELCOME TO OUR ONLINE BOOKINGS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   9375
      End
   End

   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "Book a parking space"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   7395
      Left            =   120
      Picture         =   "dashboard.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15090
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Form4.Show
Me.Hide
End Sub

Private Sub Command6_Click()
Form6.Show
Me.Hide
End Sub
