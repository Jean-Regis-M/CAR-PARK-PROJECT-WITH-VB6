VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00808080&
   Caption         =   "Form6"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15870
   LinkTopic       =   "Form6"
   ScaleHeight     =   8490
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "booking.frx":0000
      Height          =   3015
      Left            =   8160
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8040
      Top             =   6960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"booking.frx":0015
      OLEDBString     =   $"booking.frx":009D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Payment"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bookings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton Command6 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   23
         Top             =   7080
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   22
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   20
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   19
         Top             =   7080
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ADD NEW"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   18
         Top             =   7080
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "Ticket Number"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   5280
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PAY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "Time spent"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         DataField       =   "Car Plate Number"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         DataField       =   "First Name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         DataField       =   "Phone Number"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Text            =   "(000) 000-000"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         DataField       =   "Last Name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Car Model"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2040
         TabIndex        =   2
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "Parking Preference"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2040
         TabIndex        =   1
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Ticket Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "First name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Last name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Phone Number*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Car Plate Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Car Model"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Parking Preference"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Hours you'll spend"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, tot As Integer
a = Val(Text2.Text)
tot = a * 1000
MsgBox "You'll pay" & tot
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Form1.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
Combo1.AddItem "TOYOTA"
Combo1.AddItem "AUDI"
Combo1.AddItem "MERCEDES BENZ"
Combo1.AddItem "FORD"
Combo1.AddItem "HYUNDAI"
Combo1.AddItem "VOLKSWAGEN"
Combo1.AddItem "CHEVROLET"
Combo1.AddItem "JEEP"
Combo1.AddItem "ROLLS-ROYCE"
Combo2.AddItem "ANGULAR"
Combo2.AddItem "PARALLEL"
Combo2.AddItem "PERPENDICULAR"
End Sub

