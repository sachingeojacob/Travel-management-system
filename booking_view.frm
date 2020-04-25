VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form booking_view 
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox DataGrid1 
      Height          =   6615
      Left            =   1080
      ScaleHeight     =   6555
      ScaleWidth      =   17115
      TabIndex        =   1
      Top             =   1320
      Width           =   17175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   840
      Top             =   8520
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1931
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
      Connect         =   $"booking_view.frx":0000
      OLEDBString     =   $"booking_view.frx":00AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "booking"
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
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   735
      Left            =   7440
      TabIndex        =   0
      Top             =   8400
      Width           =   2775
   End
End
Attribute VB_Name = "booking_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub

