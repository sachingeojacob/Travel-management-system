VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form salary_view 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22545
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   22545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   2160
      Top             =   9000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1720
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
      Connect         =   $"salary_view.frx":0000
      OLEDBString     =   $"salary_view.frx":00B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "salary_management"
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
      Height          =   975
      Left            =   8760
      TabIndex        =   1
      Top             =   9000
      Width           =   3495
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8115
      ScaleWidth      =   21675
      TabIndex        =   0
      Top             =   0
      Width           =   21735
   End
End
Attribute VB_Name = "salary_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub
