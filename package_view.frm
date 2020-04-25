VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form package_view 
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21495
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   21495
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1455
      Left            =   1920
      Top             =   8640
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
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
      Connect         =   $"package_view.frx":0000
      OLEDBString     =   $"package_view.frx":00B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "packages"
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
      Height          =   1095
      Left            =   7560
      TabIndex        =   1
      Top             =   8520
      Width           =   3495
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   21435
      TabIndex        =   0
      Top             =   0
      Width           =   21495
   End
End
Attribute VB_Name = "package_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub
