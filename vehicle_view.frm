VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form vehicle_view 
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22230
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   22230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   1800
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   $"vehicle_view.frx":0000
      OLEDBString     =   $"vehicle_view.frx":00B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "vehicle"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   7920
      Width           =   3495
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6795
      ScaleWidth      =   22155
      TabIndex        =   0
      Top             =   0
      Width           =   22215
   End
End
Attribute VB_Name = "vehicle_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub
