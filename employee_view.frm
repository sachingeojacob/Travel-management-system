VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form employee_view 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21525
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   21525
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   615
      Left            =   10560
      TabIndex        =   1
      Top             =   7800
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   6840
      Top             =   7440
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
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
      Connect         =   $"employee_view.frx":0000
      OLEDBString     =   $"employee_view.frx":00B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "employee_management"
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
   Begin VB.PictureBox DataGrid1 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   22755
      TabIndex        =   0
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "employee_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub command1_Click()
Unload Me
End Sub
