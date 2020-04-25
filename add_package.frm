VERSION 5.00
Begin VB.Form add_package 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "add_package.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   14040
      TabIndex        =   9
      Top             =   7080
      Width           =   4095
   End
   Begin VB.TextBox Text9 
      Height          =   735
      Left            =   14040
      TabIndex        =   8
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox Text8 
      Height          =   735
      Left            =   14040
      TabIndex        =   7
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   7920
      TabIndex        =   6
      Top             =   8640
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   7920
      TabIndex        =   5
      Top             =   7560
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   6480
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   7920
      TabIndex        =   3
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   7920
      TabIndex        =   2
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   7920
      TabIndex        =   1
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   7920
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   15960
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   12000
      Top             =   10320
      Width           =   2655
   End
End
Attribute VB_Name = "add_package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public cmd As New ADODB.Command
Public str As String
Public str1 As String

Public Sub connect()
con.Provider = "sqloledb"
str1 = "server=DESKTOP-HOTR91D\;database=master;trusted_connection=yes"
con.Open str1
End Sub
           



Private Sub Form_Load()
Dim n As Integer

'If con.State = adStateOpen Then
'rs.Close
'con.Close
'End If

Call connect
str1 = "select * from packages"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("package_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Image1_Click()
If Text2.Text = "" Then
MsgBox "Please Enter State !", vbExclamation, "Add packages"
Exit Sub
End If


If Text3.Text = "" Then
MsgBox "Please Enter a Spot !", vbExclamation, "Add packages"
Exit Sub
End If



If Text4.Text = "" Then
MsgBox "Please Enter Total distance !", vbExclamation, "Add packages"
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Please Enter car charges !", vbExclamation, "Add packages"
Exit Sub
End If



If Text6.Text = "" Then
MsgBox "Please Enter Bus Charges !", vbExclamation, "Add packages"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please Enter Stay Cost !", vbExclamation, "Add packages"
Exit Sub
End If


If Text8.Text = "" Then
MsgBox "Please Enter hotel1 !", vbExclamation, "Add packages"
Exit Sub
End If

If Text9.Text = "" Then
MsgBox "Please Enter hotel !", vbExclamation, "Add packages"
Exit Sub
End If

If Text10.Text = "" Then
MsgBox "Please Enter hotel3 !", vbExclamation, "Add packages"
Exit Sub
End If

Call connect

str1 = "insert into packages(package_id,state,spot,distance,car_charges,bus_charges,stay_cost,hotel1,hotel2,hotel3)values (" & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & Text4.Text & "," & Text5.Text & "," & Text6.Text & "," & Text7.Text & ",'" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "')"

con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Saved  package Details"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""


con.Close



Call connect
str1 = "select * from packages"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Dim n As Integer
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("package_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub
