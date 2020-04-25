VERSION 5.00
Begin VB.Form delete_package 
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21495
   LinkTopic       =   "Form1"
   Picture         =   "delete_package.frx":0000
   ScaleHeight     =   11070
   ScaleWidth      =   21495
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      Height          =   855
      Left            =   14640
      TabIndex        =   9
      Top             =   6600
      Width           =   4215
   End
   Begin VB.TextBox Text8 
      Height          =   855
      Left            =   14640
      TabIndex        =   8
      Top             =   5400
      Width           =   4215
   End
   Begin VB.TextBox Text7 
      Height          =   855
      Left            =   14640
      TabIndex        =   7
      Top             =   4200
      Width           =   4215
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   7920
      TabIndex        =   6
      Top             =   8760
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   7920
      TabIndex        =   5
      Top             =   7560
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   7920
      TabIndex        =   4
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   7920
      TabIndex        =   3
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   7920
      TabIndex        =   2
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7920
      TabIndex        =   1
      Top             =   3120
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Text            =   "SELECT PACKAGE ID"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   15960
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   12000
      Top             =   10320
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   11400
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "DELETE_PACKAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim str1 As String
Call connect
str1 = "select package_id from packages"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields("package_id").Value)
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Image1_Click()
Call connect
str1 = "select * from packages where package_id= '" & Combo1.Text & "' "
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Text = rs.Fields("state").Value
Text2.Text = rs.Fields("spot").Value
Text3.Text = rs.Fields("distance").Value
Text4.Text = rs.Fields("car_charges").Value
Text5.Text = rs.Fields("bus_charges").Value
Text6.Text = rs.Fields("stay_cost").Value
Text7.Text = rs.Fields("hotel1").Value
Text8.Text = rs.Fields("hotel2").Value
Text9.Text = rs.Fields("hotel3").Value

rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub Image2_Click()

If Combo1.Text = "SELECT PACKAGE ID" Then
MsgBox "Select Package ID to delete", vbExclamation, "Delete package"
Exit Sub
End If

Call connect

confirm = MsgBox("Are you sure you want to delete this package ?", vbYesNo, "Delete a Package")
If confirm = vbYes Then

str1 = "delete from packages where package_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "Package Deleted Successfully...", vbInformation, "Delete a Package"

DELETE_PACKAGE.Refresh

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

Else
MsgBox "Package not deleted ", vbInformation, "Delete a package"
End If
con.Close

End Sub

Private Sub Image3_Click()
Unload Me
End Sub

