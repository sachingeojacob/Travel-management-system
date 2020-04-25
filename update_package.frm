VERSION 5.00
Begin VB.Form update_package 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "update_package.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      Height          =   855
      Left            =   14520
      TabIndex        =   9
      Top             =   6600
      Width           =   4455
   End
   Begin VB.TextBox Text8 
      Height          =   855
      Left            =   14520
      TabIndex        =   8
      Top             =   5400
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   855
      Left            =   14520
      TabIndex        =   7
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   7680
      TabIndex        =   6
      Top             =   8640
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   7680
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   7680
      TabIndex        =   4
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   4320
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7680
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
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   735
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
      Left            =   11160
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "update_package"
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
MsgBox "Please Select package id !", vbExclamation, "Update packages"
Exit Sub
End If


If Text1.Text = "" Then
MsgBox "Please Enter State !", vbExclamation, "Update packages"
Exit Sub
End If


If Text2.Text = "" Then
MsgBox "Please Enter a Spot !", vbExclamation, "Update packages"
Exit Sub
End If



If Text3.Text = "" Then
MsgBox "Please Enter Total distance !", vbExclamation, "Update packages"
Exit Sub
End If

If Text4.Text = "" Then
MsgBox "Please Enter car charges !", vbExclamation, "Update packages"
Exit Sub
End If



If Text5.Text = "" Then
MsgBox "Please Enter Bus Charges !", vbExclamation, "Update packages"
Exit Sub
End If


If Text6.Text = "" Then
MsgBox "Please Enter Stay Cost !", vbExclamation, "Update packages"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please Enter hotels1 !", vbExclamation, "Update packages"
Exit Sub
End If

If Text8.Text = "" Then
MsgBox "Please Enter hotels2 !", vbExclamation, "Update packages"
Exit Sub
End If

If Text9.Text = "" Then
MsgBox "Please Enter hotels3 !", vbExclamation, "Update packages"
Exit Sub
End If

Call connect
str1 = "update packages set state= '" & Text1.Text & "',spot='" & Text2.Text & "',distance=" & Text3.Text & ",car_charges=" & Text4.Text & ",bus_charges=" & Text5.Text & ",stay_cost=" & Text6.Text & ",hotel1='" & Text7.Text & "',hotel2='" & Text8.Text & "',hotel3='" & Text9.Text & "' where package_id=" & Combo1.Text & " "

con.Execute str1
MsgBox "package Details Updated Successfully...", vbInformation, "Update package Details"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""


con.Close
End Sub

Private Sub Image3_Click()
Unload Me
End Sub
