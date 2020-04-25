VERSION 5.00
Begin VB.Form delete_booking 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21450
   LinkTopic       =   "Form1"
   Picture         =   "delete_booking.frx":0000
   ScaleHeight     =   11085
   ScaleWidth      =   21450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   6600
      Width           =   4335
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   14160
      TabIndex        =   9
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   14160
      TabIndex        =   8
      Top             =   3360
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   9360
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   8280
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   7200
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Text            =   "SELECT BOOKING ID"
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image cancel 
      Height          =   1095
      Left            =   15360
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Image delete 
      Height          =   1095
      Left            =   12120
      Top             =   9240
      Width           =   3015
   End
End
Attribute VB_Name = "delete_booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Combo1_Click()


Call connect
str1 = "select * from booking where booking_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset


Text1.Text = rs.Fields("customer_id").Value
Text2.Text = rs.Fields("customer_name")
Text3.Text = rs.Fields("mobile")
Text4.Text = rs.Fields("state")
Text5.Text = rs.Fields("hotel")
Text6.Text = rs.Fields("number_of_persons")
Text7.Text = rs.Fields("boarding_date")
Text8.Text = rs.Fields("package_id")
Text9.Text = rs.Fields("email_id")
Text10.Text = rs.Fields("spot")
rs.MoveNext



rs.Close
con.Close

End Sub



Private Sub delete_Click()


If Combo1.Text = "SELECT BOOKING ID" Then
MsgBox "Select BOOKING ID to delete", vbExclamation, "Delete Booking"
Exit Sub
End If


Call connect

confirm = MsgBox("Are you sure you want to delete this Booking ?", vbYesNo, "Delete a Booking")
If confirm = vbYes Then
str1 = "delete from booking where booking_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "Booking Deleted Successfully...", vbInformation, "Delete a Booking"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Combo1.Text = "select another booking id"

Else
MsgBox "Booking not deleted ", vbInformation, "Delete a Booking"
End If

con.Close

End Sub

Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If

Call connect
str1 = "select booking_id from booking"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF


Combo1.AddItem (rs.Fields(0))

rs.MoveNext
Loop
rs.Close
con.Close

End Sub
