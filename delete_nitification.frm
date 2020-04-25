VERSION 5.00
Begin VB.Form delete_nitification 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "delete_nitification.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   855
      Left            =   9480
      TabIndex        =   6
      Top             =   8760
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   12120
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   10800
      TabIndex        =   4
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   9360
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   9360
      TabIndex        =   2
      Top             =   5880
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   9360
      TabIndex        =   1
      Top             =   4560
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9360
      TabIndex        =   0
      Text            =   "SELECT NOTIFICATION ID"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   12480
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   16920
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   16920
      Top             =   4680
      Width           =   2415
   End
End
Attribute VB_Name = "delete_nitification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If

Call connect
str1 = "select notification_id from notifications"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF


Combo1.AddItem (rs.Fields(0))

rs.MoveNext
Loop
rs.Close
con.Close
End Sub



Private Sub Image1_Click()

If Combo1.Text = "" Then
MsgBox "Select Notification ID to delete", vbExclamation, "Delete Notification"
Exit Sub
End If


Call connect

confirm = MsgBox("Are you sure you want to delete this Notification ?", vbYesNo, "Delete a Notification")
If confirm = vbYes Then
str1 = "delete from notifications where notification_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "Notification Deleted Successfully...", vbInformation, "Delete a Notification"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

Combo1.Text = ""

Else
MsgBox "Notification not deleted ", vbInformation, "Delete a Notification"
End If

con.Close
End Sub


Private Sub Image2_Click()
Unload Me
End Sub



Private Sub Image3_Click()
Call connect
str1 = "select * from notifications where notification_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset

Text1.Text = rs.Fields(1)
Text2.Text = rs.Fields(2)
Text3.Text = rs.Fields(3)
Text4.Text = rs.Fields(4)
Text5.Text = rs.Fields(5)
Text6.Text = rs.Fields(6)
rs.MoveNext



rs.Close
con.Close
End Sub
