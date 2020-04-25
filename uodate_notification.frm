VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form update_notification 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21075
   LinkTopic       =   "Form1"
   Picture         =   "uodate_notification.frx":0000
   ScaleHeight     =   11040
   ScaleWidth      =   21075
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   9480
      TabIndex        =   6
      Top             =   6000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1296
      _Version        =   393216
      Format          =   127008769
      CurrentDate     =   43410
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   12720
      TabIndex        =   5
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   9480
      TabIndex        =   4
      Top             =   8760
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   11040
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   9480
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   9480
      TabIndex        =   1
      Top             =   4440
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9480
      TabIndex        =   0
      Text            =   "SELECT NOTIFICATION ID"
      Top             =   3360
      Width           =   4215
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
Attribute VB_Name = "update_notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()

Call connect
str1 = "select * from notifications where notification_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset

Text1.Text = rs.Fields(1)
DTPicker1.Value = rs.Fields(2)
Text3.Text = rs.Fields(3)
Text4.Text = rs.Fields(4)
Combo2.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)

rs.MoveNext



rs.Close
con.Close
End Sub

Private Sub Form_Load()



If con.State = adStateOpen Then
rs.Close
con.Close
End If


Call connect
str1 = "select * from notifications"
rs.Open str1, con, adOpenKeyset

Do While Not rs.EOF

Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub Image1_Click()


Call connect
str1 = "update notifications set purpose='" & Text1.Text & "' , notification_date='" & DTPicker1.Value & "',notification_time_H=" & Text3.Text & ",notification_time_M=" & Text4.Text & ",am_pm='" & Combo2.Text & "',venue='" & Text5.Text & "'where notification_id=" & Combo1.Text & " "
con.Execute str1
MsgBox "Notification Updated Successfully", vbInformation, "Update Notification"
con.Close

End Sub

Private Sub Image2_Click()
Unload Me
End Sub

