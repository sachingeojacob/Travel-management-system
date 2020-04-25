VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form add_notification 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "add_notification.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   855
      Left            =   9480
      TabIndex        =   6
      Top             =   5880
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1508
      _Version        =   393216
      Format          =   124452865
      CurrentDate     =   43410
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   9360
      TabIndex        =   5
      Top             =   8880
      Width           =   6015
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   9360
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   11880
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   13920
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   9360
      TabIndex        =   1
      Top             =   4440
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   9360
      TabIndex        =   0
      Top             =   3120
      Width           =   5895
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
Attribute VB_Name = "add_notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If

With Combo3
    

    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    
    End With
    
    
    
With Combo2
    

    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    .AddItem "32"
    .AddItem "33"
    .AddItem "34"
    .AddItem "35"
    .AddItem "36"
    .AddItem "37"
    .AddItem "38"
    .AddItem "39"
    .AddItem "40"
    .AddItem "41"
    .AddItem "42"
    .AddItem "43"
    .AddItem "44"
    .AddItem "45"
    .AddItem "46"
    .AddItem "47"
    .AddItem "48"
    .AddItem "49"
    .AddItem "50"
    .AddItem "51"
    .AddItem "52"
    .AddItem "53"
    .AddItem "54"
    .AddItem "55"
    .AddItem "56"
    .AddItem "57"
    .AddItem "58"
    .AddItem "59"
    

    
    End With
    


With Combo1
.AddItem "AM"
.AddItem "PM"

End With
Call connect
str1 = "select * from notifications"
rs.Open str1, con, adOpenKeyset
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("notification_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub


Private Sub Image1_Click()

Call connect

str1 = "insert into notifications(notification_id,purpose,notification_date,notification_time_H,notification_time_M,am_pm,venue)values (" & Text1.Text & ",'" & Text2.Text & "','" & DTPicker1.Value & "','" & Combo3.Text & "','" & Combo2.Text & "','" & Combo1.Text & "',' " & Text4.Text & "')"
con.Execute str1
MsgBox "Notification added successfully", vbInformation, "Add Notification"



Text1.Text = ""
Text2.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Text4.Text = ""
Combo1.Text = ""
con.Close

Call connect
str1 = "select * from notifications"
rs.Open str1, con, adOpenKeyset
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("notification_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub



Private Sub Image2_Click()
Unload Me
End Sub





