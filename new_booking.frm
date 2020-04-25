VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form new_booking 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "new_booking.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   14160
      TabIndex        =   9
      Top             =   5280
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   8280
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Text            =   "SELECT CUSTOMER ID"
      Top             =   3360
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   9360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   128253953
      CurrentDate     =   43414
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   6240
      Width           =   4095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "new_booking.frx":55CB6
      Left            =   5400
      List            =   "new_booking.frx":55CB8
      TabIndex        =   4
      Text            =   "SELECT HOTEL TO STAY"
      Top             =   7320
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   5280
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   4320
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   2400
      Width           =   4095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   14160
      TabIndex        =   0
      Text            =   "SELECT PACKAGE ID"
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Image CANCEL 
      Height          =   1095
      Left            =   15360
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Image SAVE 
      Height          =   1095
      Left            =   12120
      Top             =   9240
      Width           =   3015
   End
End
Attribute VB_Name = "new_booking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Combo1_Click()


Text5.Text = ""


Call connect
str1 = "select customer_name,mobile,email_id from customer where customer_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text2.Text = rs.Fields("customer_name").Value
Text3.Text = rs.Fields("mobile").Value
Text6.Text = rs.Fields("email_id").Value


rs.MoveNext

Loop
rs.Close
str2 = "select number_of_persons,boarding_date from booking where customer_id=" & Combo1.Text & ""
rs.Open str2, con, adOpenKeyset
Do While Not rs.EOF

DTPicker1.Value = rs.Fields("boarding_date").Value
Text5.Text = rs.Fields("number_of_persons").Value

rs.MoveNext

Loop
rs.Close
con.Close

End Sub



Private Sub Combo2_click()



Text5.Text = ""

Call connect
str1 = "select state,hotel1,hotel2,hotel3,spot from packages where package_id=" & Combo2.Text & ""
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF



With Combo3
    

    .AddItem rs.Fields("hotel1").Value
    .AddItem rs.Fields("hotel2").Value
    .AddItem rs.Fields("hotel3").Value
   
    
    End With

Text4.Text = rs.Fields("state").Value
Text7.Text = rs.Fields("spot").Value



rs.MoveNext

Loop
rs.Close
str2 = "select number_of_persons,boarding_date from booking where package_id=" & Combo2.Text & ""
rs.Open str2, con, adOpenKeyset
Do While Not rs.EOF

DTPicker1.Value = rs.Fields("boarding_date").Value
Text5.Text = rs.Fields("number_of_persons").Value

rs.MoveNext

Loop
rs.Close
con.Close


End Sub

Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If



Call connect
str1 = "select * from booking"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("booking_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close


Call connect
str1 = "select customer_id from customer"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close

Call connect
str1 = "select package_id from packages"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo2.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub save_Click()


If Combo1.Text = "SELECT CUSTOMER ID" Then
MsgBox "Please choose customer ID", vbExclamation, "booking Management"
Exit Sub
End If

If Combo2.Text = "SELECT PACKAGE ID" Then
MsgBox "Please choose package ID", vbExclamation, "booking Management"
Exit Sub
End If

If Combo3.Text = "SELECT HOTEL TO STAY" Then
MsgBox "Please choose hotel to stay", vbExclamation, "booking Management"
Exit Sub
End If





Call connect
str1 = "insert into booking(booking_id,customer_id,package_id,customer_name,mobile,state,hotel,number_of_persons,boarding_date,email_id,spot) values (" & Text1.Text & "," & Combo1.Text & "," & Combo2.Text & ",'" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Combo3.Text & "','" & Text5.Text & "','" & DTPicker1 & " ','" & Text6.Text & "','" & Text7.Text & "')"
con.Execute str1
MsgBox "booking saved successfully...", vbInformation, "booking Management"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = "select another customer id"
Combo2.Text = "select another package id"
Combo3.Text = "select hotel to stay"
con.Close


Call connect
str1 = "select * from salary_management"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("salary_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close

End Sub

