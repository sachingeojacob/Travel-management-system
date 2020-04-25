VERSION 5.00
Begin VB.Form update_customer 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21495
   LinkTopic       =   "Form1"
   Picture         =   "update_customer.frx":0000
   ScaleHeight     =   11085
   ScaleWidth      =   21495
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   13560
      TabIndex        =   9
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   13560
      TabIndex        =   8
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   9000
      Width           =   5055
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   7920
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   7080
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   6240
      Width           =   5055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   4320
      Width           =   5055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   0
      Text            =   "SELECT CUSTOMER ID"
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Image cancel 
      Height          =   1095
      Left            =   14880
      Top             =   9840
      Width           =   3495
   End
   Begin VB.Image update 
      Height          =   1095
      Left            =   10320
      Top             =   9840
      Width           =   3495
   End
End
Attribute VB_Name = "update_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Me
End Sub


Private Sub Combo1_click()
Call connect
str1 = "select * from customer where customer_id=" & Combo1.Text & ""

rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF


If rs.Fields("gender").Value = Option1.Caption Then
Option1.Value = True
Else
Option2.Value = True
End If



Text1.Text = rs.Fields("customer_name").Value
Text2.Text = rs.Fields("street").Value
Text3.Text = rs.Fields("city").Value
Text4.Text = rs.Fields("states").Value
Text5.Text = rs.Fields("pin_code").Value
Text6.Text = rs.Fields("email_id").Value
Text7.Text = rs.Fields("mobile").Value
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
str1 = "select customer_id from customer"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub update_Click()

''male

If Option1.Value = True Then
Call connect
str1 = "update customer set customer_name= '" & Text1.Text & "',gender='" & Option1.Caption & "',street= '" & Text2.Text & "',city='" & Text3.Text & "',states='" & Text4.Text & "',pin_code=" & Text5.Text & ",email_id='" & Text6.Text & "',mobile=" & Text7.Text & " where customer_id=" & Combo1.Text & " "

con.Execute str1
MsgBox "Customer Details Updated Successfully...", vbInformation, "Update Customer Details"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = "Select customer to update"


con.Close

End If

''female

If Option2.Value = True Then
Call connect
str1 = "update customer set customer_name= '" & Text1.Text & "',gender='" & Option2.Caption & "',street= '" & Text2.Text & "',city='" & Text3.Text & "',states='" & Text4.Text & "',pin_code=" & Text5.Text & ",email_id='" & Text6.Text & "',mobile=" & Text7.Text & " where customer_id=" & Combo1.Text & " "

con.Execute str1
MsgBox "Customer Details Updated Successfully...", vbInformation, "Update Customer Details"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = "Select customer to update"

con.Close

End If

End Sub
