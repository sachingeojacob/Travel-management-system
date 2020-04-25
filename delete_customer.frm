VERSION 5.00
Begin VB.Form delete_customer 
   Caption         =   "Form1"
   ClientHeight    =   11745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19305
   LinkTopic       =   "Form1"
   Picture         =   "delete_customer.frx":0000
   ScaleHeight     =   11745
   ScaleWidth      =   19305
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   13680
      TabIndex        =   8
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   13680
      TabIndex        =   7
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   9000
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   7920
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   7080
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   6120
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   5280
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   4320
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Text            =   "SELECT CUSTOMER ID"
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Image cancel 
      Height          =   975
      Left            =   15000
      Top             =   9840
      Width           =   3375
   End
   Begin VB.Image delete 
      Height          =   975
      Left            =   10320
      Top             =   9840
      Width           =   3495
   End
End
Attribute VB_Name = "delete_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Combo1_click()

Call connect
str1 = "select * from customer where customer_id= '" & Combo1.Text & "' "
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Text = rs.Fields("customer_name").Value
Text2.Text = rs.Fields("gender").Value
Text3.Text = rs.Fields("street").Value
Text4.Text = rs.Fields("city").Value
Text5.Text = rs.Fields("states").Value
Text6.Text = rs.Fields("pin_code").Value
Text7.Text = rs.Fields("email_id").Value
Text8.Text = rs.Fields("mobile").Value


rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub delete_Click()
If Combo1.Text = "SELECT CUSTOMER ID" Then
MsgBox "Select Customer ID to delete", vbExclamation, "Delete customer"
Exit Sub
End If

Call connect

confirm = MsgBox("Are you sure you want to delete this Customer ?", vbYesNo, "Delete a Customer")
If confirm = vbYes Then

str1 = "delete from customer where customer_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "customer Deleted Successfully...", vbInformation, "Delete a customer"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Combo1.Text = "select another customer to delete"

Else
MsgBox "Employee not deleted ", vbInformation, "Delete an Employee"
End If
con.Close
End Sub

Private Sub Form_Load()

Dim str1 As String
Call connect
str1 = "select customer_id from customer"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields("customer_id").Value)
rs.MoveNext
Loop
rs.Close
con.Close

End Sub
