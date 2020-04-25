VERSION 5.00
Begin VB.Form add_customer 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18675
   LinkTopic       =   "Form1"
   Picture         =   "add_customer.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   18675
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   13680
      TabIndex        =   9
      Top             =   6120
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   13680
      TabIndex        =   8
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   9000
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   7920
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   7080
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   6240
      Width           =   4695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Image cancel 
      Height          =   1095
      Left            =   14880
      Top             =   9840
      Width           =   3495
   End
   Begin VB.Image save 
      Height          =   975
      Left            =   10320
      Top             =   9840
      Width           =   3495
   End
End
Attribute VB_Name = "add_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim n As Integer

Call connect
str1 = "select * from customer"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("customer_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub save_Click()

If Text2.Text = "" Then
MsgBox "Please Enter Customer Name !", vbExclamation, "Add Employee"
Exit Sub
End If


If Option1.Value = False And Option2.Value = False Then
MsgBox "Please Select Gender !", vbExclamation, "Add Customer"
Exit Sub
End If


If Text3.Text = "" Then
MsgBox "Please Enter Street !", vbExclamation, "Add Customer"
Exit Sub
End If



If Text4.Text = "" Then
MsgBox "Please Enter City !", vbExclamation, "Add Customer"
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Please Enter State !", vbExclamation, "Add Customer"
Exit Sub
End If



If Text6.Text = "" Then
MsgBox "Please Enter Pin code !", vbExclamation, "Add Customer"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please Enter email !", vbExclamation, "Add Customer"
Exit Sub
End If


If Text8.Text = "" Then
MsgBox "Please Enter mobile !", vbExclamation, "Add Customer"
Exit Sub
End If



If Option1.Value = True Then

Call connect

str1 = "insert into customer(customer_id,customer_name,gender,street,city,states,pin_code,mobile,email_id)values (" & Text1.Text & ",'" & Text2.Text & "','" & Option1.Caption & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "'," & Text6.Text & "," & Text8.Text & ",'" & Text7.Text & "')"

con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save customer Details"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""


Option1.Value = False
con.Close

End If


If Option2.Value = True Then

Call connect

str1 = "insert into customer(customer_id,customer_name,gender,street,city,states,pin_code,mobile,email_id)values (" & Text1.Text & ",'" & Text2.Text & "','" & Option2.Caption & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "'," & Text6.Text & "," & Text8.Text & ",'" & Text7.Text & "')"


con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save customer details"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""


Option2.Value = False

con.Close


End If

''id

Call connect
str1 = "select * from customer"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Dim n As Integer
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("customer_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub


Private Sub Text8_Validate(cancel As Boolean)

'mobile

If Not (IsNumeric(Text8.Text)) Then
MsgBox " Please Enter Phone Number Numerically !", vbExclamation, "Add Employee"
End If

If Len(Text8.Text) < 10 Or Len(Text8.Text) > 10 Then
MsgBox "Enter the phone number in 10 digits !", vbExclamation, "Add Customer"
End If

End Sub

Private Sub Text6_Validate(cancel As Boolean)

'pin

If Not (IsNumeric(Text6.Text)) Then
MsgBox " Please Enter PIN Code Numerically !", vbExclamation, "Add Customer"
End If

If Len(Text6.Text) < 6 Or Len(Text6.Text) > 6 Then
MsgBox "Please Enter a Valid PIN Code !", vbExclamation, "Add Customer"
End If
End Sub

'Private Sub Text2_Validate(cancel As Boolean)

'name
'If (Not (isNotaNumeric(Text2.Text))) Then
 'MsgBox " Please Enter correct name !", vbExclamation, "Add Customer"
'End If
'End Sub
