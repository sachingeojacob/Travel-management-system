VERSION 5.00
Begin VB.Form create_account 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21285
   LinkTopic       =   "Form1"
   Picture         =   "create_account.frx":0000
   ScaleHeight     =   11085
   ScaleWidth      =   21285
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   9000
      TabIndex        =   5
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   9000
      TabIndex        =   4
      Top             =   7320
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   9000
      TabIndex        =   3
      Top             =   6120
      Width           =   4695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "EMPLOYEE"
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   2160
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ADMIN"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9000
      TabIndex        =   0
      Text            =   "SELECT EMPLOYEE ID"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   11520
      Top             =   9600
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4200
      Top             =   9600
      Width           =   4095
   End
End
Attribute VB_Name = "create_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

Call connect
str1 = "select employee_management.employee_name,login_page.password from employee_management,login_page where employee_management.employee_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF



Text1.Text = rs.Fields("employee_name").Value
Text2.Text = rs.Fields("password").Value

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


'Text1.Enabled = False

Call connect
str1 = "select employee_id from employee_management"
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
Call connect

st = False

'Administrator Account

If Text2.Text = Text3.Text Then
st = True
End If


If Option1.Value = True And st = True Then
str1 = "insert into login_page(user_name,employee_id,password,account_type) values('" & Text1.Text & "'," & Combo1.Text & ",'" & Text2.Text & "','" & Option1.Caption & "')"
con.Execute str1
MsgBox "Account created successfully", vbInformation, "Create Account"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End If

'Employee Account Account

If Option2.Value = True And st = True Then
str1 = "insert into login_page(user_name,employee_id,password,account_type) values('" & Text1.Text & "'," & Combo1.Text & ",'" & Text2.Text & "','" & Option2.Caption & "')"

con.Execute str1
MsgBox "Account created successfully", vbInformation, "Create Account"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End If

If st = False Then
MsgBox "Password Mismatch", vbInformation, "Create Account"
End If

con.Close
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

