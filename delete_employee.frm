VERSION 5.00
Begin VB.Form delete_employee 
   Caption         =   "Form2"
   ClientHeight    =   11100
   ClientLeft      =   1410
   ClientTop       =   1110
   ClientWidth     =   20460
   LinkTopic       =   "Form2"
   Picture         =   "delete_employee.frx":0000
   ScaleHeight     =   11100
   ScaleWidth      =   20460
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text15 
      Height          =   615
      Left            =   4440
      TabIndex        =   16
      Top             =   8880
      Width           =   4695
   End
   Begin VB.TextBox Text14 
      Height          =   615
      Left            =   4440
      TabIndex        =   15
      Top             =   7080
      Width           =   4695
   End
   Begin VB.TextBox Text13 
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   6120
      Width           =   4695
   End
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   4440
      TabIndex        =   13
      Top             =   5160
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   16200
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Height          =   615
      Left            =   14160
      TabIndex        =   11
      Top             =   9480
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   9720
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   8040
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Text            =   "SELECT EMPLOYEE ID"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   14160
      TabIndex        =   6
      Top             =   8520
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   14160
      TabIndex        =   5
      Top             =   7560
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   14160
      TabIndex        =   4
      Top             =   6600
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   14160
      TabIndex        =   3
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   14160
      TabIndex        =   2
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   16920
      Top             =   10560
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   11880
      Top             =   10560
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   10800
      Top             =   1320
      Width           =   3135
   End
End
Attribute VB_Name = "delete_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub Form_Load()
'Public con As New ADODB.Connection
'Public rs As New ADODB.Recordset
'Public cmd As New ADODB.Command
'Public str As String
'Public str1 As String
'Public Sub connect()
'con.Provider = "sqloledb"
'str1 = "server=DESKTOP-HOTR91D\SQLEXPRESS;database=master;trusted_connection=yes"
'con.Open str1

'End Sub

'Private Sub Combo1_click()
'Call connect
'str1 = "select * from employee_management where employee_id= '" & Combo1.Text & "' "
'rs.Open str1, con, adOpenKeyset
'Do While Not rs.EOF
'Text1.Text = rs.Fields("employee_id").Value
'Text2.Text = rs.Fields("employee_name").Value
'Text3.Text = rs.Fields("mobile").Value
'Text4.Text = rs.Fields("email_id").Value
'Text5.Text = rs.Fields("basic_pay").Value
'Text6.Text = rs.Fields("house_name").Value
'Text7.Text = rs.Fields("village").Value
'Text8.Text = rs.Fields("city").Value
'Text9.Text = rs.Fields("town").Value
'Text10.Text = rs.Fields("pin_code").Value
'Text11.Text = rs.Fields("states").Value
'Text12.Text = rs.Fields("date_of_birth").Value
'Text13.Text = rs.Fields("gender").Value
'Text14.Text = rs.Fields("date_of_join").Value
'Text15.Text = rs.Fields("email_id").Value
'Picture1.Picture = LoadPicture(rs.Fields("pic"))


'rs.MoveNext
'Loop
'rs.Close
'con.Close




'End Sub



'Private Sub Command3_Click()
'DELETEEMPLOYEE.Refresh
'Combo1.Refresh

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = " "

'Combo1.Text = " "

'End Sub

Private Sub Form_Load()
Dim str1 As String
Call connect
str1 = "select employee_id from employee_management"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields("employee_id").Value)
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub Image1_Click()
Call connect
str1 = "select * from employee_management where employee_id= '" & Combo1.Text & "' "
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Text = rs.Fields("employee_id").Value
Text2.Text = rs.Fields("employee_name").Value
Text3.Text = rs.Fields("mobile").Value
Text4.Text = rs.Fields("basic_pay").Value
Text5.Text = rs.Fields("house_name").Value
Text6.Text = rs.Fields("village").Value
Text7.Text = rs.Fields("city").Value
Text8.Text = rs.Fields("town").Value
Text9.Text = rs.Fields("pin_code").Value
Text10.Text = rs.Fields("states").Value
Text11.Text = rs.Fields("country").Value
Text12.Text = rs.Fields("date_of_birth").Value
Text13.Text = rs.Fields("gender").Value
Text14.Text = rs.Fields("date_of_join").Value
Text15.Text = rs.Fields("email_id").Value
Picture1.Picture = LoadPicture(rs.Fields("pic"))


rs.MoveNext
Loop
rs.Close
con.Close

End Sub

Private Sub Image2_Click()
'Call connect
'str1 = "delete from employee_management where employee_id= '" & Combo1.Text & "' "
'MsgBox " SUCCESSFULLY DELETED "


'con.Execute str1
'con.Close

If Combo1.Text = "" Then
MsgBox "Select Employee ID to delete", vbExclamation, "Delete Employee"
Exit Sub
End If

Call connect

confirm = MsgBox("Are you sure you want to delete this Employee ?", vbYesNo, "Delete an Employee")
If confirm = vbYes Then

str1 = "delete from employee_management where employee_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "Employee Deleted Successfully...", vbInformation, "Delete an Employee"

delete_employee.Refresh

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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Picture1.Picture = Nothing

Else
MsgBox "Employee not deleted ", vbInformation, "Delete an Employee"
End If
con.Close
End Sub

Private Sub Image3_Click()
Unload Me
End Sub
