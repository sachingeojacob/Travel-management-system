VERSION 5.00
Begin VB.Form salary_details 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   945
   ClientTop       =   630
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "salary_details.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   855
      Left            =   11160
      TabIndex        =   7
      Top             =   8280
      Width           =   5295
   End
   Begin VB.TextBox Text6 
      Height          =   855
      Left            =   15840
      TabIndex        =   6
      Top             =   6600
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   15840
      TabIndex        =   5
      Top             =   4680
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   15840
      TabIndex        =   4
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   6840
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   5280
      Width           =   5175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      Text            =   "SELECT EMPLOYEE ID"
      Top             =   4080
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5760
      TabIndex        =   0
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Image total_salary 
      Height          =   855
      Left            =   6360
      Top             =   8280
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   13920
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   9840
      Top             =   10320
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   5760
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   5040
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Image Image_save 
      Height          =   855
      Left            =   1800
      Top             =   10320
      Width           =   2535
   End
End
Attribute VB_Name = "salary_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

Call connect
str1 = "select employee_name,basic_pay from employee_management where employee_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF

Text2.Text = rs.Fields(0)
Text3.Text = rs.Fields(1)


rs.MoveNext

Loop
rs.Close
str2 = "select da,hra,providend_fund,total_salary from salary_management where employee_id=" & Combo1.Text & ""
rs.Open str2, con, adOpenKeyset
Do While Not rs.EOF

Text4.Text = rs.Fields("da").Value
Text5.Text = rs.Fields("hra").Value
Text6.Text = rs.Fields("providend_fund").Value
Text7.Text = rs.Fields("total_salary").Value

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





Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image4_Click()
salary_view.Show

End Sub

Private Sub total_salary_Click()
Text7.Text = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text)
End Sub

Private Sub Image_save_Click()

If Combo1.Text = "" Then
MsgBox "Please choose Employee ID", vbExclamation, "Salary Management"
Exit Sub
End If

If Text4.Text = "" Then
MsgBox "Please Enter Da !", vbExclamation, "Salary Management"
Exit Sub
End If


If Text5.Text = "" Then
MsgBox "Please Enter  HRA !", vbExclamation, "Salary Management"
Exit Sub
End If



If Text6.Text = "" Then
MsgBox "Please Enter PF !", vbExclamation, "Salary Managemet"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please calculate Total salary", vbExclamation, "Salary Management"
Exit Sub
End If



Call connect
str = "insert into salary_management(salary_id,employee_id,employee_name,basic_pay,hra,da,providend_fund,total_salary) values (" & Text1.Text & "," & Combo1.Text & ",'" & Text2.Text & "','" & Text3.Text & "'," & Text4.Text & ",'" & Text5.Text & "'," & Text6.Text & "," & Text7.Text & ")"
con.Execute str
MsgBox "Salary saved successfully...", vbInformation, "Salary Management"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
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

'Private Sub Command8_Click()
'Call connect
'str1 = " update salary set da= ' " & Text2.Text & " ',hra=' " & Text3.Text & " ',pf=' " & Text4.Text & " ' where eid=' " & Combo1.Text & "'"
'con.Execute str1
'MsgBox " succesfully updated", vbInformation
'con.Close
'Adodc1.Recordset.Update
'End Sub




Private Sub Image5_Click()
Call connect

If Combo1.Text = "SELECT EMPLOYEE ID" Then
MsgBox "Select Emoloyee ID to delete", vbExclamation, "Delete Salary"
Exit Sub
End If

confirm = MsgBox("Are you sure you want to delete this Salary details ?", vbYesNo, "Delete this salary details")
If confirm = vbYes Then

str1 = "delete from salary_management where employee_id= '" & Combo1.Text & "' "
con.Execute str1
MsgBox "successfully deleted ", vbInformation
End If
con.Close
End Sub
