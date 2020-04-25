VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form add_employee 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   3435
   ClientTop       =   2925
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "add_employee.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16560
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   14040
      TabIndex        =   16
      Top             =   9480
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   17520
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Height          =   615
      Left            =   14040
      TabIndex        =   14
      Top             =   8520
      Width           =   4815
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   14040
      TabIndex        =   13
      Top             =   7560
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   14040
      TabIndex        =   12
      Top             =   6600
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   14040
      TabIndex        =   11
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   14040
      TabIndex        =   10
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   14040
      TabIndex        =   9
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   9360
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   6600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      Format          =   129040385
      CurrentDate     =   43324
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   5520
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   5520
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   4560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      Format          =   129040385
      CurrentDate     =   43324
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   8400
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   7440
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   17640
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   16920
      Top             =   10560
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   600
      Top             =   10560
      Width           =   3135
   End
End
Attribute VB_Name = "add_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public cmd As New ADODB.Command
Public str As String
Public str1 As String

Public str2 As String

Public Sub connect()
con.Provider = "sqloledb"
str1 = "server=DESKTOP-HOTR91D;database=master;trusted_connection=yes"
con.Open str1
End Sub
           


Private Sub Image1_Click()
CommonDialog1.ShowOpen
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Form_Load()
Dim n As Integer

Call connect
str1 = "select * from employee_management"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("employee_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image2_Click()
If Text2.Text = "" Then
MsgBox "Please Enter Employee Name !", vbExclamation, "Add Employee"
Exit Sub
End If


If Option1.Value = False And Option2.Value = False Then
MsgBox "Please Select Gender !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text3.Text = "" Then
MsgBox "Please Enter Mobile Number !", vbExclamation, "Add Employee"
Exit Sub
End If



If Text4.Text = "" Then
MsgBox "Please Enter E-Mail ID !", vbExclamation, "Add Employee"
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Please Enter Basic Pay !", vbExclamation, "Add Employee"
Exit Sub
End If



If Text6.Text = "" Then
MsgBox "Please Enter House Name !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please Enter Village !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text8.Text = "" Then
MsgBox "Please Enter City !", vbExclamation, "Add Employee"
Exit Sub
End If

If Text9.Text = "" Then
MsgBox "Please Enter Town !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text10.Text = "" Then
MsgBox "Please Enter PIN Code !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text11.Text = "" Then
MsgBox "Please Enter State !", vbExclamation, "Add Employee"
Exit Sub
End If



If Text12.Text = "" Then
MsgBox "Please Enter Country !", vbExclamation, "Add Employee"
Exit Sub
End If




If Picture1.Picture = 0 Then
MsgBox "Please Upload Employees Photo !", vbExclamation, "Add Employee"
Exit Sub
End If






If Option1.Value = True Then

Call connect

str1 = "insert into employee_management(employee_id,employee_name,date_of_birth,gender,date_of_join,mobile,email_id,basic_pay,house_name,village,city,town,pin_code,states,country,pic)values (" & Text1.Text & ",'" & Text2.Text & "','" & DTPicker1 & "','" & Option1.Caption & "','" & DTPicker2 & "'," & Text3.Text & ",'" & Text4.Text & "'," & Text5.Text & ",'" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "'," & Text10.Text & ",'" & Text11.Text & "','" & Text12.Text & "','" & str & "')"

con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save Employee Details"

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




Picture1.Picture = Nothing


Option1.Value = False

'Option2.Value = False



con.Close

End If

If Option2.Value = True Then

Call connect

str1 = "insert into employee_management(employee_id,employee_name,date_of_birth,gender,date_of_join,mobile,email_id,basic_pay,house_name,village,city,town,pin_code,states,country,pic)values (" & Text1.Text & ",'" & Text2.Text & "','" & DTPicker1 & "','" & Option2.Caption & "','" & DTPicker2 & "'," & Text3.Text & ",'" & Text4.Text & "'," & Text5.Text & ",'" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "'," & Text10.Text & ",'" & Text11.Text & "','" & Text12.Text & "','" & str & "')"


con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save EmployeeDetails"

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

Option2.Value = False

Picture1.Picture = Nothing

con.Close


End If

''id

Call connect
str1 = "select * from employee_management"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Dim n As Integer
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("employee_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub



Private Sub Text5_Validate(cancel As Boolean)

If Not (IsNumeric(Text5.Text)) Then
MsgBox " Please Enter Basic Pay Numerically !", vbExclamation, "Add Employee"
End If

End Sub






Private Sub Text10_Validate(cancel As Boolean)

'pin

If Not (IsNumeric(Text10.Text)) Then
MsgBox " Please Enter PIN Code Numerically !", vbExclamation, "Add Employee"
End If

If Len(Text10.Text) < 6 Or Len(Text10.Text) > 6 Then
MsgBox "Please Enter a Valid PIN Code !", vbExclamation, "Add Employee"
End If

End Sub






Private Sub Text3_Validate(cancel As Boolean)

'mobile

If Not (IsNumeric(Text3.Text)) Then
MsgBox " Please Enter Phone Number Numerically !", vbExclamation, "Add Employee"
End If

If Len(Text3.Text) < 10 Or Len(Text3.Text) > 10 Then
MsgBox "Enter the phone number in 10 digits !", vbExclamation, "Add Employee"
End If

End Sub

'Private Sub Text2_Validate(cancel As Boolean)

'name

'If (Not (isNotNumeric(Text2.Text))) Then
'MsgBox " Please Enter correct name !", vbExclamation, "Add Employee"
'End If




