VERSION 5.00
Begin VB.Form addemp 
   Caption         =   "ADDEMP"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "addemp.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "BROWSE"
      Height          =   495
      Left            =   16560
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   15720
      ScaleHeight     =   2115
      ScaleWidth      =   2595
      TabIndex        =   22
      Top             =   1800
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C000&
      Caption         =   "FEMALE"
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   6720
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C000&
      Caption         =   "MALE"
      Height          =   495
      Left            =   5520
      TabIndex        =   20
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE"
      Height          =   735
      Left            =   11760
      TabIndex        =   18
      Top             =   8400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   11760
      TabIndex        =   17
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "SAVE"
      Height          =   735
      Left            =   11760
      TabIndex        =   16
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   8040
      TabIndex        =   15
      Top             =   9960
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   9960
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   9000
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   7800
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   5760
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "GENDER"
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      Caption         =   "PASSWORD"
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "USER NAME"
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      Caption         =   "BASIC PAY"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   9000
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      Caption         =   "QUALIFICATION"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "EMAIL ID"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "MOBILE NUMBER"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label empid 
      BackColor       =   &H00C0C000&
      Caption         =   "EMPLOYEE ID"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "NAME"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
   End
End
Attribute VB_Name = "addemp"
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
Dim status As String
Dim t As String
Dim commondialog1 As String

Public SQL As String
Public Sub connect()
con.Provider = "SQLOLEDB"
str1 = "server=DESKTOP-HOTR91D\SQLEXPRESS;database=master;trusted_connection=yes"
con.Open str1
End Sub


Private Sub Command3_Click()
Unload Me
End Sub





Private Sub Form_Load()
Dim n As Integer

'If con.State = adStateOpen Then
'rs.Close
'con.Close
'End If

Call connect
str1 = "select * from emp"
rs.Open str1, con, adOpenKeyset
con.Execute str1
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("emp_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub


Private Sub Command1_Click()
If Text2.Text = "" Then
MsgBox "Please Enter Employee Name !", vbExclamation, "Add Employee"
Exit Sub
End If


If Option1.Value = False And Option2.Value = False Then
MsgBox "Please Select Gender !", vbExclamation, "Add Employee"
Exit Sub
End If


If Text6.Text = "" Then
MsgBox "Please Enter Basic Pay !", vbExclamation, "Add Employee"
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
MsgBox "Please Enter qualification !", vbExclamation, "Add Employee"
Exit Sub
End If









If Option1.Value = True Then

Call connect

str1 = "insert into emp(emp_id,name,,mobile_number,email_id,gender,qualification,basic_pay,user_name,password) values (" & Text1.Text & " ,' " & Text2.Text & " ',' " & Text3.Text & " ' ,' " & Text4.Text & ",'" & Option1.Caption & "','" & Text5.Text & "','" & Text6.Text & "',' " & Text7.Text & "','" & Text8.Text & "')"

con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save Employee Details"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""



Picture1.Picture = Nothing


Option1.Value = False

'Option2.Value = False



con.Close

End If

If Option2.Value = True Then

Call connect

str1 = "insert into emp(emp_id,name,,mobile_number,email_id,gender,qualification,basic_pay,user_name,password)values (" & Text1.Text & " ,' " & Text2.Text & " ',' " & Text3.Text & " ' ,' " & Text4.Text & ",'" & Option2.Caption & "','" & Text5.Text & "','" & Text6.Text & "',' " & Text7.Text & "','" & Text8.Text & "')"

con.Execute str1
MsgBox "Details saved successfully...", vbInformation, "Save EmployeeDetails"

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""



Option2.Value = False

Picture1.Picture = Nothing

con.Close


End If

''id

Call connect
str1 = "emp"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Dim n As Integer
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("emp_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub


Private Sub text6_Validate(cancel As Boolean)

If Not (IsNumeric(Text6.Text)) Then
MsgBox " Please Enter Basic Pay Numerically !", vbExclamation, "Add Employee"
End If

End Sub

Private Sub text3_Validate(cancel As Boolean)

'mobile

If Not (IsNumeric(Text3.Text)) Then
MsgBox " Please Enter Phone Number Numerically !", vbExclamation, "Add Employee"
End If

If Len(Text3.Text) < 10 Or Len(Text3.Text) > 10 Then
MsgBox "Enter the phone number in 10 digits !", vbExclamation, "Add Employee"
End If

End Sub

