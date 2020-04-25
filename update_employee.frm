VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form update_employee 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   2010
   ClientTop       =   1485
   ClientWidth     =   19815
   LinkTopic       =   "Form1"
   Picture         =   "update_employee.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   19815
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   615
      Left            =   4440
      TabIndex        =   17
      Top             =   6720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   127008769
      CurrentDate     =   43346
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   4440
      TabIndex        =   16
      Top             =   4800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   127008769
      CurrentDate     =   43346
   End
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   14160
      TabIndex        =   15
      Top             =   9480
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16680
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   17520
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Height          =   615
      Left            =   14160
      TabIndex        =   13
      Top             =   8520
      Width           =   4935
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   14160
      TabIndex        =   12
      Top             =   7560
      Width           =   4935
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   14160
      TabIndex        =   11
      Top             =   6600
      Width           =   4935
   End
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   14160
      TabIndex        =   10
      Top             =   5640
      Width           =   4935
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   14160
      TabIndex        =   9
      Top             =   4560
      Width           =   4935
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   14160
      TabIndex        =   8
      Top             =   3600
      Width           =   4935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FEMALE"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   5880
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   9600
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   8640
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   7680
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Text            =   "SELECT EMPLOYEE ID"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   840
      Top             =   10560
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   16920
      Top             =   10560
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   17640
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   8880
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "update_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If


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

Private Sub Image2_Click()
CommonDialog1.ShowOpen
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image1_Click()

Call connect
str1 = "select * from employee_management where employee_id=" & Combo1.Text & ""

rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Enabled = False
DTPicker1.Value = rs.Fields("date_of_birth").Value


If rs.Fields(3) = Option1.Caption Then
Option1.Value = True
Else
Option2.Value = True
End If

DTPicker2.Value = rs.Fields("date_of_join").Value



Text1.Text = rs.Fields("employee_id").Value
Text2.Text = rs.Fields("employee_name").Value
Text3.Text = rs.Fields("mobile").Value
Text4.Text = rs.Fields("email_id").Value
Text5.Text = rs.Fields("basic_pay").Value
Text6.Text = rs.Fields("house_name").Value
Text7.Text = rs.Fields("village").Value
Text8.Text = rs.Fields("city").Value
Text9.Text = rs.Fields("town").Value
Text10.Text = rs.Fields("pin_code").Value
Text11.Text = rs.Fields("states").Value
Text12.Text = rs.Fields("country").Value
Picture1.Picture = LoadPicture(rs.Fields("pic"))


str = rs.Fields("pic")
Picture1.Picture = LoadPicture(str)
rs.MoveNext
Loop
rs.Close
con.Close


End Sub

Private Sub Image4_Click()

''male

If Option1.Value = True Then
Call connect
str1 = "update employee_management set employee_name= '" & Text2.Text & "',date_of_birth='" & DTPicker1.Value & "',gender= '" & Option1.Caption & "',date_of_join='" & DTPicker2.Value & "',mobile=" & Text3.Text & ",email_id='" & Text4.Text & "',basic_pay=" & Text5.Text & ",house_name='" & Text6.Text & "',village='" & Text7.Text & "',city='" & Text8.Text & "',town='" & Text9.Text & "',pin_code='" & Text10.Text & "',states='" & Text11.Text & "',country='" & Text12.Text & "',pic ='" & str & "' where employee_id=" & Combo1.Text & " "

con.Execute str1
MsgBox "Employee Details Updated Successfully...", vbInformation, "Update Employee Details"

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


con.Close

End If

''female

If Option2.Value = True Then
Call connect
str1 = "update employee_management set employee_name= '" & Text2.Text & "',date_of_birth='" & DTPicker1.Value & "',gender= '" & Option2.Caption & "',date_of_join='" & DTPicker2.Value & "',mobile=" & Text3.Text & ",email_id='" & Text4.Text & "',basic_pay=" & Text5.Text & ",house_name='" & Text6.Text & "',village='" & Text7.Text & "',city='" & Text8.Text & "',town='" & Text9.Text & "',pin_code='" & Text10.Text & "',states='" & Text11.Text & "',country='" & Text12.Text & "',pic ='" & str & "' where employee_id=" & Combo1.Text & " "

con.Execute str1
MsgBox "Employee Details Updated Successfully...", vbInformation, "Update Employee Details"


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
con.Close

End If
End Sub
