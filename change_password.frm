VERSION 5.00
Begin VB.Form change_password 
   Caption         =   "Form1"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19290
   LinkTopic       =   "Form1"
   Picture         =   "change_password.frx":0000
   ScaleHeight     =   10080
   ScaleWidth      =   19290
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   9000
      TabIndex        =   3
      Top             =   7920
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   9000
      TabIndex        =   2
      Top             =   6360
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   9000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4680
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9000
      TabIndex        =   0
      Text            =   "SELECT USERNAME"
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   10920
      Top             =   9720
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   3960
      Top             =   9840
      Width           =   4575
   End
End
Attribute VB_Name = "change_password"
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
Public Sub connect()
con.Provider = "sqloledb"
str1 = "server=DESKTOP-HOTR91D\SQLEXPRESS;database=master;trusted_connection=yes"
con.Open str1

End Sub
Private Sub Combo1_Click()

Call connect
str1 = "select user_name,password from login_page where user_name='" & Combo1.Text & "'"
rs.Open str1, con, adOpenKeyset
Do While Not rs.EOF
Text1.Text = rs.Fields("password").Value
rs.MoveNext
Loop
rs.Close
con.Close





End Sub




Private Sub Form_Load()
Dim str1 As String
Call connect
str1 = "select user_name from login_page"
rs.Open str1, con, adOpenKeyset
con.Execute str1
Do While Not rs.EOF
Combo1.AddItem (rs.Fields("user_name").Value)
rs.MoveNext
Loop
rs.Close
con.Close



End Sub

Private Sub Image1_Click()
Call connect
If Text2.Text = Text3.Text Then

str1 = "update login_page set password='" & Text2.Text & "' where user_name= '" & Combo1.Text & "' "
con.Execute str1
MsgBox "successfully updated"
Else
MsgBox "current password is incorrect"

End If
con.Close
Combo1.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub




Private Sub Image2_Click()
Unload Me
End Sub
