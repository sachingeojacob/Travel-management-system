VERSION 5.00
Begin VB.Form login 
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   2280
   ClientTop       =   795
   ClientWidth     =   19530
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   19530
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   12375
      Left            =   -240
      Picture         =   "login.frx":0000
      ScaleHeight     =   12315
      ScaleWidth      =   22995
      TabIndex        =   0
      Top             =   -600
      Width           =   23055
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000011&
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         MaskColor       =   &H00808080&
         TabIndex        =   4
         Top             =   7440
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000B&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Rockwell Extra Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   800
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   3
         Top             =   7440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "#"
         TabIndex        =   2
         Top             =   6000
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   2880
         TabIndex        =   1
         Top             =   4560
         Width           =   3735
      End
   End
End
Attribute VB_Name = "login"
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
Public SQL As String
Public Sub connect()
con.Provider = "SQLOLEDB"
str1 = "server=DESKTOP-HOTR91D;database=master;trusted_connection=yes"
con.Open str1
End Sub



Private Sub command1_Click()
Call connect
str1 = "select * from login_page"
rs.Open str1, con, adOpenKeyset
status = False
Do
If rs.EOF Then
MsgBox "invalid login , user name and password are not correct"
con.Close
End If
If rs("user_name").Value = Text1.Text And rs("password").Value = Text2.Text Then
status = True
t = rs("account_type").Value
Exit Do
Else
rs.MoveNext
End If
Loop Until rs.EOF
If status = True And t = "ADMIN" Then
MDIForm1.Show
Else
If status = True And t = "EMPLOYEE" Then
MDIForm1.Show
MDIForm1.ADDEMPLOYEE.Visible = False
MDIForm1.DELETEEMPLOYEE.Visible = False
MDIForm1.UPDATEEMPLOYEE.Visible = False
MDIForm1.SEARCHEMPLOYEE.Visible = False
MDIForm1.EMPLOYEEMANAGEMENT.Visible = False
MDIForm1.SALARYMANAGEMENT.Visible = False
MDIForm1.ADDNOTIFICATION.Visible = False

MDIForm1.CREATEACCOUNT.Visible = False
MDIForm1.DELETEACCOUNT.Visible = False
MDIForm1.DELETENOTIFICATION.Visible = False

MDIForm1.UPDATENOTIFICATION.Visible = False




Else
MsgBox "invalid username or password"
End If
End If
con.Close
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

