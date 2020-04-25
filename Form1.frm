VERSION 5.00
Begin VB.Form first 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   18900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   12315
      ScaleWidth      =   22755
      TabIndex        =   0
      Top             =   0
      Width           =   22815
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   14160
         Top             =   9960
         Width           =   5055
      End
   End
End
Attribute VB_Name = "first"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
login.Show
End Sub
