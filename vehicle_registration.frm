VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vehicle_registration 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "vehicle_registration.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   15840
      TabIndex        =   8
      Top             =   6360
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   15840
      TabIndex        =   7
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   15840
      TabIndex        =   6
      Top             =   3720
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   9960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   121962497
      CurrentDate     =   43414
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   8760
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   7560
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   6240
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      Text            =   "SELECT VEHICLE TYPE"
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Image cancel 
      Height          =   1215
      Left            =   15480
      Top             =   8760
      Width           =   4095
   End
   Begin VB.Image register 
      Height          =   1215
      Left            =   10800
      Top             =   8760
      Width           =   4095
   End
End
Attribute VB_Name = "vehicle_registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()

If con.State = adStateOpen Then
rs.Close
con.Close
End If

With Combo1
    

    .AddItem "car"
    .AddItem "bus"
    
    End With
    
Call connect
str1 = "select * from vehicle"
rs.Open str1, con, adOpenKeyset
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("vehicle_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub



Private Sub register_Click()



If Combo1.Text = "SELECT VEHICLE TYPE" Then
MsgBox "Please select vehicle type State !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text2.Text = "" Then
MsgBox "Please Enter Fuel Type !", vbExclamation, "Add Vehicle"
Exit Sub
End If


If Text3.Text = "" Then
MsgBox "Please Enter Company Name !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text4.Text = "" Then
MsgBox "Please Enter Vehicle Model !", vbExclamation, "Add Vehicle"
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Please Enter Vehicle Number !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text6.Text = "" Then
MsgBox "Please Enter Total Seats !", vbExclamation, "Add vehicle"
Exit Sub
End If


If Text7.Text = "" Then
MsgBox "Please Enter Fuel Capacity !", vbExclamation, "Add Vehicle"
Exit Sub
End If


Call connect

str1 = "insert into vehicle(vehicle_id,vehicle_type,fuel_type,company_name,vehicle_model,registration_date,vehicle_number,total_seats,fuel_capacity)values (" & Text1.Text & ",'" & Combo1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & DTPicker1.Value & "',' " & Text5.Text & "'," & Text6.Text & ",' " & Text7.Text & "')"
con.Execute str1
MsgBox "Vehicle registered successfully", vbInformation, "Add vehicle"



Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""

con.Close

Call connect
str1 = "select * from vehicle"
rs.Open str1, con, adOpenKeyset
If rs.RecordCount = 0 Then
Text1.Text = "1"
Else
rs.MoveLast
n = rs("vehicle_id").Value
Text1.Text = n + 1
End If
rs.Close
con.Close
End Sub



Private Sub cancel_Click()
Unload Me
End Sub


