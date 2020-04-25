VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form search_vehicle 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "update_vehicle.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   15840
      TabIndex        =   8
      Top             =   6360
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   15840
      TabIndex        =   7
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   15840
      TabIndex        =   6
      Top             =   3840
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   9960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   121962497
      CurrentDate     =   43414
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   8760
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   7560
      Width           =   4335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Text            =   "SELECT VEHICLE TYPE"
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   6240
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Text            =   "SELECT VEHICLE ID"
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Image clear 
      Height          =   975
      Left            =   15120
      Top             =   9840
      Width           =   4095
   End
   Begin VB.Image delete 
      Height          =   975
      Left            =   10560
      Top             =   9840
      Width           =   4095
   End
   Begin VB.Image cancel 
      Height          =   975
      Left            =   15120
      Top             =   8280
      Width           =   4215
   End
   Begin VB.Image update 
      Height          =   975
      Left            =   10560
      Top             =   8280
      Width           =   4095
   End
End
Attribute VB_Name = "search_vehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Combo1_click()
Call connect
str1 = "select * from vehicle where vehicle_id=" & Combo1.Text & ""
rs.Open str1, con, adOpenKeyset

Text1.Text = rs.Fields("fuel_type").Value
Combo2.Text = rs.Fields("vehicle_type").Value
Text2.Text = rs.Fields("company_name").Value
Text3.Text = rs.Fields("vehicle_model").Value
DTPicker1.Value = rs.Fields("registration_date").Value
Text4.Text = rs.Fields("vehicle_number").Value
Text5.Text = rs.Fields("total_seats").Value
Text6.Text = rs.Fields("fuel_capacity").Value
rs.MoveNext



rs.Close
con.Close
End Sub



Private Sub Form_Load()
If con.State = adStateOpen Then
rs.Close
con.Close
End If

With Combo2
    

    .AddItem "car"
    .AddItem "bus"
    
    End With


Call connect
str1 = "select * from vehicle"
rs.Open str1, con, adOpenKeyset

Do While Not rs.EOF

Combo1.AddItem (rs.Fields("vehicle_id").Value)
rs.MoveNext
Loop
rs.Close
con.Close
End Sub

Private Sub cancel_Click()
Unload Me
End Sub

 
Private Sub update_Click()


If Combo1.Text = "SELECT VEHICLE ID" Then
MsgBox "Select Vehicle ID to UPDATE", vbExclamation, "Delete vehicle"
Exit Sub
End If



If Combo2.Text = "SELECT VEHICLE TYPE" Then
MsgBox "Please select vehicle type State !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text1.Text = "" Then
MsgBox "Please Enter Fuel Type !", vbExclamation, "Add Vehicle"
Exit Sub
End If


If Text2.Text = "" Then
MsgBox "Please Enter Company Name !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text3.Text = "" Then
MsgBox "Please Enter Vehicle Model !", vbExclamation, "Add Vehicle"
Exit Sub
End If

If Text4.Text = "" Then
MsgBox "Please Enter Vehicle Number !", vbExclamation, "Add Vehicle"
Exit Sub
End If



If Text5.Text = "" Then
MsgBox "Please Enter Total Seats !", vbExclamation, "Add vehicle"
Exit Sub
End If


If Text6.Text = "" Then
MsgBox "Please Enter Fuel Capacity !", vbExclamation, "Add Vehicle"
Exit Sub
End If


Call connect
str1 = "update vehicle set vehicle_type='" & Combo2.Text & "',fuel_type='" & Text1.Text & "',company_name='" & Text2.Text & "',vehicle_model='" & Text3.Text & "',registration_date='" & DTPicker1.Value & "',vehicle_number='" & Text4.Text & "',total_seats=" & Text5.Text & ",fuel_capacity='" & Text6.Text & "'where vehicle_id=" & Combo1.Text & " "
con.Execute str1
MsgBox "Vehicle details Updated Successfully", vbInformation, "Update vehicle"
con.Close

End Sub

Private Sub delete_Click()

If Combo1.Text = "" Then
MsgBox "Select Vehicle ID to delete", vbExclamation, "Delete vehicle"
Exit Sub
End If


Call connect

confirm = MsgBox("Are you sure you want to delete this vehicle ?", vbYesNo, "Delete a vehicle")
If confirm = vbYes Then
str1 = "delete from vehicle where vehicle_id=" & Combo1.Text & ""
con.Execute str1
MsgBox "Vehicle Deleted Successfully...", vbInformation, "Delete vehicle"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Text = ""
Combo1.Text = ""

Else
MsgBox "Vehicle not deleted ", vbInformation, "Delete Vehicle"
End If

con.Close

End Sub

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
End Sub
