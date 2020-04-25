Attribute VB_Name = "Module1"

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
