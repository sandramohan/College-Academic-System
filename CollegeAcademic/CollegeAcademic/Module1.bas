Attribute VB_Name = "Module1"

Public con As New ADODB.Connection
Public cmd As New ADODB.Command
Public usr, id As String
Public Sub getcon()
If con.State = 1 Then con.Close
con.ConnectionString = "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=sa;Initial Catalog=Academic;Data Source=."
con.Open
cmd.ActiveConnection = con
End Sub



