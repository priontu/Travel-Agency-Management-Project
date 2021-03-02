Attribute VB_Name = "Module1"
Public con As Connection
Public rs As Recordset

Public Sub dblink()
Set con = New Connection
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\MAIN PROJECT\travel agency database.accdb"

End Sub

