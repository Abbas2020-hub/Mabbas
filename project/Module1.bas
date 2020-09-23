Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection

Public Function connect()
conn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=."
End Function


Public Sub Main()
Call Splash.Show
Splash.SetFocus
End Sub
