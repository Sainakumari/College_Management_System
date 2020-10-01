Attribute VB_Name = "Module1"
Public c As ADODB.Connection
Public r As ADODB.Recordset
Public str As String
Public sql As String
Public Sub conn()
Set c = New ADODB.Connection
c.Open "provider=msdaora.1;user id=cms/saina;persist security info=true"
Set r = New ADODB.Recordset
End Sub

