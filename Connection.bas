Attribute VB_Name = "Module1"
Option Explicit

Public c As adodb.Connection
Public r As adodb.Recordset
Public sql As String
Public Function conn()
Set c = New adodb.Connection
c.Open "provider=msdaora.1; user id=moon/admin; persist security info =false"
Set r = New adodb.Recordset
End Function
