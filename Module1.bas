Attribute VB_Name = "Module1"
Global CnnStr As ADODB.Connection
Global un As String
Sub main()
Set CnnStr = New ADODB.Connection
CnnStr.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\login.mdb;Persist Security Info=False"
splash.Show
End Sub

Public Sub OpenRecordSet(ByRef rsSend As ADODB.Recordset, Optional sSQL As String)
Set rsSend = New ADODB.Recordset
rsSend.ActiveConnection = CnnStr
rsSend.CursorLocation = adUseClient
rsSend.CursorType = adOpenKeyset
rsSend.LockType = adLockPessimistic

If sSQL <> "" Then
rsSend.Open sSQL
End If
End Sub
