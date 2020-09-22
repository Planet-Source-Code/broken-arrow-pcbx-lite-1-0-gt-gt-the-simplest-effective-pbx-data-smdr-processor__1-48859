Attribute VB_Name = "modADO"
Option Explicit

Public GlobalADOConnectionString As String, MDBDatabase As String

Public Function Query(SQL As String, Optional ConnectionString As String = "") As Recordset
On Error GoTo Query_Error

Set Query = New ADODB.Recordset
Query.Open SQL, ADOConnection, adOpenStatic, adLockOptimistic
Query.MoveLast
Query.MoveFirst

Query_Exit: Exit Function

Query_Error:
Select Case Err.Number
Case 3021 'No current record!
    Resume Next
Case Else
    MsgBox "Error#" & Err.Number & ": " & Err.Description
    Resume Next
End Select
End Function

Public Function QueryExec(SQL As String, Optional ConnectionString As String)
Dim CMD As New ADODB.Command
CMD.ActiveConnection = ADOConnection(ConnectionString)
CMD.CommandText = SQL
CMD.Execute
End Function

Private Function ADOConnection(Optional ConnectionString As String = "") As ADODB.Connection
Set ADOConnection = New ADODB.Connection
If ConnectionString = "" Then ConnectionString = GlobalADOConnectionString
ADOConnection.Open ConnectionString
End Function
