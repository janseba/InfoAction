Option Explicit
Private mcnConnection As ADODB.Connection
Function GetRecordSet(ByRef sSQL As String) As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    
    If mcnConnection Is Nothing Then bCreateDBConnection gsDatabasePath & gsDATABASE_FILE
    mcnConnection.Open
    
    Set rsData = New ADODB.Recordset
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open sSQL, mcnConnection, , , adCmdText
    End With
    Set rsData.ActiveConnection = Nothing
    mcnConnection.Close
    Set GetRecordSet = rsData
End Function
Sub UpdateRecords(ByRef rsData As ADODB.Recordset)
    If mcnConnection Is Nothing Then bCreateDBConnection gsDatabasePath & gsDATABASE_FILE
    mcnConnection.Open
    rsData.ActiveConnection = mcnConnection
    rsData.UpdateBatch
    Set rsData.ActiveConnection = Nothing
    mcnConnection.Close
End Sub
Public Function bCreateDBConnection(ByRef sFullName As String) As Boolean
    Dim bReturn As Boolean, sConnect As String
    
    On Error GoTo ErrorHandler
    bReturn = True
    sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sFullName & ";"
    Set mcnConnection = New ADODB.Connection
    mcnConnection.ConnectionString = sConnect
    mcnConnection.Open
    mcnConnection.Close
    
ErrorExit:
    bCreateDBConnection = bReturn
    Exit Function
    
ErrorHandler:
    bReturn = False
    Resume ErrorExit
End Function
Public Sub RunSQL(ByVal sql As String)
    If mcnConnection Is Nothing Then bCreateDBConnection gsDatabasePath & gsDATABASE_FILE
    mcnConnection.Open
    mcnConnection.Execute sql, , adExecuteNoRecords
    mcnConnection.Close
End Sub
