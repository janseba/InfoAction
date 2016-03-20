Option Explicit

Sub getCareweb()
    Dim carewebCsv As String
    carewebCsv = getFileName("Selecteer csv bestand uit Careweb")
    FileCopy carewebCsv, gsDatabasePath & "import.csv"
    
End Sub
Sub importCSV()
    Dim csvConnnection As ADODB.Connection, csvRecords As ADODB.Recordset, dbRecords As ADODB.Recordset, f As Variant
    Set csvConnnection = New ADODB.Connection
    Set csvRecords = New ADODB.Recordset
    csvConnnection.Provider = "Microsoft.Jet.OLEDB.4.0"
    csvConnnection.ConnectionString = "Data Source=" & gsDatabasePath & ";" & _
        "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
    csvConnnection.Open
    csvRecords.Open "SELECT [Hoofdbehandelaar diagnosefase] AS PID FROM [" & gsCSV_FILE & "]", csvConnnection
    csvRecords.MoveFirst
    Set dbRecords = GetRecordSet("SELECT * FROM tblTest WHERE DBCNo IS NULL")
    Do Until csvRecords.EOF
        dbRecords.AddNew
        For Each f In csvRecords.Fields
            dbRecords.Fields(f.Name) = f.Value
        Next f
        csvRecords.MoveNext
    Loop
    UpdateRecords dbRecords
End Sub

Sub maintainSettings()
    Dim frmMaintainSettings As FrmSettings
    Set frmMaintainSettings = New FrmSettings
    With frmMaintainSettings
        .Show
        If .OK Then
            SaveSetting "BuurtzorgT", "Settings", "DBPath", .DBPath
            InitGlobals 'Zorgt er voor dat de nieuwe settings van kracht zijn
        End If
    End With
    Unload frmMaintainSettings
End Sub

