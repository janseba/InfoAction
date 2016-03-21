Option Explicit
Private mcurrentRS As ADODB.Recordset
Sub getCareweb()
    Dim carewebCsv As String
    carewebCsv = getFileName("Selecteer csv bestand uit Careweb")
    FileCopy carewebCsv, gsDatabasePath & "import.csv"
    importCSV
    RunSQL "INSERT INTO tblDBCMutatie (DBCNo) SELECT a.DBCNo FROM tblDBC AS a WHERE a.DBCNo NOT IN (SELECT DBCNo FROM tblDBCMutatie)"
    MsgBox "De import is gereed.", vbInformation
End Sub
Sub importCSV()
    Dim csvConnnection As ADODB.Connection, csvRecords As ADODB.Recordset, dbRecords As ADODB.Recordset, f As Variant
    Set csvConnnection = New ADODB.Connection
    Set csvRecords = New ADODB.Recordset
    csvConnnection.Provider = "Microsoft.Jet.OLEDB.4.0"
    csvConnnection.ConnectionString = "Data Source=" & gsDatabasePath & ";" & _
        "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
    csvConnnection.Open
    csvRecords.Open "SELECT [DBCNo],[CliÌÇnt ID] AS ClientID, [Hoofdbehandelaar diagnosefase] AS PID" & _
        ",[Behandelende eenheid] AS BehandelendeEenheid, Financier,[Startdatum DBC] AS StartdatumDBC," & _
        "[Einddatum DBC] AS EinddatumDBC,[Berekend bedrag huidige registratie] AS BerekendBedragDeclaratie," & _
        "[Laatste declaratiedocument: Detailstatus] AS LaatsteDeclaratieDetailStatus,Onderwerp," & _
        "IIF([Uitgezonderd van automatische decl#]=1,'Ja','Nee') AS UitgezonderdAutomatischeDeclaratie FROM [" & gsCSV_FILE & "]", csvConnnection
    csvRecords.MoveFirst
    Set dbRecords = GetRecordSet("SELECT * FROM tblDBC WHERE DBCNo IS NULL")
    Do Until csvRecords.EOF
        dbRecords.AddNew
        For Each f In csvRecords.Fields
            dbRecords.Fields(f.Name) = f.Value
        Next f
        csvRecords.MoveNext
    Loop
    RunSQL "DELETE FROM tblDBC"
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
Sub OpenAddInWorkbook()
    Dim wkbBook As Workbook
    On Error Resume Next
    Set wkbBook = Application.Workbooks(gsDatabasePath & gsADD_IN_WORKBOOK)
    On Error GoTo 0
    If wkbBook Is Nothing Then
        Set wkbBook = Application.Workbooks.Open(Filename:=gsDatabasePath & gsADD_IN_WORKBOOK, ReadOnly:=True)
    Else
        wkbBook.Activate
    End If
End Sub
Sub GetAll()
    Set mcurrentRS = GetRecordSet("SELECT * FROM View_DBC ORDER BY a.DBCNo")
    mcurrentRS.MoveFirst
    OpenAddInWorkbook
    updateSingleView
End Sub
Sub updateSingleView()
    Dim n As Name, pos As String, last As String
    Application.EnableEvents = False
    pos = mcurrentRS.AbsolutePosition
    last = mcurrentRS.RecordCount
    For Each n In ActiveWorkbook.Names
        If Left(n.Name, 3) = "db_" Then
            n.RefersToRange.Value = mcurrentRS.Fields(Mid(n.Name, 4, 99))
        End If
    Next n
    With ActiveSheet
        .Shapes("recordCounter").TextFrame.Characters.Text = pos & " van " & last
        .Shapes("cmdFirst").OnAction = "cmdFirst"
        .Shapes("cmdPrevious").OnAction = "cmdPrevious"
        .Shapes("cmdNext").OnAction = "cmdNext"
        .Shapes("cmdLast").OnAction = "cmdLast"
        If pos = 1 Then
            .Shapes("cmdPrevious").DrawingObject.Font.ColorIndex = 16
        Else
            .Shapes("cmdPrevious").DrawingObject.Font.ColorIndex = 1
        End If
        If pos = mcurrentRS.RecordCount Then
            .Shapes("cmdNext").DrawingObject.Font.ColorIndex = 16
        Else
            .Shapes("cmdNext").DrawingObject.Font.ColorIndex = 1
        End If
    End With
    Application.EnableEvents = True
End Sub
Sub cmdFirst()
    mcurrentRS.MoveFirst
    updateSingleView
End Sub
Sub cmdPrevious()
    If mcurrentRS.AbsolutePosition > 1 Then
        mcurrentRS.MovePrevious
        updateSingleView
    End If
End Sub
Sub cmdNext()
    If mcurrentRS.AbsolutePosition < mcurrentRS.RecordCount Then
        mcurrentRS.MoveNext
        updateSingleView
    End If
End Sub
Sub cmdLast()
    mcurrentRS.MoveLast
    updateSingleView
End Sub
Public Sub processChange(ByVal rng As Range)
    With ActiveSheet
        If rng.Address = .Range("db_AkteVanCessie").Address Then
            mcurrentRS.Fields("AkteVanCessie").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_Verwijsbrief").Address Then
            mcurrentRS.Fields("Verwijsbrief").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_Gedeclareerd").Address Then
            mcurrentRS.Fields("Gedeclareerd").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_Declaratiedatum").Address Then
            mcurrentRS.Fields("Declaratiedatum").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_Factuurnummer").Address Then
            mcurrentRS.Fields("Factuurnummer").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_DatumOntvangst").Address Then
            mcurrentRS.Fields("DatumOntvangst").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_OntvangenBedrag").Address Then
            mcurrentRS.Fields("OntvangenBedrag").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_OntvangstGecontroleerd").Address Then
            mcurrentRS.Fields("OntvangstGecontroleerd").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_Afgeletterd").Address Then
            mcurrentRS.Fields("Afgeletterd").Value = rng.Value
            UpdateRecords mcurrentRS
        ElseIf rng.Address = .Range("db_OpgenomenInBoekhouding").Address Then
            mcurrentRS.Fields("OpgenomenInBoekhouding").Value = rng.Value
            UpdateRecords mcurrentRS
        End If
    End With
End Sub

