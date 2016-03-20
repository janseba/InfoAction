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
    csvRecords.Open "SELECT [DBCNo],[CliÌÇnt ID] AS ClientID, [Hoofdbehandelaar diagnosefase] AS PID,[Behandelende eenheid] AS BehandelendeEenheid FROM [" & gsCSV_FILE & "]", csvConnnection
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

