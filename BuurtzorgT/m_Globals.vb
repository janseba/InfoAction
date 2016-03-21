Option Explicit

Public gsDatabasePath As String
Public Const gsDATABASE_FILE As String = "declaraties.accdb"
Public Const gsCSV_FILE As String = "import.csv"
Public Const gsADD_IN_WORKBOOK = "declaraties.xlsm"

Public Sub InitGlobals()
    
    gsDatabasePath = GetSetting("BuurtzorgT", "Settings", "DBPath")
    If Right$(gsDatabasePath, 1) <> "\" Then gsDatabasePath = gsDatabasePath & "\"
    If Not FileExists(gsDatabasePath) Then
        MsgBox "De map " & gsDatabasePath & " bestaat niet. Pas dit aan in de instellingen.", vbExclamation
    End If
    
End Sub

