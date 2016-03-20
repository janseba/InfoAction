Option Explicit

Public gsDatabasePath As String

Public Sub InitGlobals()
    
    gsDatabasePath = GetSetting("BuurtzorgT", "Settings", "DBPath")
    If Right$(gsDatabasePath, 1) <> "\" Then gsDatabasePath = gsDatabasePath & "\"
    If Not FileExists(gsDatabasePath) Then
        MsgBox "De map " & gsDatabasePath & " bestaat niet. Pas dit aan in de instellingen.", vbExclamation
    End If
    
End Sub

