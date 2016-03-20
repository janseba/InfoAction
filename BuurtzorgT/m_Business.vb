Option Explicit

Sub getCareweb()
    Dim FilePath As String
    FilePath = getFileName("Selecteer csv bestand uit Careweb")
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

