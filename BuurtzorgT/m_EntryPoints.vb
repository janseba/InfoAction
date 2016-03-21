Option Explicit
Public Sub ribbonClick(control As IRibbonControl)

'On Error GoTo ErrorHandler

    Select Case control.Tag
        Case "getCareweb"
            getCareweb
        Case "maintainSettings"
            maintainSettings
        Case "getAll"
            GetAll
    End Select

ErrorExit:
     Exit Sub
    
ErrorHandler:
    MsgBox "Er heeft zich een onverwachte fout voorgedaan" & _
        Chr(13) & "De volledige omschrijving van de fout is: " & _
        Err.Description, vbCritical, "BuurtzorgT"
    Resume ErrorExit

End Sub

