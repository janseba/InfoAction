Option Explicit
Dim mbOK As Boolean
Private Sub UserForm_Initialize()
    
    Me.txtDBLocation = gsDatabasePath

End Sub
Private Sub cmdBrowseDBFolder_Click()

    Dim sPath As String
    sPath = BrowseFolder("Selecteer locatie database")
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    Me.txtDBLocation = sPath
    
End Sub
Private Sub cmdCancel_Click()
    
    mbOK = False
    Me.Hide
    
End Sub
Private Sub cmdOK_Click()

    If CheckInput() Then
        mbOK = True
        Me.Hide
    End If
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click
        Cancel = True
    End If
    
End Sub
Public Property Get OK() As Boolean
    OK = mbOK
End Property
Public Property Get DBPath() As String
    DBPath = Me.txtDBLocation.Value
End Property
Function CheckInput() As Boolean
    Dim bInputOK As Boolean
    
    bInputOK = True
        
    If Not FileExists(Me.txtDBLocation.Value) Then
        bInputOK = False
        MsgBox "De map " & Me.txtDBLocation.Value & " is niet gevonden.", vbExclamation
    End If
    
    CheckInput = bInputOK
    
End Function


