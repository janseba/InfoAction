Option Explicit
Public Function getFileName(sTitle As String) As String

    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = sTitle
        .Show
        If .SelectedItems.Count > 0 Then
            getFileName = .SelectedItems(1)
        Else
            getFileName = ""
        End If
    End With

End Function
Function BrowseFolder(Title As String, _
        Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = _
            msoFileDialogViewList) As String
            
    Dim V As Variant
    Dim InitFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        V = .SelectedItems(1)
        If Err.Number <> 0 Then
            V = vbNullString
        End If
    End With
    BrowseFolder = CStr(V)
    
End Function
Function FileExists(sPath As String) As Boolean
    Dim FilePath As String
    Dim TestStr As String

    FilePath = sPath
    If FilePath = "" Then
        FileExists = False
        Exit Function
    End If

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function
