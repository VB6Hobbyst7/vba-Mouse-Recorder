Attribute VB_Name = "Dialogues"
Function PickRecord(Optional initFolder As String) As String
    If initFolder = "" Then initFolder = Environ("USERprofile") & "\Desktop\"
    Dim strFile As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "MouseRecord", "*.txt"
        .Title = "Choose Mouse Record"
        .AllowMultiSelect = False
        .InitialFileName = initFolder
        If .Show = True Then
            strFile = .SelectedItems(1)
            PickRecord = strFile
        End If
    End With
End Function

Public Function SelectFolder(Optional initFolder As String) As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder"
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems.Item(1)
        Else
            'MsgBox "Folder is not selected."
        End If
    End With
    
End Function


Function InputboxString(Optional sTitle As String = "Select String", Optional sPrompt As String = "Select String") As String
    Dim stringVariable As String
    stringVariable = Application.InputBox( _
                     Title:=sTitle, _
                     Prompt:=sPrompt, _
                     Type:=2)
    InputboxString = CStr(stringVariable)
End Function

