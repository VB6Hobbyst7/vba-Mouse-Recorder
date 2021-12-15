VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uMouseRecorder 
   Caption         =   "Mouse Macro"
   ClientHeight    =   5496
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2112
   OleObjectBlob   =   "uMouseRecorder.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uMouseRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub crearREC_Click()
newRecord
End Sub

Private Sub dragRec_Click()
recordDrag
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    uDEV.Show
End Sub

Private Sub RecordLoad_Click()
    LoadRecord
    Me.LoadedRecording.ControlTipText = Me.LoadedRecording.Caption
End Sub

Private Sub PickRecordFolder_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim recFolder As String
    recFolder = Dialogues.SelectFolder
    If recFolder <> "" Then
        Me.LoadedFolder.Caption = IIf(FolderExists(recFolder) = True, recFolder, "NONE")
        LoadedFolder.ControlTipText = Me.LoadedFolder.Caption
        ws.Range("recFolder") = recFolder
    End If
End Sub

Private Sub RecordSave_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    If Not IsFolderSelected Then
        MsgBox "Pick a folder for your recordings first"
        Exit Sub
    End If
    If IsRecordSaved Then
        MsgBox "Record already saved"
        Exit Sub
    End If
    If ws.Range("A3") = "" Then
        MsgBox "Record something first"
        Exit Sub
    End If

    Dim fName As String
    fName = InputboxString
    If Len(fName) <> 0 And fName <> "False" Then
        ws.Range("recFile") = fName
        saveRecord
    End If
End Sub

Private Sub recWholeMotion_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    ws.Range("recWholeMotion") = recWholeMotion
    wholeMotionCaption
End Sub

Sub wholeMotionCaption()
    If recWholeMotion = True Then
        recWholeMotion.Caption = "Record whole motion"
    Else
        recWholeMotion.Caption = "Record clicks only"
    End If
End Sub

Private Sub ReplayRecord_Click()
    Mouse.MouseReplay
End Sub

Private Sub StartRecord_Click()
    Mouse.RecordStart Me.recWholeMotion.Value
    checkFile
End Sub

Private Sub UserForm_Initialize()
    LoadUserformPosition
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim recFolder As String
    recFolder = ws.Range("recFolder").Text
    Me.LoadedFolder.Caption = IIf(FolderExists(recFolder) = True, recFolder, "NONE")
    checkFile
    recWholeMotion = ws.Range("recWholeMotion")
    wholeMotionCaption
    Me.Show
    flashControl Me.info

End Sub

Sub checkFile()
    Dim recFile As String
    recFile = RecordFileFullName
    Me.LoadedRecording.Caption = IIf(FileExists(recFile) = True, recFile, "NONE")
    Me.LoadedRecording.ControlTipText = Me.LoadedRecording.Caption
    Me.LoadedFolder.ControlTipText = Me.LoadedFolder.Caption
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveUserformPosition
End Sub

Sub LoadUserformPosition()
    '/load position
    If GetSetting("My Settings Folder", Me.Name, "Left Position") = "" _
                                                                    And GetSetting("My Settings Folder", Me.Name, "Top Position") = "" Then
        Me.StartUpPosition = 1        ' CenterOwner
    Else
        Me.Left = GetSetting("My Settings Folder", Me.Name, "Left Position")
        Me.Top = GetSetting("My Settings Folder", Me.Name, "Top Position")
    End If
    'load position/
End Sub

Sub SaveUserformPosition()
    '/save position
    'must have uf position set to manual
    SaveSetting "My Settings Folder", Me.Name, "Left Position", Me.Left
    SaveSetting "My Settings Folder", Me.Name, "Top Position", Me.Top
End Sub


