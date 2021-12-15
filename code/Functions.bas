Attribute VB_Name = "Functions"
Private Declare PtrSafe Function GetTickCount Lib "Kernel32" () As Long
Private Const Black As Long = &H80000012
Private Const Red As Long = &HFF&

Sub flashControl(ctr As MSForms.Control)
    Dim lngTime As Long
    Dim i As Integer
    For i = 1 To 20
        lngTime = GetTickCount
        If ctr.Visible = True Then
            ctr.Visible = False
        Else
            ctr.Visible = True
        End If
        DoEvents
        Do While GetTickCount - lngTime < 200
        Loop
    Next
End Sub
' credited to ndu
Function Filter2DArray(ByVal sArray, ByVal ColIndex As Long, ByVal FindStr As String, ByVal HasTitle As Boolean)
  Dim tmpArr, i As Long, j As Long, arr, Dic, TmpStr, Tmp, Chk As Boolean, TmpVal As Double
  On Error Resume Next
  Set Dic = CreateObject("Scripting.Dictionary")
  tmpArr = sArray
  ColIndex = ColIndex + LBound(tmpArr, 2) - 1
  Chk = (InStr("><=", Left(FindStr, 1)) > 0)
  For i = LBound(tmpArr, 1) - HasTitle To UBound(tmpArr, 1)
    If Chk Then
      TmpVal = CDbl(tmpArr(i, ColIndex))
      If Evaluate(TmpVal & FindStr) Then Dic.Add i, ""
    Else
      If UCase(tmpArr(i, ColIndex)) Like UCase(FindStr) Then Dic.Add i, "" 'This finds only exact matches, if you need *FindStr* use:  If UCase(tmpArr(i, ColIndex)) Like UCase("*" & FindStr & "*") Then Dic.Add i, ""

    End If
  Next
  If Dic.Count > 0 Then
    Tmp = Dic.Keys
    ReDim arr(LBound(tmpArr, 1) To UBound(Tmp) + LBound(tmpArr, 1) - HasTitle, LBound(tmpArr, 2) To UBound(tmpArr, 2))
    For i = LBound(tmpArr, 1) - HasTitle To UBound(Tmp) + LBound(tmpArr, 1) - HasTitle
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(i, j) = tmpArr(Tmp(i - LBound(tmpArr, 1) + HasTitle), j)
      Next
    Next
    If HasTitle Then
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(LBound(tmpArr, 1), j) = tmpArr(LBound(tmpArr, 1), j)
      Next
    End If
  End If
  Filter2DArray = arr
End Function
Function FolderExists(ByVal Folder As String) As Boolean
    If Right(Folder, 1) = "\" Then Folder = Left(Folder, Len(Folder) - 1)
    FolderExists = (Dir(Folder, vbDirectory + vbArchive + vbHidden + vbSystem) <> "") And (Dir(Folder, vbArchive + vbHidden + vbSystem) = "")
End Function

Public Function FileExists(ByVal filename As String) As Boolean
    If InStr(1, filename, "\") = 0 Then Exit Function
    If Right(filename, 1) = "\" Then filename = Left(filename, Len(filename) - 1)
    FileExists = (Dir(filename, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

Function IsFolderSelected() As Boolean
    IsFolderSelected = Not uMouseRecorder.LoadedFolder.Caption = "NONE"
End Function


' Enum MouseButtonConstants
' vbLeftButton
' vbMiddleButton
' vbRightButton
' End Enum
'
''simulate the MouseDown event
' Sub ButtonDown(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    Dim lFlag As Long
'    If Button = vbLeftButton Then
'        lFlag = MOUSEEVENTF_LEFTDOWN
'    ElseIf Button = vbMiddleButton Then
'        lFlag = MOUSEEVENTF_MIDDLEDOWN
'    ElseIf Button = vbRightButton Then
'        lFlag = MOUSEEVENTF_RIGHTDOWN
'    End If
'    mouse_event lFlag, 0, 0, 0, 0
'End Sub
'
''simulate the MouseUp event
'
' Sub ButtonUp(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    Dim lFlag As Long
'    If Button = vbLeftButton Then
'        lFlag = MOUSEEVENTF_LEFTUP
'    ElseIf Button = vbMiddleButton Then
'        lFlag = MOUSEEVENTF_MIDDLEUP
'    ElseIf Button = vbRightButton Then
'        lFlag = MOUSEEVENTF_RIGHTUP
'    End If
'    mouse_event lFlag, 0, 0, 0, 0
'End Sub
'
''simulate the MouseClick event
'
' Sub ButtonClick(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    ButtonDown Button
'    ButtonUp Button
'End Sub
'
''simulate the MouseDblClick event
'
' Sub ButtonDblClick(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    ButtonClick Button
'    ButtonClick Button
'End Sub

