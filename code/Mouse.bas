Attribute VB_Name = "Mouse"
Public MouseArray() As Variant

'declaration for keys event reading
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'declaration for mouse events
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000

'declaration for setting mouse position
Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'declaration for getting mouse position
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type

'Declare sleep
Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)


Sub dragMouse(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
   SetCursorPos x0, y0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Sleep 20
   SetCursorPos x1, y1
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Sub moveFromAtoB(x0 As Long, y0 As Long, x1 As Long, y1 As Long)
Dim steep As Boolean: steep = Abs(y1 - y0) > Abs(x1 - x0)
Dim t As Integer
If steep Then
       '// swap(x0, y0);
        t = x0
        x0 = y0
        y0 = t
       ' // swap(x1, y1);
        t = x1
        x1 = y1
        y1 = t
End If
If x0 > x1 Then
    '// swap(x0, x1);
    t = x0
    x0 = x1
    x1 = t
    '// swap(y0, y1);
    t = y0
    y0 = y1
    y1 = t
End If
    Dim deltax As Integer: deltax = x1 - x0
    Dim deltay As Integer: deltay = Abs(y1 - y0)
    Dim deviation As Integer: deviation = deltax / 2
    Dim ystep As Integer
    Dim y  As Integer: y = y0
    If y0 < y1 Then
        ystep = 1
    Else
        ystep = -1
    End If
    Dim x As Integer
    For x = x0 To x1
        If steep Then
            SetCursorPos y, x
        Else
            SetCursorPos x, y
        End If
        
        deviation = deviation - deltay
        If deviation < 0 Then
            y = y + ystep
            deviation = deviation + deltax
        End If
        DoEvents
        Sleep 2
    Next
End Sub
Function nextRangeUntill(c As Range, off As Long, findme As Variant)
Dim cell As Range
Set cell = c
Do While UCase(cell.Offset(off)) <> UCase(findme)
Set cell = cell.Offset(1)
Loop
Set nextRangeUntill = cell
End Function
Sub MouseReplay(Optional smoothMovement As Boolean)
    ActiveWindow.WindowState = xlMaximized
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim cell As Range
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion
    Set rng = rng.Offset(1).Resize(rng.rows.Count - 1, 1)
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    For Each cell In rng
        'option to smooth with sub MoveFromAtoB
        If smoothMovement = True Then
            Dim lngCurPos As POINTAPI, activeX As Long, activeY As Long
            GetCursorPos lngCurPos
            activeX = lngCurPos.x
            activeY = lngCurPos.y
            moveFromAtoB activeX, activeY, CLng(cell.Value), CLng(cell.Offset(0, 1).Value)
        Else
            SetCursorPos cell, cell.Offset(, 1)
        End If
        
        If cell.Offset(0, 2) > 1 Then
            dragMouse cell.Value, cell.Offset(0, 1), cell.Offset(0, 2), cell.Offset(0, 3)
        ElseIf cell.Offset(0, 2) = 1 Then
            If cell.Offset(1, 2) = 0 Then
                If cell.Offset(2, 2) = 0 And cell.Offset(-1, 2) = 0 Then
                    SingleClick
                    Set cell = cell.Offset(2, 2)
                ElseIf cell.Offset(2, 2) = 1 Then
                    DoubleClick
                        Set cell = cell.Offset(2, 2)
                End If
            End If
        ElseIf cell.Offset(0, 3) = 1 Then
            RightClick
        End If
        
        DoEvents
        Sleep 20
    Next
LoopEnd:
    If Err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Do While GetAsyncKeyState(1) = 1
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        DoEvents
        Loop
        
    End If
End Sub

Sub newRecord()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    ws.Range("A2").CurrentRegion.Offset(1).ClearContents
    ws.Range("R7").CurrentRegion.Offset(1).ClearContents
    ws.Range("recFile").CurrentRegion.ClearContents
End Sub
Sub recordDrag()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim previousX As Long, previousY As Long, activeX As Long, activeY As Long
    Dim previousL As Long, previousR As Long, activeL As Long, activeR As Long
    Erase MouseArray
    Dim arrayCounter As Long: arrayCounter = 1
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        ReDim Preserve MouseArray(1 To arrayCounter)
        GetCursorPos lngCurPos
        activeL = IIf(GetAsyncKeyState(1) = 0, 0, 1)
        activeR = IIf(GetAsyncKeyState(2) = 0, 0, 1)
        activeX = lngCurPos.x
        activeY = lngCurPos.y
            If previousX <> lngCurPos.x Or previousY <> lngCurPos.y Or previousL <> activeL Or previousR <> activeR Then
                previousX = activeX
                previousY = activeY
                previousL = activeL
                previousR = activeR
                MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
                arrayCounter = arrayCounter + 1
            End If
    Loop
LoopEnd:
    If Err = 18 Then
        Application.EnableCancelKey = xlInterrupt
                    Dim arr
            arr = MouseArray
            arr = Filter(arr, ",1,", , vbTextCompare)

        Dim recRng As Range
            Set recRng = ws.Range("RecFile").Offset(1)
            If recRng <> "" Then Set recRng = ws.Cells(rows.Count, recRng.Column).End(xlUp).Offset(1)

            recRng.Resize(UBound(arr), 1) = WorksheetFunction.Transpose(arr)

        Set rng = ws.Range("A" & rows.Count).End(xlUp).Offset(1, 0)
        Set rng = rng.Resize(UBound(arr), 1)
        rng = WorksheetFunction.Transpose(arr)
        rng.TextToColumns rng, comma:=True
        rng.Offset(0, 4) = "DRAG"
        rng.Offset(2).Resize(rng.rows.Count - 3, 5).Delete xlUp
        rng.Resize(1, 5).Delete xlUp
        rng.Offset(0, 2).Resize(1, 2).Value = rng.Offset(1).Resize(1, 2).Value
        rng.Offset(1).Resize(1, 5).Delete xlUp
        MsgBox "Drag recorded."
         Debug.Print "Drag recorded at rows: " & rng.Row & " to " & rng.Row + rng.rows.Count
    End If
End Sub
Sub RecordStart(Optional recordWholeMotion As Boolean)
    ActiveWindow.WindowState = xlMaximized
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim previousX As Long, previousY As Long, activeX As Long, activeY As Long
    Dim previousL As Long, previousR As Long, activeL As Long, activeR As Long
    Erase MouseArray
    Dim arrayCounter As Long: arrayCounter = 1
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        ReDim Preserve MouseArray(1 To arrayCounter)
        GetCursorPos lngCurPos
        activeL = IIf(GetAsyncKeyState(1) = 0, 0, 1)
        activeR = IIf(GetAsyncKeyState(2) = 0, 0, 1)
        activeX = lngCurPos.x
        activeY = lngCurPos.y
        If recordWholeMotion Then
            If previousX <> lngCurPos.x Or previousY <> lngCurPos.y Or previousL <> activeL Or previousR <> activeR Then
                previousX = activeX
                previousY = activeY
                previousL = activeL
                previousR = activeR
                MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
                arrayCounter = arrayCounter + 1
                DoEvents
            End If
        Else
            If previousL <> activeL Or previousR <> activeR Then
                previousX = activeX
                previousY = activeY
                previousL = activeL
                previousR = activeR
                MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
                arrayCounter = arrayCounter + 1
                DoEvents
            End If
        End If
    Loop
LoopEnd:
    If Err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Set rng = ws.Range("A" & rows.Count).End(xlUp).Offset(1, 0)
        Set rng = rng.Resize(UBound(MouseArray), 1)
        rng = WorksheetFunction.Transpose(MouseArray)
        rng.TextToColumns rng, comma:=True
        ws.Range("recFile").Offset(1).Resize(UBound(MouseArray)) = MouseArray
        ws.Range("A3:D3").Delete Shift:=xlUp
        MsgBox "Macro recorded."
        Debug.Print "Macro recorded at rows: " & rng.Row & " to " & rng.Row + rng.rows.Count
    End If
End Sub

Sub SingleClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Sub DoubleClick()
    'Double click as a quick series of two clicks
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Sub RightClick()
    'Right click
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub


Function RecordFileFullName() As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    RecordFileFullName = ws.Range("recFolder") & "\" & ws.Range("recFile") & "_mr.txt"
End Function

Function IsRecordSaved() As Boolean
    IsRecordSaved = Not ThisWorkbook.Sheets("MouseRecord").Range("recFile").Text = ""
End Function

Function RecordedMacro() As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion.Offset(1)
    Set rng = rng.Resize(rng.rows.Count - 1)
    Dim arr
    arr = rng.Value
    RecordedMacro = ArrayToString(arr)
End Function

Sub saveRecord()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")

    If Not IsFolderSelected Then
        MsgBox "Select a folder to store records first"
        Exit Sub
    End If
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion
    Set rng = rng.Offset(1).Resize(rng.rows.Count - 1)
    TXToverwrite RecordFileFullName, RecordedMacro
End Sub

Sub LoadRecord()
    If Not IsFolderSelected Then
        MsgBox "Pick a folder for your recordings first"
        Exit Sub
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("MouseRecord")
    Dim fName As String
    fName = Dialogues.PickRecord(ws.Range("recFolder"))
    If fName = "" Or Right(fName, 7) <> "_mr.txt" Then
        MsgBox "No valid file selected"
        Exit Sub
    End If
    newRecord
    fName = Mid(fName, InStrRev(fName, "\") + 1)
    fName = Left(fName, InStr(1, fName, "_") - 1)
    uMouseRecorder.LoadedRecording.Caption = fName
    ws.Range("recFile") = fName
    Dim recFile As String
    recFile = RecordFileFullName
    Dim arr
    arr = TXTtoArray(recFile)
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion.Offset(1)
    rng.ClearContents
    rng.Resize(UBound(arr, 1), 4) = arr
End Sub

