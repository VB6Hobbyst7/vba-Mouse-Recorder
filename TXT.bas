Attribute VB_Name = "TXT"
'https://newbedev.com/vb-vba-import-csv-to-array-code-example

'VBA function to open a CSV file in memory and parse it to a 2D
'array without ever touching a worksheet:

Function TXTtoArray(sFile$)
    Dim c&, i&, j&, p&, d$, s$, rows&, cols&, a, r, v
    Const Q = """", QQ = Q & Q
    Const ENQ = ""        'Chr(5)
    Const ESC = ""        'Chr(27)
    Const COM = ","
    
    d = OpenTextFile$(sFile)
    If LenB(d) Then
        r = Split(Trim(d), vbCrLf)
        rows = UBound(r) + 1
        cols = UBound(Split(r(0), ",")) + 1
        ReDim v(1 To rows, 1 To cols)
        For i = 1 To rows
            s = r(i - 1)
            If LenB(s) Then
                If InStrB(s, QQ) Then s = Replace(s, QQ, ENQ)
                For p = 1 To Len(s)
                    Select Case Mid(s, p, 1)
                    Case Q:   c = c + 1
                    Case COM: If c Mod 2 Then Mid(s, p, 1) = ESC
                    End Select
                Next
                If InStrB(s, Q) Then s = Replace(s, Q, "")
                a = Split(s, COM)
                For j = 1 To cols
                    s = a(j - 1)
                    If InStrB(s, ESC) Then s = Replace(s, ESC, COM)
                    If InStrB(s, ENQ) Then s = Replace(s, ENQ, Q)
                    v(i, j) = s
                Next
            End If
        Next
        TXTtoArray = v
    End If
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
        .Close
    End With
End Function

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
'
'@AUTHOR ROBERT TODAR
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    
    Dim Temp As String
    
    Select Case ArrayDimensionLength(SourceArray)
        'SINGLE DIMENTIONAL ARRAY
    Case 1
        Temp = Join(SourceArray, Delimiter)
        
        '2 DIMENSIONAL ARRAY
    Case 2
        Dim RowIndex As Long
        Dim ColIndex As Long
            
        'LOOP EACH ROW IN MULTI ARRAY
        For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
            'LOOP EACH COLUMN ADDING VALUE TO STRING
            For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                Temp = Temp & SourceArray(RowIndex, ColIndex)
                If ColIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
            Next ColIndex
                
            'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
            If RowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
        
        Next RowIndex
    End Select
    
    ArrayToString = Temp
    
End Function

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    Dim i As Integer
    Dim Test As Long

    On Error GoTo catch
    Do
        i = i + 1
        Test = UBound(SourceArray, i)
    Loop
    
catch:
    ArrayDimensionLength = i - 1

End Function

Function TXTread(sPath As String) As String
    If Dir(sPath) = "" Then
        MsgBox "File was not found."
        Exit Function
    End If
    '    Close
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TXTread = TXTread & sTXT & vbLf
    Loop
    Close
    If Len(TXTread) = 0 Then
        TXTread = ""
    Else
        TXTread = Left(TXTread, Len(TXTread) - 1)
    End If
End Function

Function TXToverwrite(sFile As String, sText As String)
    On Error GoTo Err_Handler
    Dim FileNumber As Integer
 
    FileNumber = FreeFile        ' Get unused file number
    Open sFile For Output As #FileNumber        ' Connect to the file
    Print #FileNumber, sText        ' Append our string
    Close #FileNumber        ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: TXToverwrite" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function TxtAppend(sFile As String, sText As String)
    On Error GoTo Err_Handler
    Dim iFileNumber           As Integer
 
    iFileNumber = FreeFile        ' Get unused file number
    Open sFile For Append As #iFileNumber        ' Connect to the file
    Print #iFileNumber, sText        ' Append our string
    Close #iFileNumber        ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Txt_Append" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function


