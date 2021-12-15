Attribute VB_Name = "Module1"

Sub MacroLinkRemover()
'PURPOSE: Remove an external workbook reference from all shapes triggering macros
'Source: www.ExcelForFreelancers.com
Dim Shp As Shape
Dim MacroLink, NewLink As String
Dim SplitLink As Variant

  For Each Shp In ActiveSheet.Shapes 'Loop through each shape in worksheet
  
    'Grab current macro link (if available)
    On Error GoTo NextShp
      MacroLink = Shp.OnAction
    
    'Determine if shape was linking to a macro
      If MacroLink <> "" And InStr(MacroLink, "!") <> 0 Then
        'Split Macro Link at the exclaimation mark (store in Array)
          SplitLink = Split(MacroLink, "!")
        
        'Pull text occurring after exclaimation mark
          NewLink = SplitLink(1)
        
        'Remove any straggling apostrophes from workbook name
            If Right(NewLink, 1) = "'" Then
              NewLink = Left(NewLink, Len(NewLink) - 1)
            End If
        
        'Apply New Link
          Shp.OnAction = NewLink
      End If
NextShp:
  Next Shp
End Sub

