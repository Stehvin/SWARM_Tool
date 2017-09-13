Attribute VB_Name = "EquivalentWcds"
Function Equivilant_WCD(WCD_Num As String) As String
    
'Function matches WCD number to a equivalent WCD number
'that can be found in the 'SWARM' sheet.
'(And yes, I misspelled 'equivalent' in the function name.)
    
    Dim rngEquivalentList As Range
    Dim OriginalWCD As String
    
    On Error Resume Next
    
    Set rngEquivalentList = Worksheets("WCD Equivalency").Range("A2:E200")
    
    If WCD_Num = "" Then
        Equivilant_WCD = ""
        Exit Function
    End If
    
    'Iterate over every non-blank cell in WCD Equivalency sheet and check
    'if it has the same WCD as the input.
    For Each cell In rngEquivalentList
    
        If cell <> "" And cell Is Not Null And Not IsEmpty(cell) Then

            'If WCD is the same as input, the original WCD is located in
            'column A of the same row.
            If cell = WCD_Num Then
                OriginalWCD = Worksheets("WCD Equivalency").Range("A" & cell.Row)
                Exit For
            End If
        End If
        
    Next cell
        
    Equivilant_WCD = OriginalWCD
End Function
