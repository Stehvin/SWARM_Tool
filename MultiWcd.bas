Attribute VB_Name = "MultiWcd"
Function Multi_WCD(WCD_Num As String) As Integer
'Input: WCD Number
'Output: Number of WCD numbers found in SWARM sheet column D
    
    Dim Count As Integer
    Dim MyRange As Range
    Dim findCell As Integer
    Dim cell As Range
    
    On Error Resume Next            'Continues on to next step when Find function outputs an error
    
    Set MyRange = Worksheets("SWARM").Range("D6:D1000")     'All parts in SWARM

    Count = 0
    
    For Each cell In MyRange
        
        'Find function returns number greater than 0 if WCD_Num is in cell
        findCell = Application.WorksheetFunction.Find(WCD_Num, cell.Value)
        
        If findCell >= 1 Then
            Count = Count + 1
        End If
        
        findCell = 0            'Used for error handling. When errors occur, findCell
                                'keeps its last non-error value. So it must be reset to 0.
    Next
        
    Multi_WCD = Count
    
End Function
