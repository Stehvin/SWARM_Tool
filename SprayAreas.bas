Attribute VB_Name = "SprayAreas"
Sub SprayAreas()
'Refreshes 'Spray Areas' sheet to show all spray areas for each WCD.

'List unique WCD numbers on top row of 'Spray Areas' sheet.
UniqueWCDLister

'Find all spray areas of a given WCD and fill in 'Spray Areas' sheet.
Spray_Areas

End Sub
Private Sub UniqueWCDLister()
'Program lists unique WCD numbers found in SWARM sheet on different
'columns in Spray Area sheet.

    Dim regEx As New RegExp
    Dim strPattern As String
    Dim PartWcd As String
    Dim PartCodeCopy As String
    Dim PartCode As Range
    Dim WcdEntry As Range
    Dim arr As Variant
    Dim wcd As Variant
    Dim counter As Integer
    Dim inArray As Boolean
    
    Set PartCode = Worksheets("SWARM").Range("D6")
    Set WcdEntry = Worksheets("Spray Areas").Range("B1")
    
    strPattern = "(.*)(\()(\w{6})(\))(.*)"      'Regular Expression: anything, WCD number, anything

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With

    'Trouble with arrays led to this solution: start with a two element array
    'with each element = "1" and start the counter at 2, so it starts adding
    'WCD numbers in the third array index.
    counter = 2
    arr = Array("1", "1")
    
    'Loops over all part nouns in 'SWARM' sheet
    Do While PartCode.Row <> Worksheets("SWARM").Columns(3).Find("*", _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        If PartCode.Value <> "" Then
            inArray = False
            PartCodeCopy = PartCode.Value
            
            'Testing for WCD number in part noun. If found, assign WCD number
            'to PartWcd.
            If regEx.Test(PartCodeCopy) Then
                PartWcd = regEx.Replace(PartCodeCopy, "$3")
                
                'Testing to see if WCD is already in array.
                For Each wcd In arr
                    If PartWcd = wcd Then
                        inArray = True
                    End If
                Next wcd
                
                'If WCD is not in array, add it to 'Spray Areas' sheet and array.
                'Also, re-dimension array.
                If Not inArray Then
                    WcdEntry.Value = PartWcd
                    Set WcdEntry = WcdEntry.Offset(0, 1)
                    ReDim Preserve arr(0 To counter)
                    arr(counter) = PartWcd
                    counter = counter + 1
                End If
            End If
        End If
        'Move to the next part in SWARM sheet.
        Set PartCode = PartCode.Offset(1, 0)
    Loop
End Sub
Private Sub Spray_Areas()
'Program takes list of WCD numbers found in 'Spray Areas' sheet and
'lists every spray noun found with that WCD number under said number.

    Dim Count As Integer
    Dim MyRange As Range
    Dim SprayArea As Range
    Dim WcdCell As Range
    Dim findCell As Integer
    Dim cell As Range
    Dim WCD_Num As String
    
    On Error Resume Next            'Continues on to next step when Find function outputs an error
    
    Set MyRange = Worksheets("SWARM").Range("D6:D1000")     'All parts in SWARM
    Set SprayArea = Worksheets("Spray Areas").Range("B3")   'Sets first spray noun location
    Set WcdCell = Worksheets("Spray Areas").Range("B1")     'Sets first WCD number
    
    Dim lastCol As Long
    lastCol = Worksheets("Spray Areas").Cells(1, Columns.Count).End(xlToLeft).Column
    'Hard coded last row = 15
    Worksheets("Spray Areas").Range(Cells(3, 2), Cells(15, lastCol)).ClearContents
    
    'Loops over all WCD numbers in 'Spray Areas' sheet
    Do While WcdCell.Value <> ""
        WCD_Num = WcdCell.Value
        Count = 0
        
        'Loops over all part nouns
        For Each cell In MyRange
        
            If cell.Value <> "" Then
                'Find function returns integer greater than 0 if WCD_Num is in cell
                findCell = Application.WorksheetFunction.Find(WCD_Num, cell.Value)
                
                'If found in 'SWARM' add to 'Spray Areas' beneath WCD number.
                If findCell >= 1 Then
                    SprayArea.Value = cell.Value
                    Set SprayArea = SprayArea.Offset(1, 0)
                    Count = Count + 1
                End If
                
                findCell = 0            'Used for error handling. When errors occur, findCell
            End If                      'keeps its last non-error value. So it must be reset to 0.
        Next
        
        Set WcdCell = WcdCell.Offset(0, 1)
        Set SprayArea = SprayArea.Offset(-Count, 1)
    Loop

End Sub


