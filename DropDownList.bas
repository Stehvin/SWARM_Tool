Attribute VB_Name = "DropDownList"
Sub DropDownList()
'Program goes through each part in the SWARM sheet to create an operator drop down
'list based on the task code and data entered in the PAC TSS sheet.

    Dim oldSwarmArr As Variant
    Dim swarmArr As Variant
    Dim pacTssArr As Variant
    Dim dropListArr As Variant
    Dim rowCountArr As Variant
    Dim opCheckArr As Variant
    Dim endRowPac As Long
    Dim endColDrop As Long
    Dim arrCounter As Long
    Dim opCell As Range
    Dim namePass As Boolean
    
    opCheckArr = Worksheets("SELECTION").Range("A3:A23").Value2         'Hard coded
    
    'Clear contents of drop list sheet.
    Worksheets("DROP LIST").UsedRange.ClearContents
    
    'Set un-processed 'SWARM' array
    oldSwarmArr = Worksheets("SWARM").Range("D6:D1000").Value2          'Hard coded
    
    'Set 'PAC TSS' array
    endRowPac = Worksheets("PAC TSS").Range("B2").End(xlDown).Row
    pacTssArr = Worksheets("PAC TSS").Range("B2:D" & endRowPac).Value2
    
    'Set 'SWARM' and 'DROP LIST' arrays
    endColDrop = 1
    ReDim swarmArr(1 To endColDrop)
    For i = LBound(oldSwarmArr, 1) To UBound(oldSwarmArr, 1)
        If oldSwarmArr(i, 1) <> "" And oldSwarmArr(i, 1) <> 0 Then
            swarmArr(endColDrop) = oldSwarmArr(i, 1)
            endColDrop = endColDrop + 1
            ReDim Preserve swarmArr(1 To endColDrop)
        End If
    Next i
    endColDrop = endColDrop - 1
    ReDim Preserve swarmArr(1 To endColDrop)
    ReDim dropListArr(1 To 20, 1 To endColDrop)                    'Semi-hard coded
    ReDim rowCountArr(1 To 1, 1 To endColDrop)
    
    'Loop over parts in 'SWARM'
    For i = 1 To UBound(swarmArr, 1)

        'Set task code as first row in 'DROP LIST' array
        dropListArr(1, i) = Left(swarmArr(i), 12)
        
        arrCounter = 1
        
        'Loop over 'PAC TSS' task codes
        For j = LBound(pacTssArr, 1) To UBound(pacTssArr, 1)
        
            'If task code match, add operator to drop list array, and increment
            'the array row counter.
            If pacTssArr(j, 1) = dropListArr(1, i) Then
            
                namePass = False
                
                For k = LBound(opCheckArr, 1) To UBound(opCheckArr, 1)
                    If pacTssArr(j, 3) = opCheckArr(k, 1) Then
                        namePass = True
                    End If
                Next k
                
                If namePass Then
                    dropListArr(1 + arrCounter, i) = pacTssArr(j, 3)
                    arrCounter = arrCounter + 1
                End If
            End If
        Next j
        
        rowCountArr(1, i) = arrCounter
    
    Next i
    
    'Paste drop list array into 'DROP LIST'
    Col_letter = Split(Cells(1, endColDrop).Address(True, False), "$")(0)
    Worksheets("DROP LIST").Range("A1:" & Col_letter & 20) = dropListArr    'Semi-hard coded
                
    Set opCell = Worksheets("SWARM").Range("AA6")
    Dim i2 As Long
    Dim bottomDrop As Long
    i2 = 0
    
    'Loop over parts in un-processed 'SWARM'
    For i = LBound(oldSwarmArr, 1) To UBound(oldSwarmArr, 1)
        
        If oldSwarmArr(i, 1) <> "" And oldSwarmArr(i, 1) <> 0 Then
            'Create and format a drop down list with the range of names
            'in 'DROP LIST' sheet
            i2 = i2 + 1
            columnLetter = Split(Cells(1, i2).Address(True, False), "$")(0)
            bottomDrop = rowCountArr(1, i2)
            
            With opCell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="='DROP LIST'!" & columnLetter & 2 & ":" & _
                columnLetter & bottomDrop
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
            End With
            
        End If
        
        Set opCell = opCell.Offset(1, 0)
    Next i
End Sub
