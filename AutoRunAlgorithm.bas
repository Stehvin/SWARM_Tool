Attribute VB_Name = "AutoRunAlgorithm"
Sub SimpleChoiceAlgorithm()
'Starting by A-date and considering masked column, look up spray area in SWARM,
'randomly choose a booth and an operator, assign operator to booth, remove
'operator from other booth lists, and add time to booth and operator time totals.
'Go to next part. Check if masked, then repeat.

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

'Dim StartTime As Double
'Dim TotalTime As Double
'
'StartTime = Timer

'Make dynamic POS/SWARM sheet operator list, considering who is available to spray
'and what booths are available.
DynamicListCreator

'Save dynamic list as an array, TodaysDynamicList is a constant starting point,
'dynamicList will change over the course of the program.
Dim TodaysDynamicList As Variant
Dim dynamicList As Variant
TodaysDynamicList = Worksheets("Dynamic Drop List").UsedRange.Value2

'Initialize 'SELECTION' sheet arrays
Dim sprayAreaArr, itnArr, adateArr, maskedArr As Variant

'Initialize best run score and arrays
Dim BestRunNumParts, BestRunNames, BestRunITNs As Variant
Dim BestRunScore As Long
Dim LastAdate As Date
Dim LastAdateSpot As Range

'Initialize miscellaneous variables
Dim RandomTaskCode, RandomOperator, IndividualSprayArea As String
Dim AdateSum, seq, booth As Long
Dim nounsArray As Variant

'Initialize 'SWARM' sheet arrays and booth time-keeping array
Dim swarmNumArr, swarmItnArr, swarmOpArr, swarmPartsArr, swarmTimesArr As Variant
Dim swarmNumBlank, swarmItnBlank, swarmOpBlank As Variant
Dim boothArr(1 To 25) As Variant                                'Assumes 25 booths
    
'Initialize BestRunScore to -1 (in case winning run is 0) and find the
'last Adate entry (to be used to calculate the Adate score).
BestRunScore = -1
Set LastAdateSpot = Worksheets("SELECTION").Range("J1000").End(xlUp)    'Hard coded.
LastAdate = LastAdateSpot.Value
While (LastAdate = "12:00:00 AM")
    Set LastAdateSpot = LastAdateSpot.Offset(-1, 0)
    LastAdate = LastAdateSpot.Value
Wend

'Set 'SELECTION' sheet arrays
sprayAreaArr = Worksheets("SELECTION").Range("O2:O200").Value2          'Hard coded.
itnArr = Worksheets("SELECTION").Range("R2:R200").Value2                'Hard coded.
adateArr = Worksheets("SELECTION").Range("J2:J200").Value               'Hard coded.
maskedArr = Worksheets("SELECTION").Range("M2:M200").Value2             'Hard coded.

'Clear 'SWARM' sheet inputs
Worksheets("SWARM").Range("E6:E1000").ClearContents                 'Hard coded.
Worksheets("SWARM").Range("AA6:AA1000").ClearContents               'Hard coded.
Worksheets("SWARM").Range("AC6:AC1000").ClearContents               'Hard coded.

'Set 'SWARM' sheet blank and read-only arrays
swarmNumBlank = Worksheets("SWARM").Range("E6:E1000").Value2        'Hard coded.
swarmItnBlank = Worksheets("SWARM").Range("AC6:AC1000").Value2      'Hard coded.
swarmOpBlank = Worksheets("SWARM").Range("AA6:AA1000").Value2       'Hard coded.
swarmPartsArr = Worksheets("SWARM").Range("D6:D1000").Value2        'Hard coded.
swarmTimesArr = Worksheets("SWARM").Range("H6:H1000").Value2        'Hard coded.

'Loop for trials of random selection runs. Best run is saved.
For trial = 1 To 100

    'Set dynamicList at its beginning point, set 'SWARM' arrays as initially blank,
    'and set booth time-keeping array to 1 shift for all booths.
    dynamicList = TodaysDynamicList
    swarmNumArr = swarmNumBlank
    swarmItnArr = swarmItnBlank
    swarmOpArr = swarmOpBlank
    For boothCount = LBound(boothArr) To UBound(boothArr)
        boothArr(boothCount) = 420                      'Hard coded, assumes 1 shift
    Next boothCount

    'Set Adate score to 0 at the beginning of each trial.
    AdateSum = 0

    'Iterate over parts by A-date and check to see if they are masked.
    'If no task code or operator, loop to next part. Keep looping while
    'any operator still exists on dynamiclist or any parts are still left.
    For i = 1 To 199                                                    'Hard coded.

        If maskedArr(i, 1) = "No" Or sprayAreaArr(i, 1) = "" Then       'Or Adate doesn't exist?
            GoTo NextIteration
        End If

        'Do loop and nounsArray handle cases when more than one spray area is
        'selected for a single part.
        seq = 0
        nounsArray = Split(sprayAreaArr(i, 1), "; ")
        
        Do
            IndividualSprayArea = nounsArray(seq)
            booth = 0

            'Pick a random booth/task code for the part.
            RandomTaskCode = ChooseTaskCode(Right(IndividualSprayArea, _
            Len(IndividualSprayArea) - 13), dynamicList)

            'Pick a random operator (who is qualified) to spray the part.
            RandomOperator = ChooseOperator(RandomTaskCode, dynamicList)

            'Skip to next part if task code doesn't exist or booth is full.
            If RandomOperator = "" Then
                GoTo NextIteration
            End If

            'Remove chosen operator from other booths and remove all other
            'operators from chosen booth.
            Call RemoveOperators(RandomTaskCode, RandomOperator, dynamicList, booth)
'            Worksheets("Dynamic Drop List").UsedRange = dynamicList

            'Input booth/operator into 'SWARM' sheet arrays and remove booth if
            'time exceeds limit.
            Call InputSwarmGetTime(RandomTaskCode & " " & Right(IndividualSprayArea, _
            Len(IndividualSprayArea) - 13), RandomOperator, itnArr(i, 1), booth, _
            dynamicList, swarmNumArr, swarmItnArr, swarmOpArr, swarmPartsArr, _
            swarmTimesArr, boothArr)

            seq = seq + 1
        Loop While seq < UBound(nounsArray) - LBound(nounsArray) + 1

        'Update Adate score
        AdateSum = AdateSum + DateDiff("d", adateArr(i, 1), LastAdate)

NextIteration:
    Next i
    
'    MsgBox ("This run's Adate score: " & AdateSum)

    'If this run was better than the best run, save the info.
    If AdateSum > BestRunScore Then
        BestRunScore = AdateSum
        BestRunNumParts = swarmNumArr
        BestRunNames = swarmOpArr
        BestRunITNs = swarmItnArr
    End If
Next trial

'Fill 'SWARM' sheet with best run info.
For k = LBound(BestRunNumParts) To UBound(BestRunNumParts)
    If BestRunNumParts(k, 1) <> "" Then
        Worksheets("SWARM").Range("E" & 5 + k) = BestRunNumParts(k, 1)
        Worksheets("SWARM").Range("AA" & 5 + k) = BestRunNames(k, 1)
        Worksheets("SWARM").Range("AC" & 5 + k) = BestRunITNs(k, 1)
    End If
Next k

'MsgBox ("Best Adate score: " & BestRunScore)
'TotalTime = Round(Timer - StartTime, 2)
'MsgBox (TotalTime)

'Tell the user when the program is complete.
MsgBox ("Auto run complete.")

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Private Sub DynamicListCreator()
'Make dynamic POS/SWARM sheet operator list, considering who is available to spray
'and which booths are available.

    Worksheets("Dynamic Drop List").UsedRange.ClearContents

    Dim dynDropArr As Variant
    Dim pacTssArr As Variant
    
    'Get dimensions of dynamic drop list array (dynDropArr)
    dynDropArr = Worksheets("DROP LIST").UsedRange.Value2
    maxColDynDrop = UBound(dynDropArr, 2)
    maxRowDynDrop = UBound(dynDropArr, 1) + 1
    
    'Paste dynDropArr back into 'Dynamic Drop List' sheet, but leave row 1 blank
    Col_letter = Split(Cells(1, maxColDynDrop).Address(True, False), "$")(0)
    Worksheets("Dynamic Drop List").Range("A2:" & Col_letter & maxRowDynDrop) _
    .Value2 = dynDropArr
    
    'Put (almost) blank row 1 back into dynDropArr
    Worksheets("Dynamic Drop List").Range("A1").Value2 = 0
    dynDropArr = Worksheets("Dynamic Drop List").UsedRange.Value2

    pacTssArr = Worksheets("PAC TSS").UsedRange.Value2
    maxRowPacTss = UBound(pacTssArr, 1)

'Iterate over task codes to find booth numbers.
    'Iterate over task codes in dynDropArr
    For i = 1 To maxColDynDrop

        'Iterate over task codes in pactssArr
        For j = 2 To maxRowPacTss

            'If task codes match, copy booth into 1st row of dynDropArr
            If dynDropArr(2, i) = pacTssArr(j, 2) Then
                dynDropArr(1, i) = pacTssArr(j, 1)
                Exit For
            End If
        Next j
    Next i

'Remove all booths that aren't available for production.
    Dim boothArr As Variant

    boothArr = Worksheets("SELECTION").Range("A26:B45").Value2      'Hard coded

    'Iterate over booths, checking if each one is available, then removing
    'from dynDropArr.
    For i = 1 To 20                                                 'Hard coded
        If boothArr(i, 2) = "No" Then
            For j = 1 To maxColDynDrop
                If boothArr(i, 1) = dynDropArr(1, j) Then
                    'Delete whole column
                    For k = 1 To maxRowDynDrop
                        dynDropArr(k, j) = ""
                    Next k
                End If
            Next j
        End If
    Next i

'Remove any operators who aren't available to spray.
    Dim operatorArr As Variant

    operatorArr = Worksheets("SELECTION").Range("A3:B23").Value2    'Hard coded

    'Iterate over operators, checking if each one is available, then removing
    'from dynDropArr.
    For i = 1 To 21                                                 'Hard coded
        If operatorArr(i, 2) = "No" Then
            For j = 1 To maxColDynDrop
                For k = 1 To maxRowDynDrop
                    If operatorArr(i, 1) = dynDropArr(k, j) Then
                        dynDropArr(k, j) = ""
                    End If
                Next k
            Next j
        End If
    Next i

'Remove task codes without booths
    For i = 1 To maxColDynDrop
        If dynDropArr(1, i) = "" Then
            For j = 1 To maxRowDynDrop
                dynDropArr(j, i) = ""
            Next j
        End If
    Next i

'Paste dynDropArr back into 'Dynamic Drop List' sheet.
    Worksheets("Dynamic Drop List").UsedRange.ClearContents
    Worksheets("Dynamic Drop List").Range("A1:" & Col_letter & maxRowDynDrop) _
    .Value2 = dynDropArr

End Sub

Private Function ChooseTaskCode(ByVal part As String, dynamicList As Variant) _
As String
'Randomly choose booth (task code) from SWARM sheet, given part text
'and considering booth availability.
    
    Dim size As Long
    Dim swarmParts As Variant
    Dim TaskCodes As Variant
    
    size = 0
    ReDim TaskCodes(0 To size) As Variant
    swarmParts = Worksheets("SWARM").Range("D6:D1000").Value2
    
    'Find task codes in 'SWARM' sheet, then add them to task code array.
    For i = LBound(swarmParts, 1) To UBound(swarmParts, 1)
        If InStr(swarmParts(i, 1), part) > 0 Then
            TaskCodes(size) = Left(swarmParts(i, 1), 12)
            
            'Re-dimension task code array if a new task code is found
            size = size + 1
            ReDim Preserve TaskCodes(0 To size)
        End If
    Next i
    
    Dim in_array As Boolean
    Dim elementInList As Boolean
    in_array = False
    
    'Check if each task code is in dynamic list.
    For i2 = LBound(TaskCodes) To UBound(TaskCodes)
        elementInList = False
        For j2 = LBound(dynamicList, 2) To UBound(dynamicList, 2)
            If TaskCodes(i2) <> "" And dynamicList(2, j2) = TaskCodes(i2) Then
                elementInList = True
                in_array = True
            End If
        Next j2
        
        'If task code isn't in dynamic list, set task code array element as blank.
        If Not elementInList Then
            TaskCodes(i2) = ""
        End If
    Next i2
    
    'Randomly pick a task code. If none available, return a blank string
    Dim choice As String
    If Not in_array Then
        ChooseTaskCode = ""
    Else
        Do
            Randomize
            choice = TaskCodes(Int((UBound(TaskCodes) - LBound(TaskCodes) + 1) * Rnd))
        Loop While choice = ""
        ChooseTaskCode = choice
    End If
End Function

Private Function ChooseOperator(ByVal TaskCode As String, dynamicList As Variant) _
As String
'Randomly choose operator from dynamic drop list, given task code.

    Dim operatorArr As Variant
    Dim size As Long
    
    size = 0
    ReDim operatorArr(0 To size) As Variant
    operatorArr(size) = ""
    
    'Find task code in dynamicList and populate operatorArr with operators
    For j = LBound(dynamicList, 2) To UBound(dynamicList, 2)
        If dynamicList(2, j) = TaskCode Then
            For i = 3 To UBound(dynamicList, 1)
                If dynamicList(i, j) <> "" Then
                    operatorArr(size) = dynamicList(i, j)
                    
                    'Re-dimension operator array if new operator found
                    size = size + 1
                    ReDim Preserve operatorArr(0 To size)
                End If
            Next i
        End If
    Next j
    
    'Randomly choose operator
    'If there is only element, assign it to ChooseOperator. Otherwise, randomly pick
    If size = 0 Then
        ChooseOperator = ""
    ElseIf size = 1 Then
        ChooseOperator = operatorArr(0)
    Else
        'Randomly pick operator from Dynamic Drop List.
        Dim choice As String
        Do
            Randomize
            choice = operatorArr(Int((UBound(operatorArr) - LBound(operatorArr) _
            + 1) * Rnd))
        Loop While choice = ""
        ChooseOperator = choice
    End If
End Function

Private Sub RemoveOperators(ByVal TaskCode As String, ByVal Operator As String, _
ByRef dynamicList As Variant, ByRef booth As Long)
'Remove operator from other booth lists and remove all other operators
'from current booth.

    'Get booth number for a given task code
    Dim pacTssArr As Variant
    pacTssArr = Worksheets("PAC TSS").UsedRange.Value2
    For i = LBound(pacTssArr, 1) To UBound(pacTssArr, 1)
        If pacTssArr(i, 2) = TaskCode Then
            booth = pacTssArr(i, 1)
            Exit For
        End If
    Next i

    'Search for booth and operator in array, and remove entries as appropriate.
    For Irow = LBound(dynamicList, 1) To UBound(dynamicList, 1)
        For Icol = LBound(dynamicList, 2) To UBound(dynamicList, 2)
            
            'Delete other operators in same booth
            If dynamicList(Irow, Icol) = booth Then
                For i = Irow + 2 To UBound(dynamicList, 1)
                    If dynamicList(i, Icol) <> Operator Then
                        dynamicList(i, Icol) = ""
                    End If
                Next i
            End If
            
            'Delete operator in other booths
            If dynamicList(Irow, Icol) = Operator And _
            dynamicList(1, Icol) <> booth Then
                dynamicList(Irow, Icol) = ""
            End If
        Next Icol
    Next Irow
End Sub

Private Sub InputSwarmGetTime(ByVal TaskCodeAndNoun As String, _
ByVal Operator As String, ByVal ITN As String, ByVal booth As Long, _
ByRef dynamicList As Variant, ByRef swarmNumArr As Variant, _
ByRef swarmItnArr As Variant, ByRef swarmOpArr As Variant, _
ByVal swarmPartsArr As Variant, ByVal swarmTimesArr As Variant, _
ByRef boothArr As Variant)
'Input operator selection into SWARM arrays, check total booth time, and remove
'booth when time exceeds limit.

    'Find part index with task code and part noun.
    For sRow = LBound(swarmPartsArr) To UBound(swarmPartsArr)
        If swarmPartsArr(sRow, 1) = TaskCodeAndNoun Then
            
            'Input 1 greater than current number of parts
            If swarmNumArr(sRow, 1) = "" Then
                swarmNumArr(sRow, 1) = 1
            Else
                swarmNumArr(sRow, 1) = swarmNumArr(sRow, 1) + 1
            End If
            
            'Input operator
            swarmOpArr(sRow, 1) = Operator
            
            'Input ITN(s)
            If swarmItnArr(sRow, 1) = "" Then
                swarmItnArr(sRow, 1) = ITN
            Else
                swarmItnArr(sRow, 1) = swarmItnArr(sRow, 1) & ", " & ITN
            End If
            
            'Subtract booth time and remove if too large
            boothArr(booth) = boothArr(booth) - swarmTimesArr(sRow, 1)
            If boothArr(booth) <= 0 Then
                For Icol = LBound(dynamicList, 2) To UBound(dynamicList, 2)
                    If dynamicList(1, Icol) = booth Then
                        For i = 3 To UBound(dynamicList, 1)
                            dynamicList(i, Icol) = ""
                        Next i
                    End If
                Next Icol
            End If
        End If
    Next sRow
End Sub
