Attribute VB_Name = "WipRefresh"
Sub InventoryOnWorkRefresh()
'Refreshes the shop's WIP, as displayed in the Inventory application.

'Run three subroutines defined below.
ITN_Database_Refresh
GetFromInventory
RefreshSprayAreas

Worksheets("SELECTION").Activate

'Tell the user when the program is done.
MsgBox ("WIP refresh complete.")

End Sub
Private Sub ITN_Database_Refresh()

'Saves all spray areas and masking yes/no dropdowns in 'SELECTION' sheet by ITN.
'This information is saved in the 'ITN Database' sheet.

    Dim itnArr As Variant
    Dim sprayArr As Variant
    Dim maskedArr As Variant

    'Make arrays for 3 columns
    itnArr = Worksheets("SELECTION").Range("R2:R200").Value2        'Hard-coded
    sprayArr = Worksheets("SELECTION").Range("O2:O200").Value2      'Hard-coded
    maskedArr = Worksheets("SELECTION").Range("M2:M200").Value2     'Hard-coded

    'Paste arrays into 'ITN Database'
    Worksheets("ITN Database").Range("A2:A200") = itnArr            'Hard-coded
    Worksheets("ITN Database").Range("B2:B200") = sprayArr          'Hard-coded
    Worksheets("ITN Database").Range("C2:C200") = maskedArr         'Hard-coded

End Sub
Private Sub GetFromInventory()
    
'Pastes Inventory info into 'Inventory WIP' sheet and sorts it by Adate.
    
    'Clear 'Inventory Drop'.
    'Must be done twice b/c first clear unhides all the hidden rows,
    'then second clear actually clears all the previously hidden data.
    Worksheets("Inventory Drop").UsedRange.Clear
    Worksheets("Inventory Drop").UsedRange.Clear
    
    'Use connection objects to get all Inventory info and put it in 'Inventory Drop'.
    'Made by macro recording. Data tab --> From Access

    With Worksheets("Inventory Drop").ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;" _
        "Data Source=redacted;Mode=Share Deny Write;Extended Properties="""";" _
        "Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";" _
        "Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=6;" _
        "Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;" _
        "Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";" _
        "Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;" _
        "Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without " _
        "Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex " _
        "Data=False;Jet OLEDB:Bypass UserInfo Validation=False"), _
        Destination:=Worksheets("Inventory Drop").Range("$A$1")).QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array("redacted")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = "redacted"
        .Refresh BackgroundQuery:=False
    End With
    
    'Format the Adate column as a date.
    'Only currently need the Adate column to be formated as dates. If more
    'dates are needed, do the same command with their respective colummns.
    Worksheets("Inventory Drop").Columns(41).NumberFormat = "MM/DD/YYYY"

    'Filter shop-specific parts, then sort by ascending Adate.
    'Made using macro recording (with some slight updating).
    Worksheets("Inventory Drop").ListObjects("Table_Inventory.accde"). _
        Range.AutoFilter Field:=19, Criteria1:="redacted"
    Worksheets("Inventory Drop").ListObjects("Table_Inventory.accde"). _
        Sort.SortFields.Add Key:=Range _
        ("Table_Inventory.accde[[#All], [InvAdate]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Worksheets("Inventory Drop").ListObjects( _
        "Table_Inventory.accde").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Copy info from 'Inventory Drop' to 'Inventory WIP'.
    'Why? --> Need to keep the row value of the first part static.
    Worksheets("Inventory WIP").UsedRange.Clear
    Worksheets("Inventory Drop").UsedRange.Copy
    Worksheets("Inventory WIP").Range("A1").PasteSpecial Paste:=xlPasteFormats
    Worksheets("Inventory WIP").Range("A1").PasteSpecial Paste:=xlPasteValues
    
    'Hopefully aleviates some memory.
    Worksheets("Inventory Drop").UsedRange.Clear
    Worksheets("Inventory Drop").UsedRange.Clear

End Sub
Private Sub RefreshSprayAreas()
    
'Program goes through each WCD in 'SELECTION' sheet to create a part noun
'drop down list based on the data entered in the 'Spray Areas' sheet.
    Dim rngSelectArea As Range
    Dim selectArr As Variant
    Dim sprayAreaArr As Variant
    Dim itnDbArr As Variant
    Dim selectSprayArr As Variant
    Dim selectMaskArr As Variant
    Dim endWcd As Long
    Dim maxColSprayArea As Long

    'Clear all current drop downs. ITN database has saved selections.
    Set rngSelectArea = Worksheets("SELECTION").Range("O2")
    Worksheets("SELECTION").Range("O2:O200").ClearContents
    Worksheets("SELECTION").Range("M2:M200") = "No"

    endWcd = Worksheets("Spray Areas").Range("B1").End(xlToRight).Column
    'Converts column number to letter.
    Col_letter = Split(Cells(1, endWcd).Address(True, False), "$")(0)

    selectArr = Worksheets("SELECTION").Range("M2:S200").Value2         'Hard coded
    sprayAreaArr = Worksheets("Spray Areas").Range("B1:" & Col_letter & "8").Value2
    maxColSprayArea = endWcd - 1
    itnDbArr = Worksheets("ITN Database").Range("A2:C200").Value2       'Hard coded
    selectSprayArr = Worksheets("SELECTION").Range("O2:O200").Value2    'Hard coded
    selectMaskArr = Worksheets("SELECTION").Range("M2:M200").Value2     'Hard coded

    'Loop over all WCD numbers in 'SELECTION' array
    For i = 1 To 199

        'If the WCD cell is blank or has an error, continue to the next WCD
        If selectArr(i, 7) = "" Then
        
            'Populate drop down cell with blank
            With rngSelectArea.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="='Spray Areas'!A20"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
            End With
            
        Else

            'Loops over WCD entries in 'Spray Areas' sheet.
            For i2 = 1 To maxColSprayArea

                'If WCDs match, populate dropdown box
                If sprayAreaArr(1, i2) = selectArr(i, 7) Then
                    Column_letter = Split(Cells(1, i2 + 1).Address(True, _
                        False), "$")(0)
                    lastRow = Worksheets("Spray Areas").Range(Column_letter _
                        & 1).End(xlDown).Row

                    'Populate drop down cell
                    With rngSelectArea.Validation
                        .Delete
                        .Add Type:=xlValidateList, _
                            AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="='Spray Areas'!" & _
                            Column_letter & 3 & ":" & Column_letter & lastRow
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

            Next i2
        End If

        'If ITN in ITN database, automatically add spray area
        ' and masking yes/no already selected.
        For i3 = 1 To 199
            If itnDbArr(i3, 1) = CStr(selectArr(i, 6)) Then
                selectSprayArr(i, 1) = itnDbArr(i3, 2)
                selectMaskArr(i, 1) = itnDbArr(i3, 3)
                Exit For
            End If
        Next i3

        Set rngSelectArea = rngSelectArea.Offset(1, 0)
    Next i

    'Paste spray area and masked arrays into 'SELECTION' sheet
    Worksheets("SELECTION").Range("O2:O200") = selectSprayArr
    Worksheets("SELECTION").Range("M2:M200") = selectMaskArr
End Sub
