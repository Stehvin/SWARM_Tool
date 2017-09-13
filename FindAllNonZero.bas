Attribute VB_Name = "FindAllNonZero"
Function FindAllNonZero(SearchRange As Range) As String
'Function combines all non-blank cells in the search range,
'using a comma delimiter.

Dim outputStr As String
outputStr = ""

'Search every cell in SearchRange, adding every non-blank cell to
'outputStr.
For Each cell In SearchRange
    If cell.Value <> "" Then
        If outputStr <> "" Then
            outputStr = outputStr & ", " & cell.Value
        Else
            outputStr = cell.Value
        End If
    End If
Next cell

FindAllNonZero = outputStr

End Function
