Attribute VB_Name = "WcdFinder"
Function WCD_Finder(MyRange As Range) As String
'Input: A range containing a WCD number
'Output: The first WCD number, if any are found
'        "Error", if no WCD numbers are found

    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strOutput As String
    Dim regular As String

    strPattern = "(.*)(\()(\w{6})(\))(.*)"      'Regular Expression: anything, WCD number, anything

    If strPattern <> "" Then
        strInput = MyRange.Value

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If regEx.Test(strInput) Then
            WCD_Finder = regEx.Replace(strInput, "$3")
        Else
            WCD_Finder = "Error"
        End If
    End If
End Function
