
Sub ExtractIPs()
    Dim regEx As Object
    Dim strPattern As String: strPattern = _
        "(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)"
    Dim strInput As String
    Dim MyRange As Range
    Dim i As Long
    Dim j As Long
    Dim objFSO As Object
    Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile("output.txt")

    Set MyRange = ThisWorkbook.ActiveSheet.UsedRange
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = strPattern
    regEx.Global = True

    For i = 1 To MyRange.Rows.Count
        For j = 1 To MyRange.Columns.Count
            strInput = MyRange.Cells(i, j).Value
            If regEx.test(strInput) Then
                objFile.WriteLine regEx.Execute(strInput)(0)
            End If
        Next j
    Next i
    objFile.Close
End Sub
