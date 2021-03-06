Function DoesSheetExists(sh As String) As Boolean
    Dim ws As Worksheet
    
    REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")
    
    On Error Resume Next
    Set ws = Workbooks(REGISTERfile).Sheets(sh)
    On Error GoTo 0

    If Not ws Is Nothing Then DoesSheetExists = True
End Function
