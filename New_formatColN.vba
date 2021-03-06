Sub ColN()
'''''''''''''''''''''''''''''''''''''''''''''''Change date format in column N''''''''''''''''''''''''''''''''''''''''''
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

'From request file copy MUR
Workbooks(REGISTERfile).Activate

Dim s As String: s = "MUR"
        
'If sheet MUR is found
If DoesSheetExists(s) Then
    'Change format of column N to date
    Range("N:N").NumberFormat = "dd/MM/yyyy"
    
End If


Dim s_EUR As String: s_EUR = "EUR"
        
'If sheet EUR is found
If DoesSheetExists(s_EUR) Then
    'Change format of column N to date
    Range("N:N").NumberFormat = "dd/MM/yyyy"
End If


Dim s_USD As String: s_USD = "USD"
        
'If sheet USD is found
If DoesSheetExists(s_USD) Then
    'Change format of column N to date
    Range("N:N").NumberFormat = "dd/MM/yyyy"
End If


Dim s_GBP As String: s_GBP = "GBP"
        
'If sheet GBP is found
If DoesSheetExists(s_GBP) Then
    'Change format of column N to date
    Range("N:N").NumberFormat = "dd/MM/yyyy"
End If

End Sub

Sub New_FormatColN()
''''''''''''''''''''''''''''''''''''Loop through all sheets to change format in column N''''''''''''''''''''''''''''''''


REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

Workbooks(REGISTERfile).Activate

' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For i = 1 To WS_Count

   ' Insert your code here.
   ' The following line shows how to reference a sheet within
   ' the loop by displaying the worksheet name in a dialog box.
   
   'MsgBox ActiveWorkbook.Worksheets(I).Name
   
    Worksheets(i).Select
    
   'Change format of column N to date
    Range("O:O").NumberFormat = "mm/dd/yyyy"
    Range("Q:Q").NumberFormat = "mm/dd/yyyy"

Next i

End Sub

Sub CHANGE_FormatDate()
''''''''''''''''''''''''''''''''''''Loop through all sheets to change format in column''''''''''''''''''''''''''''''''


REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

Workbooks(REGISTERfile).Activate

' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For i = 1 To WS_Count

   ' Insert your code here.
   ' The following line shows how to reference a sheet within
   ' the loop by displaying the worksheet name in a dialog box.
   
   'MsgBox ActiveWorkbook.Worksheets(I).Name
   
    Worksheets(i).Select
    
   'Change format of column to date
    Range("F:F").NumberFormat = "dd/mm/yyyy"
    Range("O:O").NumberFormat = "dd/mm/yyyy"
    Range("Q:Q").NumberFormat = "dd/mm/yyyy"
    Range("R:R").NumberFormat = "dd/mm/yyyy"

Next i

End Sub

