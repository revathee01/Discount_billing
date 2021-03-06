Sub Currency_ColI()
''''''''''''''''''''''''''''''''''''Currency - Col I BILLS_DISCOUNTED_REGISTER ''''''''''''''''''''''''''''''''''''''''''''''''''
REQUESTfile = ThisWorkbook.Sheets("Setup").Range("E3")
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

'From request file copy MUR
Workbooks(REQUESTfile).Activate


Dim X As Integer
Dim NumRows As Long
      ' Set numrows = number of rows of data.
      NumRows = Range("I" & Rows.Count).End(xlUp).Row
      ' Select cell a1.
      Range("J2").Select
      ' Establish "For" loop to loop "numrows" number of times.
      
      For X = 2 To NumRows
         ' Insert your code here.
         ' Selects cell down 1 row from active cell.
         ActiveCell.Offset(1, 0).Select
         Debug.Print ActiveCell.Offset(1, 0).Address(False, False)
         Debug.Print ActiveCell.Offset(1, 0).Value
         
         
         If Trim(ActiveCell.Value) = Trim("MUR") Then
         ActiveCell.Copy
         
         Workbooks(REGISTERfile).Activate
         Sheets("MUR").Select
         
         Dim NumRowsRegister As Long
         NumRowsRegister = Range("J" & Rows.Count).End(xlUp).Row
         Range("J" & NumRowsRegister).Select
         
         ActiveCell.Paste
         
         End If
      Next



End Sub

Sub test()
    
    'Dim RequestLastRow As Range
    'Set the formula in the working
    ''
    ''
        'Ctrl + Shift + End
    MainCurrLastRow = ThisWorkbook.Worksheets("Main").Cells(Worksheets("Main").Rows.Count, "A").End(xlUp).Row
    
   For i = 1 To MainCurrLastRow
   curr = ThisWorkbook.Worksheets("Main").Range("A" & i).Value
    'Ctrl + Shift + End
    'selext sheet
    RequestLastRow = Worksheets("Request").Cells(Worksheets("Request").Rows.Count, "A").End(xlUp).Row
   ''LOOP FOR CURRENCY
    'Insert a column to to find if there is any date
    Worksheets("Request").Columns("I:I").Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Worksheets("Request").Range("I2").Value = "DateCurrency"
    Worksheets("Request").Range("I3:I" & RequestLastRow).Formula = "=IF(TRIM(J3)=""" & curr & """,TRUE,FALSE)"
    Worksheets("Request").Range("I3:I" & RequestLastRow).Value = Worksheets("Request").Range("I3:I" & RequestLastRow).Value


    ''The false value should be inserted into exception
    With Worksheets("Request").Range("A1:AAB" & RequestLastRow)
        .AutoFilter Field:=9, Criteria1:=">=TRUE"
       
        .SpecialCells(xlCellTypeVisible).Copy

        ''paste on the other sheet
        ActiveSheet.Paste
       
    End With
    
    'delete column
    Worksheets("Request").Columns(9).EntireColumn.Delete

Next i

End Sub

Sub NewColIfilter()
    
    Dim curr As String
    'Dim RequestLastRow As Range
    'Set the formula in the working
    ''
    ''
        'Ctrl + Shift + End
    REQUESTfile = ThisWorkbook.Sheets("Setup").Range("E3")
    REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")
    
    MainCurrLastRow = ThisWorkbook.Sheets("Main").Range("A1")
    
    'Dim lastrowMain As Long

    lastrowMain = ThisWorkbook.Sheets("Main").Range("A" & Rows.Count).End(xlUp).Row
     
    'Cells(Worksheets("Main").Rows.Count, "A").End(xlUp).Row
    
   For i = 1 To lastrowMain
   curr = ThisWorkbook.Worksheets("Main").Range("A" & i).Value
    'Ctrl + Shift + End
    'selext sheet
    Workbooks(REQUESTfile).Activate
    RequestLastRow = Worksheets("Request").Cells(Worksheets("Request").Rows.Count, "A").End(xlUp).Row
    
   ''LOOP FOR CURRENCY
    'Insert a column to find if there is any date
    Worksheets("Request").Columns("I:I").Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Worksheets("Request").Range("I2").Value = "DateCurrency"
    Worksheets("Request").Range("I3:I" & RequestLastRow).Formula = "=IF(TRIM(J3)=""" & curr & """,TRUE,FALSE)"
    Worksheets("Request").Range("I3:I" & RequestLastRow).Value = Worksheets("Request").Range("I3:I" & RequestLastRow).Value


    ''The false value should be inserted into exception
    With Worksheets("Request").Range("A2:AAB" & RequestLastRow)
        .AutoFilter Field:=9, Criteria1:=">=TRUE"
       
        Dim lastrowI As Long

        lastrowI = Range("B" & Rows.Count).End(xlUp).Row
        
            
        'If sheet is found
        Workbooks(REGISTERfile).Activate
        If DoesSheetExists(curr) Then
        
        'Sheet present
        
        Else
        'If sheet does not exist
        'sheet GBP is not found
        Workbooks(REGISTERfile).Activate

        'Create new sheet
        Sheets.Add.Name = curr


        Worksheets("MUR").Select
        Range("1:2").Copy Sheets(curr).Range("1:2")
        
        Sheets(curr).Select

                
        End If
        
        ''paste on the other sheet
        'Workbooks(REGISTERfile).Activate
        Workbooks(REQUESTfile).Activate
        Range("A3:A" & lastrowI).SpecialCells(xlCellTypeVisible).Copy
        
        Workbooks(REGISTERfile).Activate
        Worksheets(curr).Select
        Dim lastrowcurr As Long
        lastrowcurr = Range("B" & Rows.Count).End(xlUp).Row
        Range("A" & lastrowcurr + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        
        Workbooks(REQUESTfile).Activate
        Range("B3:C" & lastrowI).SpecialCells(xlCellTypeVisible).Copy
        
        Workbooks(REGISTERfile).Activate
        Worksheets(curr).Select
        'Dim lastrowcurr As Long
        lastrowcurr = Range("B" & Rows.Count).End(xlUp).Row
        Range("C" & lastrowcurr + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            
        Workbooks(REQUESTfile).Activate
        Worksheets("Request").Select
        Range("D3:H" & lastrowI).SpecialCells(xlCellTypeVisible).Copy
        
        
        Workbooks(REGISTERfile).Activate
        Worksheets(curr).Select
        Range("F" & lastrowcurr + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
        
        Workbooks(REQUESTfile).Activate
        Worksheets("Request").Select
        Range("J3:O" & lastrowI).SpecialCells(xlCellTypeVisible).Copy
        
        
        Workbooks(REGISTERfile).Activate
        Worksheets(curr).Select
        Range("K" & lastrowcurr + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
  
        'Add border
        
        Workbooks(REGISTERfile).Activate
        Worksheets(curr).Select
        Dim lastrowNewdata As Long
        lastrowNewdata = Range("C" & Rows.Count).End(xlUp).Row
        
        Dim rng1 As Range
        Set rng1 = Range("A2" & ":Q" & lastrowNewdata)
       
            With rng1.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            
        
        
    End With
  
        
        'Remove filter in sheet Request
        Workbooks(REQUESTfile).Activate
        Worksheets("Request").Select
        
        If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
        
        'delete column I
        Workbooks(REQUESTfile).Activate
        Worksheets("Request").Columns(9).EntireColumn.Delete
        
    Next i



End Sub
