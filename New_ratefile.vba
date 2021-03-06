Sub new_rate_file()
Indicativerates = ThisWorkbook.Sheets("Setup").Range("C5")
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

''''''''''''''''''''''''''''''''''''''''''''''open file - Indicativerates'''''''''''''''''''''''''''''''''''''''''''
Dim FilenameInd As String
FilenameInd = Indicativerates

'Open excel file based on path
'If path is not found display the message "File not found"
    If Dir(FilenameInd) = "" Then
    
        MsgBox "File not found " & FilenameInd
    
    Exit Sub
    
    End If

'Open file
Set Indicativerates_file = Workbooks.Open(FilenameInd, UpdateLinks:=False) 'Remove update popin window
Application.DisplayAlerts = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''Get data from incentive rates file and set in BILL DISCOUNTED REGISTER file'''''''''''''''''''''
ThisWorkbook.Sheets("Setup").Activate
Dim vRateLastrow As Long
vRateLastrow = Range("Q" & Rows.Count).End(xlUp).Row

For av = 2 To vRateLastrow
    ThisWorkbook.Sheets("Setup").Activate
    Dim vStr As String
    Dim vBuyingTT As String
    Dim vSheetCurr As String
    Dim TT As Long

    'From sheet Setup in column Q get Buying TT currency
    vBuyingTT = Cells(av, 17).Value
    vSheetCurr = Cells(av, 18).Value
    
    Indicativerates_file.Activate
    Sheets("RATE0104").Select
    

    LastRowTT = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Loop to find currency to get T.T amount
    For TT = 1 To LastRowTT
        If Cells(TT, 2).Value = vBuyingTT And Not IsEmpty(Cells(TT, 2).Value) Then
            
           vStr = Cells(TT, 5).Value
            
            Exit For
            
        End If
        
    Next TT
    
    'Paste in Register file to do calculation
    Workbooks(REGISTERfile).Activate
    Worksheets(vSheetCurr).Select
    
    Dim lastrowL As Long
    lastrowL = Cells(Rows.Count, 3).End(xlUp).Row
    
    Range("Z1").Value = vStr
    
      
    'Calculate LOC AMT BY MULTIPLYING T.T amount with FGN AMT
    For i = 3 To lastrowL
    If Cells(i, 12).Value <> " " Then
        
       FGN_AMT = Cells(i, 12).Value
       LOC_AMT = vStr * FGN_AMT
       
       Cells(i, 13).Value = LOC_AMT
        
    End If
        
    Next i
    
    'Clear data in range z1
    Range("Z1").Clear
    
    'Calculate maturity date for sheet eur
     Worksheets(vSheetCurr).Select
     'Call new_calculate_maturity
    Call calculate_maturity_date
    
Next av

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''MUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Workbooks(REGISTERfile).Activate
Dim s_MUR As String: s_MUR = "MUR"
        
'If sheet MUR is found
If DoesSheetExists(s_MUR) Then
    Worksheets("MUR").Select
    
    
    Dim lastrowMUR As Long
    lastrowMUR = Cells(Rows.Count, 3).End(xlUp).Row
    
    'Copy data from FGN AMT to LOC AMT
    For i = 3 To lastrowMUR
    If Cells(i, 12).Value <> " " Then
        
       FGN_AMT = Cells(i, 12).Value
       LOC_AMT = FGN_AMT
       
       Cells(i, 13).Value = LOC_AMT
        
    End If
        
    Next i
    
    'Calculate maturity date for sheet MUR
    'Call new_calculate_maturity
    Call calculate_maturity_date
    
End If

End Sub
