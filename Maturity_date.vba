Sub calculate_maturity_date()

    Dim LastRowQ As Long
    Dim LastRowC As Long
    LastRowQ = Cells(Rows.Count, 17).End(xlUp).Row
    LastRowC = Cells(Rows.Count, 3).End(xlUp).Row

    'Calculate Maturity date in column P
    For i = LastRowQ + 1 To LastRowC
    If Cells(i, 17).Value <> " " Then
        
       no_of_days = Cells(i, 16).Value
       DATE_OF_DOCUMENT = Cells(i, 15).Value
       
       Dim LValue As String
       LValue = Format(DATE_OF_DOCUMENT, "dd/mm/yyyy")
       
       
       Dim Maturity_date As String
       Dim Maturity_date_format As String
       Maturity_date = DateAdd("d", no_of_days, LValue)
       'Maturity_date_format = Format(Maturity_date, "dd/mm/yyyy")
       
       Cells(i, 17).Value = Maturity_date
        
    End If
        
    Next i
        
        

End Sub
