Sub New_autogeneratecolA()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''MUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''Autogenerate number - Col A BILLS_DISCOUNTED_REGISTER ''''''''''''''''''''''''''''''''''''''''''''''''''
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")


Workbooks(REGISTERfile).Activate
Dim s_MUR As String: s_MUR = "MUR"
'If sheet MUR is found
If DoesSheetExists(s_MUR) Then
    Sheets("MUR").Select
    
    Dim lastrowColA_MUR As Long
    Dim lastrowColB_MUR As Long
    lastrowColA_MUR = Range("A" & Rows.Count).End(xlUp).Row
    lastrowColB_MUR = Range("B" & Rows.Count).End(xlUp).Row
    
    Range("A" & lastrowColA_MUR).Select
    
    
    Position = ActiveCell.Row
    
    ActiveCell.Select
    
    Dim i As Integer
    
    Dim initialCount As Long
    Dim strAlpha As String
    Dim counterA As Long
    
        If Position > 2 Then
        initialCount = RetNum(Range("A" & lastrowColA_MUR).Value)
        strAlpha = RetNonNum(Range("A" & lastrowColA_MUR).Value)
        
        
        For i = Position + 1 To ((lastrowColB_MUR - lastrowColA_MUR) + Position)
            counterA = counterA + 1
            Cells(i, 1).Value = strAlpha & (initialCount + counterA)
        Next i
        
        Else
        
        ActiveCell.Offset(1, 0).Value = "BD0001"
        
        initialCount = RetNum("BD0001")
        strAlpha = RetNonNum("BD0001")
        
        
        For i = Position + 2 To ((lastrowColB_MUR - lastrowColA_MUR) + Position)
            
            counterA = counterA + 1
            Cells(i, 1).Value = strAlpha & (initialCount + counterA)
        Next i
    
    End If

End If
End Sub
Sub loop_AutofillColA()
Dim WS_Count As Integer
Dim X As Integer
Dim cid As String
Dim counterA As Long
Dim counter As Long
        

REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

Workbooks(REGISTERfile).Activate

' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For X = 1 To WS_Count
    
       ' Insert your code here.
       ' The following line shows how to reference a sheet within
       ' the loop by displaying the worksheet name in a dialog box.
       
       'MsgBox ActiveWorkbook.Worksheets(I).Name
       
        Worksheets(X).Select
        
       'Change format of column N to date
'''''''''''''''''''''''''''''''''''''''''''''''''''''Sheet MUR'''''''''''''''''''''''''''''''''''''''''''''''''
      If Worksheets(X).Name = "MUR" Then
        Worksheets(X).Select
        
        Dim lastrowColAMUR As Long
        Dim lastrowColBMUR As Long
        lastrowColAMUR = Range("B" & Rows.Count).End(xlUp).Row
        lastrowColBMUR = Range("C" & Rows.Count).End(xlUp).Row
        
        Range("B" & lastrowColAMUR).Select
        
        
        Position = ActiveCell.Row
        
        ActiveCell.Select
        
        Dim i As Integer
        
        Dim initialCount As Long
        Dim strAlpha As String
        
        'counter = 0
'        Dim counterA1 As Long
        
            If Position > 2 Then
                initialCount = RetNum(Range("B" & lastrowColAMUR).Value)
                strAlpha = RetNonNum(Range("B" & lastrowColAMUR).Value)
                
                
                For i = Position + 1 To ((lastrowColBMUR - lastrowColAMUR) + Position)
                    counterA = counterA + 1
                    counterfinal = initialCount + counterA
                    cid = Format(counterfinal, "0000")
                    Cells(i, 2).Value = strAlpha & cid
                Next i
            
            Else
            
                ActiveCell.Offset(1, 0).Value = "BD0001"
                
                initialCount = RetNum("BD0001")
                strAlpha = RetNonNum("BD0001")
                
                
                For i = Position + 2 To ((lastrowColBMUR - lastrowColAMUR) + Position)
                    
                    counterA = counterA + 1
                    counterfinal = initialCount + counterA
                    cid = Format(counterfinal, "0000")
                    Cells(i, 2).Value = strAlpha & cid
                Next i
            
            End If
        End If
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''Sheet not MUR'''''''''''''''''''''''''''''''''''''''''''''''''
             
        If Worksheets(X).Name <> "MUR" Then
        Worksheets(X).Select
        'If worksheet name is not MUR
        Worksheets(X).Select
        
'        Dim counterB As Long
'        Dim counterB1 As Long
        Dim lastrowColA As Long
        Dim lastrowColB As Long
        lastrowColA = Range("B" & Rows.Count).End(xlUp).Row
        lastrowColB = Range("C" & Rows.Count).End(xlUp).Row
        
        Range("B" & lastrowColA).Select
        
        
        Position = ActiveCell.Row
        
        ActiveCell.Select
        
        If Position > 2 Then
        
            initialCount = RetNum(Range("B" & lastrowColA).Value)
            strAlpha = RetNonNum(Range("B" & lastrowColA).Value)
            
            counter = 0
            For i = Position + 1 To ((lastrowColB - lastrowColA) + Position)
                counter = counter + 1
                counterfinal = initialCount + counter
                cid = Format(counterfinal, "0000")
                Cells(i, 2).Value = strAlpha & cid
                
            Next i
            
        Else
            
            ActiveCell.Offset(1, 0).Value = "BD" & Worksheets(X).Name & "0001"
            
            'Dim var As Variant

            var = ActiveCell.Offset(2, 0).Value
            
            'MsgBox var
            
            initialCount = RetNum("BD" & Worksheets(X).Name & "0001")
            strAlpha = RetNonNum("BD" & Worksheets(X).Name & "0001")
            
            'Dim cid As String
            'cid = Format(initialCount, "0000")
            
            
            counter = 0
            For i = Position + 2 To ((lastrowColB - lastrowColA) + Position)
                
                counter = counter + 1
                counterfinal = initialCount + counter
                cid = Format(counterfinal, "0000")
                
                Cells(i, 2).Value = strAlpha & cid
                'counter = Format(counterB, "0000")
                'Cells(I, 1).Value = strAlpha & (initialCount + counterB)
                'Cells(I, 1).Value = strAlpha & (cid + counterB)
            Next i
        
        End If
            
        
    
    End If
       
    
    Next X

End Sub

Function RetNum(str As String)
'updateby Extendoffice
    Dim xRegEx As Object
    Set xRegEx = CreateObject("vbscript.regexp")
    xRegEx.Global = True
    xRegEx.Pattern = "[^\d]+"
    RetNum = xRegEx.Replace(str, "")
    Set xRegEx = Nothing
End Function
Function RetNonNum(str As String)
    Dim xRegEx As Object
    Set xRegEx = CreateObject("vbscript.regexp")
    xRegEx.Global = True
    xRegEx.Pattern = "[\d]+"
    RetNonNum = xRegEx.Replace(str, "")
    Set xRegEx = Nothing
End Function



