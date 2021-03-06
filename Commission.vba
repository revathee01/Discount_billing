Dim CommissionFile As String
Dim advice_contra_FILE As String
Dim cAccNum As String
Dim vAccountNum As String
Sub commission_file()
''''''''''''The commission is calculated USD 50 for FCY bills or MUR 500 for MUR bills and inserted in the sheet'''''''
CommissionFile = ThisWorkbook.Sheets("Setup").Range("C7")
advice_contra_FILE = ThisWorkbook.Sheets("Setup").Range("E6")
REGISTERfile = ThisWorkbook.Sheets("Setup").Range("E4")

'Dim Filename As String
'Open Commission file to get the currency - USD or MUR
Dim Filename1 As String
Filename1 = CommissionFile
'Open excel file based on path
'If path is not found display the message "File not found"
If Dir(Filename1) = "" Then

    MsgBox "File not found " & Filename1

Exit Sub

End If

'Open file advice_contra
Set vCommissionFile = Workbooks.Open(Filename1, UpdateLinks:=False) 'Remove update popin window

'Get account number in register file to verify if it is matching in commission file
Workbooks(REGISTERfile).Activate
Range("C2").Select

 'Looping until next visible cell
ActiveCell.Offset(1, 0).Activate
Do While ActiveCell.EntireRow.Hidden = True
    ActiveCell.Offset(1, 0).Activate
Loop

vAccountNum = ActiveCell.Value

vCommissionFile.Activate
Sheets("Commission").Activate
Dim lastrowCommission As Long
lastrowCommission = Range("A" & Rows.Count).End(xlUp).Row

'Loop to find if account number is present in commission file(if present commission is USD 50)
For vAcc = 2 To lastrowCommission
    cAccNum = Cells(vAcc, 1).Value
    If cAccNum = vAccountNum Then
        Workbooks(advice_contra_FILE).Activate
        Worksheets(1).Select
        Range("D11").Value = "50"

        Dim lastrowUSD50 As Long
        lastrowUSD50 = Range("C" & Rows.Count).End(xlUp).Row
        Range("D11:D" & lastrowUSD50).Select
        Selection.FillDown
        'MsgBox Cells(vAcc, 1).Value

    End If
Next vAcc


''Loop to find if account number is present in commission file(if present commission is MUR 500)
Workbooks(advice_contra_FILE).Activate
Worksheets(1).Activate
If Range("D11").Value = "" Then
    Workbooks(advice_contra_FILE).Activate
    Worksheets(1).Select
    Range("D11").Value = "500"

    Dim lastrowMUR500 As Long
    lastrowMUR500 = Range("C" & Rows.Count).End(xlUp).Row
    Range("D11:D" & lastrowMUR500).Select
    Selection.FillDown
End If


'For vAcc1 = 2 To lastrowCommission
'    cAccNum = Cells(vAcc1, 1).Value
'    If cAccNum <> vAccountNum Then
'        Workbooks(advice_contra_FILE).Activate
'        Worksheets(1).Select
'        Range("D11").Value = "500"
'
'        Dim lastrowMUR500 As Long
'        lastrowMUR500 = Range("C" & Rows.Count).End(xlUp).Row
'        Range("D11:D" & lastrowMUR500).Select
'        Selection.FillDown
'        'MsgBox Cells(vAcc, 1).Value
'
'    End If
'Next vAcc1

End Sub
