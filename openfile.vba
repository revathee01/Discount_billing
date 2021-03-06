Dim REQUEST_BILL_DISCOUNTED As String
Dim BILLS_DISCOUNTED_REGISTER As String
Dim Indicativerates As String

Dim REGISTERfile As String
Dim REQUESTfile As String

Dim MainCurrLastRow As String
Dim lastrowMain As Long


Dim WS_Count As Integer
Dim i As Integer

Sub Open_file()

REQUEST_BILL_DISCOUNTED = ThisWorkbook.Sheets("Setup").Range("C3")
BILLS_DISCOUNTED_REGISTER = ThisWorkbook.Sheets("Setup").Range("C4")


''''''''''''''''''''''''''''''''''''''''''''''open file - REQUEST_BILL_DISCOUNTED'''''''''''''''''''''''''''''''''''''''''''
Dim Filename As String
Filename = REQUEST_BILL_DISCOUNTED

'Open excel file based on path
'If path is not found display the message "File not found"
If Dir(Filename) = "" Then

    MsgBox "File not found " & Filename

Exit Sub

End If

'Open file
Set REQUEST_BILL_DISCOUNTED_file = Workbooks.Open(Filename, UpdateLinks:=False) 'Remove update popin window



''''''''''''''''''''''''''''''''''''''''''''''open file - BILLS_DISCOUNTED_REGISTER'''''''''''''''''''''''''''''''''''''''''''
Dim Filename1 As String
Filename1 = BILLS_DISCOUNTED_REGISTER

'Open excel file based on path
'If path is not found display the message "File not found"
If Dir(Filename1) = "" Then

    MsgBox "File not found " & Filename1

Exit Sub

End If

'Open file
Set BILLS_DISCOUNTED_REGISTER_file = Workbooks.Open(Filename1, UpdateLinks:=False) 'Remove update popin window
Application.DisplayAlerts = False

''''''''''''''''''''''''''''''''''''Currency - Col I BILLS_DISCOUNTED_REGISTER ''''''''''''''''''''''''''''''''''''''''''''''''''
'Call Identify_Currency
Call NewColIfilter
''''''''''''''''''''''''''''''''''''''''''''''''Date format in column N ''''''''''''''''''''''''''''''''''''''''''''''''''
'Call ColN
Call New_FormatColN

''''''''''''''''''''''''''''''''''''Autogenerate number - Col A BILLS_DISCOUNTED_REGISTER ''''''''''''''''''''''''''''''''''''''''''''''''''
'Call Autogenerate_ColumnA
Call loop_AutofillColA
''''''''''''''''''''''''''''''''''''''''''''Open incentiverates file ''''''''''''''''''''''''''''''''''''''''''''''''''
Call new_rate_file
'Call ratefile
'''''''''''''''''''''''''''''''''''''''''''''''''''WORKFUSION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The market segment is updated in the register by picking up the information from EBox
'''''''''''''''''''''''''''''''''''''''''''''''''''WORKFUSION''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'An advice and the contra entries is created for bills being discounted customer-wise


End Sub



