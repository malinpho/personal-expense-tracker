Attribute VB_Name = "addEntryModule"
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
'OPENS USERFORM TO ADD ENTRIES TO PERSONAL EXPENSE TRACKER
Public incomeCategoryNameArray() As Variant
Public expenseCategoryNameArray() As Variant
Public entryDate As Date
Public accountNameArray() As Variant
Public partyNameArray() As Variant
Public firstCombo As String
Public secondcombo As String

Sub openUserForm()

'if a file was not intiatiated, prompts user to pick one
If selectedFile = "" Then
    Call getFile
End If

On Error GoTo databaseNotFound
    With cn
        .ConnectionString = "Data Source=" & selectedFile
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
        'proceeds with program given there is no error
        Call startprogram
        .Close
End With
Exit Sub

'if database not successfully opened, shows message
databaseNotFound:
    MsgBox "Database cannot be found or another error has occurred.", vbCritical
    Err.Clear
End Sub

'main program to start userform
Sub startprogram()

Call initializeAccountArray
Call initializePartyArray
Call initializeExpenseCategoryArray
Call initializeIncomeCategoryArray

addEntryForm.Show

End Sub

'helper method used to initialze income category arrays
Sub initializeIncomeCategoryArray()

Dim SQL As String
Dim i As Integer
Dim categoryName As String
Dim categorycount As Integer

SQL = "SELECT IncomeCategory FROM IncomeCategories"

With rs
    .Open SQL, cn
    Do Until .EOF
        categorycount = categorycount + 1
        .MoveNext
    Loop
    .Close
End With

'put array to correct size
ReDim incomeCategoryNameArray(0 To categorycount - 1)

'add categories to array
With rs
    .Open SQL, cn
    i = 0
    Do Until .EOF
        incomeCategoryNameArray(i) = .Fields("IncomeCategory")
        i = i + 1
        .MoveNext
    Loop
    .Close
End With

End Sub

'used to initialize expense category array
Sub initializeExpenseCategoryArray()

Dim SQL As String
Dim i As Integer
Dim categoryName As String
Dim categorycount As Integer

'same idea as last subroutine
SQL = "SELECT ExpenseCategory FROM ExpenseCategories"

With rs
    .Open SQL, cn
    Do Until .EOF
        categorycount = categorycount + 1
        .MoveNext
    Loop
    .Close
End With

ReDim expenseCategoryNameArray(0 To categorycount - 1)

With rs
    .Open SQL, cn
    i = 0
    Do Until .EOF
        expenseCategoryNameArray(i) = .Fields("ExpenseCategory")
        i = i + 1
        .MoveNext
    Loop
    .Close
End With
End Sub

'used to initiliaze account array
Sub initializeAccountArray()

Dim SQL As String
Dim i As Integer
Dim accountcount As Integer

SQL = "SELECT Account FROM AccountBalances"

With rs
    .Open SQL, cn
    Do Until .EOF
        accountcount = accountcount + 1
        .MoveNext
    Loop
    .Close
End With

ReDim accountNameArray(0 To accountcount - 1)

With rs
    .Open SQL, cn
    i = 0
    Do Until .EOF
        accountNameArray(i) = .Fields("Account")
        i = i + 1
        .MoveNext
    Loop
    .Close
End With

End Sub

'used to initialize party array from 'parties'
Sub initializePartyArray()

Dim SQL As String
Dim i As Integer
Dim namecount As Integer

SQL = "SELECT Party FROM Parties"

With rs
    .Open SQL, cn
    Do Until .EOF
        namecount = namecount + 1
        .MoveNext
    Loop
    .Close
End With

ReDim partyNameArray(0 To namecount - 1)

With rs
    .Open SQL, cn
    i = 0
    Do Until .EOF
        partyNameArray(i) = .Fields("Party")
        i = i + 1
        .MoveNext
    Loop
    .Close
End With

End Sub

'adds user's expense input
Sub addExpense(dateofTransaction As String, payeeString As String, fromString As String, categoryString As String, amount As String, noteString As String)

Dim execution As String

execution = "INSERT INTO Expenses ([Day], [Payee], [Amount], [FromAccount], [Category], [Notes]) VALUES (""" & dateofTransaction & """, """ & payeeString & """, """ & amount & """, """ & fromString & """, """ & categoryString & """,""" & noteString & """);"
cn.Execute execution

End Sub

'adds user's income input
Sub addIncome(dateofTransaction As String, payerString As String, toString As String, categoryString As String, amount As String, noteString As String)
    Dim execution As String
    
    execution = "INSERT INTO Income ([Day], [Payer], [Amount], [ToAccount], [Category], [Notes]) VALUES (""" & dateofTransaction & """, """ & payerString & """, """ & amount & """, """ & toString & """, """ & categoryString & """, """ & noteString & """);"
    cn.Execute execution

End Sub

'add user's input of a transfer
Sub addTransfer(dateofTransaction As String, fromString As String, toString As String, amount As String, noteString As String)
    Dim execution As String
    
    execution = "INSERT INTO AccountTransfers ([Day], [FromAccount], [ToAccount], [Amount], [Notes]) VALUES (""" & dateofTransaction & """, """ & fromString & """, """ & toString & """, """ & amount & """, """ & noteString & """);"
    cn.Execute execution

End Sub


