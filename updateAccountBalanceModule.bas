Attribute VB_Name = "updateAccountBalanceModule"
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
' NEW BUTTON: UPDATE ACCOUNT BALANCES
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
Public cashBalance As Currency
Public chequingBalance As Currency
Public giftCardBalance As Currency

'main program of updating balance
Sub updateandShowButton()

If selectedFile = "" Then
    Call getFile
End If
On Error GoTo databaseNotFound
With cn
    .ConnectionString = "Data Source=" & selectedFile
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open
    'proceeds with program given there is no error
    Call updateAndShowBalance
    .Close
End With
Exit Sub

'if database not successfully opened, shows message
databaseNotFound:
    MsgBox "Database cannot be found or another error has occurred.", vbCritical
    Err.Clear
End Sub
'main program to update and show balance
Sub updateAndShowBalance()

Dim SQL As String
Dim cashAmount As Currency
Dim chequingAmount As Currency
Dim giftCardAmount As Currency

'calls helper function
Call updateBalances

SQL = "SELECT CurrentBalance FROM AccountBalances"

'find balances for all accounts
With rs
    .Open SQL, cn
    Do Until .EOF
        cashAmount = .Fields("CurrentBalance")
        .MoveNext
        chequingAmount = .Fields("CurrentBalance")
        .MoveNext
        giftCardAmount = .Fields("CurrentBalance")
        .MoveNext
    Loop
    .Close
End With

'add new values to sheet
Range("cashbalance").value = Format(cashAmount, "Currency")
Range("chequingbalance").value = Format(chequingAmount, "Currency")
Range("giftcardsbalance").value = Format(giftCardAmount, "Currency")

End Sub

'update balance to database
Sub updateBalances()
Dim execution As String

'use helper function to recalculate balances
findBalance ("Cash")
findBalance ("Gift Cards")
findBalance ("Chequing")

'will not check to see if a balance is below zero because of overdrafting and being in debt to someone

'update new values to worksheet
execution = "UPDATE AccountBalances SET CurrentBalance=""" & cashBalance & """ WHERE Account='Cash'"
cn.Execute execution

execution = "UPDATE AccountBalances SET CurrentBalance=""" & chequingBalance & """ WHERE Account='Chequing'"
cn.Execute execution

execution = "UPDATE AccountBalances SET CurrentBalance=""" & giftCardBalance & """ WHERE Account='Gift Cards'"
cn.Execute execution

End Sub

'recalculate the balances
Sub findBalance(accountString As String)

Dim balance As Currency
Dim SQL As String

'subtract expenses incurred with account
SQL = "SELECT Amount FROM Expenses WHERE FromAccount='" & accountString & "'"
balance = 0
With rs
    .Open SQL, cn
    Do Until .EOF
        balance = balance - .Fields("Amount")
        .MoveNext
    Loop
    .Close
End With

'add income added to account
SQL = "SELECT Amount FROM Income WHERE ToAccount='" & accountString & "'"
With rs
    .Open SQL, cn
    Do Until .EOF
        balance = balance + .Fields("Amount")
        .MoveNext
    Loop
    .Close
End With

'add transferred money to account
SQL = "SELECT Amount FROM AccountTransfers WHERE FromAccount='" & accountString & "'"
With rs
    .Open SQL, cn
    Do Until .EOF
        balance = balance - .Fields("Amount")
        .MoveNext
    Loop
    .Close
End With

'subtract money transferred out of this account
SQL = "SELECT Amount FROM AccountTransfers WHERE ToAccount='" & accountString & "'"

With rs
    .Open SQL, cn
    Do Until .EOF
        balance = balance + .Fields("Amount")
        .MoveNext
    Loop
    .Close
End With

'put balance value into the right variable
If accountString = "Cash" Then
    cashBalance = balance
ElseIf accountString = "Gift Cards" Then
    giftCardBalance = balance
Else
    chequingBalance = balance
End If

End Sub

