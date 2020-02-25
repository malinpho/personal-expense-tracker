VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addEntryForm 
   Caption         =   "Add New Entry"
   ClientHeight    =   5415
   ClientLeft      =   40
   ClientTop       =   380
   ClientWidth     =   6620
   OleObjectBlob   =   "addEntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
'sub that runs when the 'add entry' button is clicked
Public payeetext As String
Public amounttext As String

Private Sub addEntryButton_Click()

payeetext = Trim(payeeComboBox.value)
amounttext = Trim(amountTextBox.value)

'first, see if the amount box is numeric
If Not IsNumeric(amountTextBox.value) Then
    MsgBox "Amount should only contain numbers"
ElseIf TransferButton.value = True Then
    If dateTextBox.value = "" Or payeetext = "" Or fromtext = "" Or amounttext = "" Then
        MsgBox "One or more fields are missing."
    End If
ElseIf dateTextBox.value = "" Or payeetext = "" Or fromComboBox.value = "" Or categoryComboBox.value = "" Or amounttext = "" Then
    MsgBox "One or more fields are missing."
Else: Call approvedInput
End If

End Sub

'goes here if the input is approved
Sub approvedInput()

Dim answer As String
Dim amountString As String

amountString = Format(amountTextBox.value, "Currency")

'send confirmation of entry to user
If ExpenseButton.value = True Then
    answer = MsgBox("Confirm that on " & dateTextBox.value & ", you spent " & amountString & " on " & categoryComboBox.value & " to " & payeeComboBox.value & " that came from " & fromComboBox.value & "?", vbOKCancel, "Confirmation")
ElseIf IncomeButton.value = True Then
    answer = MsgBox("Confirm that on " & dateTextBox.value & ", you received " & amountString & " due to a " & categoryComboBox.value & " from " & payeeComboBox.value & " that is going towards " & fromComboBox.value & "?", vbOKCancel, "Confirmation")
ElseIf TransferButton.value = True Then
    answer = MsgBox("Confirm that on " & dateTextBox.value & ", you transferred " & amountString & " from " & payeeComboBox.value & " to " & fromComboBox.value & "?", vbOKCancel, "Confirmation")
End If

'if confirmation was confirmed, add entry to database
If answer = vbOK Then
    Call approvedEntry
End If

'ask user if they would like to keep adding entries
answer = MsgBox("Would you like to keep adding entries?", vbYesNo, "Keep Adding?")
If answer = vbNo Then
    Unload Me
End If

End Sub
'adding entry to database
Private Sub approvedEntry()

If ExpenseButton.value = True Then
    Call addExpense(dateTextBox.Text, payeetext, fromComboBox.value, categoryComboBox.value, amounttext, notesTextBox.value)
ElseIf IncomeButton.value = True Then
    Call addIncome(dateTextBox.Text, payeetext, fromComboBox.value, categoryComboBox.value, amounttext, notesTextBox.value)
ElseIf TransferButton.value = True Then
    Call addTransfer(dateTextBox.Text, payeetext, fromComboBox.value, amounttext, notesTextBox.value)
End If

End Sub

'if user clicks cancel
Private Sub CancelButton_Click()

Unload Me

End Sub
'if user clicks get date button
Private Sub CommandButton1_Click()

Dim dateVariable As Date

dateVariable = CalendarForm.GetDate
dateTextBox.value = dateVariable

End Sub

'if user clicks expense radio button
Private Sub ExpenseButton_Click()

payeeLabel.Caption = "Payee:"
fromLabel.Caption = "From:"
categoryComboBox.Enabled = True
payeeComboBox.List = partyNameArray
categoryComboBox.List = expenseCategoryNameArray
payeeComboBox.value = ""

End Sub
'if user clicks income radio button
Private Sub IncomeButton_Click()

payeeLabel.Caption = "Payer:"
fromLabel.Caption = "To:"
categoryComboBox.Enabled = True
payeeComboBox.List = partyNameArray
payeeComboBox.value = ""
categoryComboBox.List = incomeCategoryNameArray

End Sub

'if user clicks transfer radio button
Private Sub TransferButton_Click()

payeeLabel.Caption = "From:"
fromLabel.Caption = "To:"
categoryComboBox.Enabled = False
payeeComboBox.List = accountNameArray
payeeComboBox.value = ""

End Sub

'initializes userform
Private Sub UserForm_Initialize()
ExpenseButton.value = True
payeeLabel.Font.Size = 10
fromLabel.Font.Size = 10
categoryLabel.Font.Size = 9
amountLabel.Font.Size = 10
notesLabel.Font.Size = 10
TransactionFrame.Font.Size = 12
addEntryButton.Font.Size = 13

'adds arrays to comboboxes
categoryComboBox.List = expenseCategoryNameArray
fromComboBox.List = accountNameArray
payeeComboBox.List = partyNameArray

End Sub

