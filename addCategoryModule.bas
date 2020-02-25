Attribute VB_Name = "addCategoryModule"
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
' BUTTON: adding categories to database
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
Public categoryName As String

Sub openAddCategoryForm()

If selectedFile = "" Then
    Call getFile
End If

On Error GoTo databaseNotFound
With cn
    .ConnectionString = "Data Source=" & selectedFile
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open
    'proceeds with program given there is no error
    addCategoryForm.Show
    .Close
End With
Exit Sub

'if database not successfully opened, shows message
databaseNotFound:
    MsgBox "Database cannot be found or another error has occurred.", vbCritical
    Err.Clear
End Sub

'execution of adding category to database
Sub addcategory(category As String, addToIncome As Boolean)

Dim execution As String

If addToIncome = True Then
    execution = "INSERT INTO IncomeCategories (IncomeCategory) VALUES (""" & category & """);"
Else
    execution = "INSERT INTO ExpenseCategories (ExpenseCategory) VALUES (""" & category & """);"
End If

cn.Execute execution

End Sub
