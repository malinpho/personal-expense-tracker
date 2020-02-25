VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addCategoryForm 
   Caption         =   "Add Category"
   ClientHeight    =   1635
   ClientLeft      =   40
   ClientTop       =   380
   ClientWidth     =   4720
   OleObjectBlob   =   "addCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'adds category to database
Private Sub addButton_Click()
Dim category As String
Dim addToIncome As Boolean
category = categoryTextBox.value

'confirm addition of category
If IncomeRadio.value = True Then
    addToIncome = True
    answer = MsgBox("Add the category " & category & " to the list of possible Income Categories?", vbOKCancel, "Confirmation")
Else
    addToIncome = False
    answer = MsgBox("Add the category " & category & " to the list of possible Expense Categories?", vbOKCancel, "Confirmation")
End If

'call function to add category to database
If answer = vbOK Then
    Call addcategory(category, addToIncome)
End If

'ask if the user would like to add more categories
answer = MsgBox("Add more categories?", vbYesNo, "Add More?")
If answer = vbNo Then
    Unload Me
End If

End Sub

'if user clicks cancel
Private Sub CancelButton_Click()
Unload Me

End Sub

'initialize userform
Private Sub UserForm_Initialize()
'make expense category button true to ensure a radio button is clicked
ExpenseRadio.value = True

End Sub
