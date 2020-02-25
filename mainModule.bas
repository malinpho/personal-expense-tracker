Attribute VB_Name = "mainModule"
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
' NOTE
' the Calendar userform is NOT my code
' could not implement a microsoft excel chooser properly
' tried Date and Time Picker but it didn't work
' so decided to find one online
Public cn As New ADODB.connection
Public rs As New ADODB.Recordset

Public selectedFile As String

'used to get name of database file

Sub getFile()
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .Show
    .InitialFileName = ThisWorkbook.Path
End With

On Error GoTo databaseNotFound

selectedFile = fd.SelectedItems(1)

'prompt user to select a file to use
''selectedFile = Application.GetOpenFilename

With cn
    .ConnectionString = "Data Source=" & selectedFile
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open
    'test connection?
    .Close
End With
Exit Sub

'if database not successfully opened, shows message
databaseNotFound:
    MsgBox "Database cannot be found or another error has occurred.", vbCritical
    Err.Clear
End Sub

