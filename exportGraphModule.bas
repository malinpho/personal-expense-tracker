Attribute VB_Name = "exportGraphModule"
Option Explicit
' +===================== CP212 Windows Application Programming ===============+
' Name: Malin Pho
' Student ID: 150858480
' Date: December 7, 2016
' Program title: Personal Expense Tracker
' Description: Uses a database to track personal expenses
' +===========================================================================+
' NEW BUTTON: ADD CHART TO WORD DOCUMENT
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
' +=================================================================================================================================+
' This only exports a chart to a word document, it does not provide any more information
' I know more is required, but could not figure it out
'main program to run export of summary

Public ChartObj As ChartObject
Public categories() As Variant
Public amount As Currency
Public categorycount As Integer
Sub exportSummaryMain()

Dim SQL As String

If selectedFile = "" Then
    Call getFile
End If

With cn
    .ConnectionString = "Data Source=" & selectedFile
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open
    'proceeds with program given there is no error
End With

'call all these helper functions
totalAllExpenses
arraytoSheet
createChart
exportToWord

'delete sheet with values used to make chart
Application.DisplayAlerts = False

Worksheets("Temp Chart Data").Delete

Application.DisplayAlerts = True

'close connection
cn.Close

End Sub

'fill category array with categories from database
Sub totalAllExpenses()

Dim SQL As String
Dim i As Integer
Dim categoryString As String
categorycount = 0

SQL = "SELECT ExpenseCategory from ExpenseCategories"

With rs
    .Open SQL, cn
    Do Until .EOF
        categorycount = categorycount + 1
        .MoveNext
    Loop
    .Close
End With

ReDim categories(0 To categorycount - 1, 0 To 1)

i = 0
With rs
    .Open SQL, cn
    Do Until .EOF
        categories(i, 0) = .Fields("ExpenseCategory")
        i = i + 1
        .MoveNext
    Loop
    .Close
End With

For i = 0 To categorycount - 1
    categoryString = categories(i, 0)
    'subroutine below finds all the transactions for the categoryString and adds or subtracts to find total
    totalExpenses (categoryString)
    'public variable 'amount' is put in for the category amount for that particular category
    categories(i, 1) = amount
Next

End Sub

'puts values of categories and money spent on them on new sheet
Sub arraytoSheet()
Dim i As Integer
Dim sheet As Worksheet

'create new sheet
ThisWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.name = "Temp Chart Data"

'fill in sheet with categories and their amounts
With Range("A1")
    .value = categories(0, 0)
    .Offset(0, 1).value = categories(0, 1)
    For i = 0 To categorycount - 2
        .Offset(i + 1, 0).value = categories(i + 1, 0)
        .Offset(i + 1, 1).value = categories(i + 1, 1)
    Next
    .Offset(i + 1, 0).value = categories(i, 0)
    .Offset(i + 1, 1).value = categories(i, 1)
End With

End Sub

'actually create the chart
Sub createChart()
Dim sourcedata As Range

'add chartobj to sheet
Set ChartObj = Sheets("Temp Chart Data").ChartObjects.Add(Left:=300, Top:=0, Width:=500, Height:=300)

Sheets("Temp Chart Data").Range(Range("A1"), Range("B1").End(xlDown).Offset(-1, 0)).Select
Set sourcedata = Selection

'set data to chart and define the chart typer
ChartObj.chart.SetSourceData Source:=sourcedata, PlotBy:=xlColumns

'edit chart
With ChartObj.chart
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Expense Category Distribution"
    
End With

End Sub

'total all expenses of a certain category
Sub totalExpenses(categoryString As String)

Dim SQL As String

amount = 0
SQL = "SELECT Amount FROM Expenses WHERE Category='" & categoryString & "'"

With rs
    .Open SQL, cn
    Do Until .EOF
        amount = amount + .Fields("Amount")
        .MoveNext
    Loop
    .Close
End With

End Sub

'export graph to word as a picture
Sub exportToWord()

Dim wordApp As Word.Application
Dim wordDoc As Word.Document

Application.ScreenUpdating = False

On Error Resume Next
    Set wordApp = GetObject("word.application")
    
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.application")
    End If

wordApp.Activate

'add new document
Set wordDoc = wordApp.Documents.Add

'let user know there was a word document made
MsgBox ("Word Document with a chart of the expenses you've incurred was made. Save this document if you would like to keep it.")

ChartObj.CopyPicture

wordDoc.Paragraphs(1).Range.Paste

wordApp.Visible = True

Application.ScreenUpdating = True
Application.CutCopyMode = False

End Sub
