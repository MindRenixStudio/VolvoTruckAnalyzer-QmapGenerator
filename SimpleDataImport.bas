Attribute VB_Name = "Module2"
Sub Click_ImportData()

Dim CurrentRowImport As Integer
Dim FilePacka As String
Dim FileFreight As String

Dim wsCopy As Worksheet
Dim wsDest As Worksheet
Dim lCopyLastRow As Long
Dim lDestLastRow As Long

'____________________________'

CurrentRowImport = 4

FilePacka = "C:\Users\" & Environ("username") & "\Desktop\p.csv"
FileFreight = "C:\Users\" & Environ("username") & "\Desktop\f.csv"
FileTriler = "C:\Users\" & Environ("username") & "\Desktop\t.xlsx"

'___Set variables for copy and destination sheets
Set wsCopy = Workbooks("p.csv").Worksheets("p")
Set wsDest = Workbooks("TrailerAnalyzer 2.2.1.0.xlsm").Worksheets("Input")
  
'___1. Find last used row in the copy range based on data in column A
lCopyLastRow = wsCopy.Cells(wsCopy.Rows.Count, "A").End(xlUp).Row
  
'___2. Find first blank row in the destination range based on data in column A
'___Offset property moves down 1 row
lDestLastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1).Row

'___3. Copy & Paste Data
wsCopy.Range("A2:BM" & lCopyLastRow).Copy _
  wsDest.Range("A" & lDestLastRow)
  
'___Set variables for copy and destination sheets
Set wsCopy = Workbooks("f.csv").Worksheets("f")
Set wsDest = Workbooks("TrailerAnalyzer 2.2.1.0.xlsm").Worksheets("Input")
  
'___1. Find last used row in the copy range based on data in column A
lCopyLastRow = wsCopy.Cells(wsCopy.Rows.Count, "A").End(xlUp).Row
  
'___2. Find first blank row in the destination range based on data in column A
'___Offset property moves down 1 row
lDestLastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1).Row

'___3. Copy & Paste Data
wsCopy.Range("A2:BM" & lCopyLastRow).Copy _
  wsDest.Range("A" & lDestLastRow)

'___Set variables for copy and destination sheets
Set wsCopy = Workbooks("t.xlsx").Worksheets("XD IN ")
Set wsDest = Workbooks("TrailerAnalyzer 2.2.1.0.xlsm").Worksheets("Planning")

'___3. Copy & Paste Data
wsCopy.Range("A6:A150").Copy _
  wsDest.Range("B2")

'___Set variables for copy and destination sheets
Set wsCopy = Workbooks("t.xlsx").Worksheets("XD IN ")
Set wsDest = Workbooks("TrailerAnalyzer 2.2.1.0.xlsm").Worksheets("Planning")

'___3. Copy & Paste Data
wsCopy.Range("I6:I150").Copy _
  wsDest.Range("C2")






























End Sub
