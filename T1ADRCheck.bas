Attribute VB_Name = "Module4"
Sub Click_Write_ADR()

Dim InputRowCountADR As Integer
Dim InputRowCountT1 As Integer
Dim InputRowCountTot As Integer

'Counter of ADR/T1/Total
InputRowCountADR = 0
InputRowCountT1 = 0
InputRowCountTot = 0

For x = 15 To 44
    If IsEmpty(Worksheets("Planning").Cells(x, 6)) = False Then
        InputRowCountADR = InputRowCountADR + 1
    End If
Next

For x = 15 To 44
    If IsEmpty(Worksheets("Planning").Cells(x, 7)) = False Then
        InputRowCountT1 = InputRowCountT1 + 1
    End If
Next

InputRowCountTot = Worksheets("Input").Range("A4").End(xlDown).Row

'______________________________'

For c = 15 To InputRowCountADR + 15

    For v = 4 To InputRowCountTot
        
        'MsgBox (Worksheets("Planning").Cells(c, 6) & " /" & Worksheets("Input").Cells(v, 3))
        
        If Worksheets("Planning").Cells(c, 6) = Worksheets("Input").Cells(v, 3) Then
        Worksheets("Input").Cells(v, 29).Value = "Y"
        End If
    
    Next

Next

For c = 15 To InputRowCountT1 + 15

    For v = 4 To InputRowCountTot
        
        'MsgBox (Worksheets("Planning").Cells(c, 6) & " /" & Worksheets("Input").Cells(v, 3))
        
        If Worksheets("Planning").Cells(c, 7) = Worksheets("Input").Cells(v, 3) Then
        Worksheets("Input").Cells(v, 29).Value = "T"
        End If
    
    Next

Next


End Sub
