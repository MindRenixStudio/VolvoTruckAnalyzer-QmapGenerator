Attribute VB_Name = "Module3"
 Sub Click_Generate_Prepare_Sheet()

    Dim InputCounter As Integer
    Dim TemplateCounter As Integer
    
    Dim InputNow As String
    Dim SortDate As String
        
    Dim InputRowCount As Integer
    
    Dim csvName As String
    
    TemplateCounter = 2
    
    '_____________________________________
    
    InputRowCount = Worksheets("Input").Range("A4").End(xlDown).Row
    
    Worksheets("Day_Prepare").Copy After:=Worksheets("Input")
    Sheets("Day_Prepare (2)").Name = "LFS_Upload"
    
    For i = 5 To InputRowCount
        
        InputNow = Worksheets("Input").Cells(i, 49)
        SortDate = Worksheets("Planning").Cells(2, 9)
        
        If InputNow = SortDate Then
                        
            For p = 1 To 66
                
                Worksheets("LFS_Upload").Cells(TemplateCounter, p) = Worksheets("Input").Cells(i, p)
                
                Worksheets("LFS_Upload").Cells(TemplateCounter, 39).Value = "Load collected"
                
                If Worksheets("LFS_Upload").Cells(TemplateCounter, 29) = "T" Then
                
                    Worksheets("LFS_Upload").Cells(TemplateCounter, 29) = "N"
                    
                End If
                                
            Next
            
            TemplateCounter = TemplateCounter + 1
            
        End If
    
    Next
    
    csvName = "LFS_Upload.csv"
    
    strFullName = "C:\Users\" & Environ("username") & "\Desktop\LFS_CSV\" & csvName
    ActiveWorkbook.SaveAs Filename:=strFullName, FileFormat:=xlCSV, CreateBackup:=False, local:=True

End Sub
