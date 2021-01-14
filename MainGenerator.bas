Attribute VB_Name = "Module1"
Sub Click_Generate_From_Planning()
    
    'Setting up variables for Splitting LN from 'Planning'
    Dim WrdArray() As String
    Dim SplittedString As String
    
    'Setting up variables for IF statements and comparison
    Dim InputText As String
    Dim PlanningText As String
    Dim InputText2 As String
    
    'Temp variable for debugging what is now splitted (LN) in 'Planning'
    Dim SplittedNow As String
    
    'Setting up basic counter for rows
    Dim InputCount As Integer
    Dim InputRowCount As Integer
    Dim TrailerSheetCounter As Integer
    Dim InternalLoadSearchCounter As Integer
    Dim SortCount As Integer
    
    'Setting up counter in TrailerList_Template that refers rows
    Dim CheckListCounter As Integer
    
    'Setting up counter that count filled rows in TrailerList
    Dim FilledColumns As Integer
    
    'Setting up variable for notification, when load number was not found
    Dim LoadExist As Integer
    
    Dim CheckTO As String
    
    Dim CheckTOControl As String
    
    Dim CheckTOCounter As Integer
    
    Dim Week As Integer
    Dim Day As Integer
    
    '________________________________'
    
    'Setting up default values for counter
    '4
    InputCount = 4
    '6
    TrailerSheetCounter = 6
    '4
    InternalLoadSearchCounter = 4
    
    CheckTOCounter = 1
    
    'Count for getting the line that is sorted and waiting to be replicated to CheckList with Color and Zone
    SortCount = 1
    CheckListCounter = 6
    
    'Count lines in 'Input' to get dynamic row count
    InputRowCount = Worksheets("Input").Range("A4").End(xlDown).Row
            
    Week = Worksheets("Planning").Cells(10, 7)
    Day = Worksheets("Planning").Cells(10, 9)
        
    'Looking for rows in "Planning" | QMAP/TIMESLOT Counter
    For i = 2 To 151
        'Adding to array row that will be splitted
        SplittedString = Worksheets("Planning").Cells(i, 2)
        'Splitting array
        WrdArray() = Split(SplittedString, "/")
        
        Worksheets("TrailerList_Template").Copy After:=Worksheets("Planning")
        Sheets("TrailerList_Template (2)").Name = "TrailerList_Printable"

        TrailerSheetCounter = 6
        CheckListCounter = 6
        
        'FilledColumns = 0
        
        'Looking in array for single splitted LN
        For o = LBound(WrdArray) To UBound(WrdArray)
            SplittedNow = vbNewLine & WrdArray(o)
            'MsgBox "Splitted now" & SplittedNow
                
            'Reseting the counters for next looping
            InputCount = 4
            'TrailerSheetCounter = 6
            InternalLoadSearchCounter = 4
            'CheckListCounter = 6
            
            'FilledColumns = 0
            
            LoadExist = 0
                      
            'Looking and comparing data from 'Planning' and 'Input' sheets
            For x = 0 To InputRowCount
            
                InputText = Worksheets("Input").Cells(InputCount, 1)
                PlanningText = WrdArray(o)
                                
                'Line to see and debug what is comparing
                'MsgBox PlanningText & " / " & InputText
                
                If InputText = PlanningText Then
                    
                    'HERE JUST TO LIMIT CREATING WHEN NEW QMAP --------------------------------------------------------------------------
                    
                    'Creating TrailerList_Printable sheet
                    '------------ BUGGED HERE CS OF NEW NAME OF TRAILERLIST_PRINTABLE (LN)
                    
                    'Worksheets("TrailerList_Template").Copy After:=Worksheets("Planning")
                    'Sheets("TrailerList_Template (2)").Name = "TrailerList_Printable"
                    
                    'Filling data to constant cells
                    'TrailerPlate
                    Worksheets("TrailerList_Printable").Cells(2, 8) = Worksheets("Input").Cells(InputCount, 57)
                    'Carrier
                    Worksheets("TrailerList_Printable").Cells(4, 2) = Worksheets("Input").Cells(InputCount, 40)
                    'Qmap
                    Worksheets("TrailerList_Printable").Cells(2, 3) = Worksheets("Planning").Cells(i, 1)
                    'Timeslot
                    Worksheets("TrailerList_Printable").Cells(4, 3) = Worksheets("Planning").Cells(i, 3)
                    'Load Number
                    Worksheets("TrailerList_Printable").Cells(33, 3) = Worksheets("Input").Cells(InputCount, 1)
                    'Name
                    Worksheets("TrailerList_Printable").Cells(7, 12) = Worksheets("Planning").Cells(2, 6)
                    'Date
                    Worksheets("TrailerList_Printable").Cells(2, 2) = "W" & Week & "D" & Day
                                        
                    'Basically the same for cycle as parent except this fills up data to dynamic cells
                    For c = 0 To InputRowCount
                    
                        InputText2 = Worksheets("Input").Cells(InternalLoadSearchCounter, 1)
                        
                        If InputText2 = PlanningText Then
                            
                            'Supplier
                            Worksheets("TrailerList_Printable").Cells(TrailerSheetCounter, 2) = Worksheets("Input").Cells(InternalLoadSearchCounter, 6)
                            'TO
                            Worksheets("TrailerList_Printable").Cells(TrailerSheetCounter, 3) = Worksheets("Input").Cells(InternalLoadSearchCounter, 3)
                            'FDP
                            Worksheets("TrailerList_Printable").Cells(TrailerSheetCounter, 4) = Worksheets("Input").Cells(InternalLoadSearchCounter, 13)
                            'Colli_INET
                            Worksheets("TrailerList_Printable").Cells(TrailerSheetCounter, 5) = Worksheets("Input").Cells(InternalLoadSearchCounter, 26)
                            'Hidden country
                            Worksheets("TrailerList_Printable").Cells(TrailerSheetCounter, 10) = Worksheets("Input").Cells(InternalLoadSearchCounter, 9)
                            
                            'Up counters
                            InternalLoadSearchCounter = InternalLoadSearchCounter + 1
                            TrailerSheetCounter = TrailerSheetCounter + 1
                            
                            LoadExist = 1
                                                    
                        ElseIf InputText2 <> PlanningText Then
                        
                            'Up counter
                            InternalLoadSearchCounter = InternalLoadSearchCounter + 1
                            
                        End If
                    
                    Next
                    
                    CheckTOCounter = 0
            
                    For r = 4 To InputRowCount
                    
                        CheckTO = Worksheets("TrailerList_Printable").Cells(6, 3)
                        CheckTOControl = Worksheets("Input").Cells(r, 3)
                        
                        If CheckTO = CheckTOControl Then
                            If Worksheets("Input").Cells(r, 29) = "T" Then
                                Worksheets("TrailerList_Printable").Cells(2, 7) = "T1"
                                Worksheets("TrailerList_Printable").Cells(2, 7).Interior.Color = RGB(255, 57, 57)
                                CheckTOCounter = CheckTOCounter + 1
                            End If
                            If Worksheets("Input").Cells(r, 9) = "PL" Then
                                Worksheets("TrailerList_Printable").Cells(3, 8) = "OUT:"
                                Worksheets("TrailerList_Printable").Cells(4, 8) = Worksheets("Planning").Cells(23, 9)
                            End If
                        End If
                        
                        CheckTOCounter = CheckTOCounter + 1
                        
                    Next
                                        
                    'Adding TrailerSheetCounter because it was used as a sub-counter of InputCount to get exact line where it ended filling dynamic data
                    InputCount = InputCount + TrailerSheetCounter
                                                                                
                    '----------------------------------------------------------------------------------------------------
                    'SETUP CHECKLIST PAGE
                    
                    
                    
                    
                    
                    '----------------------------------------------------------------------------------------------------
                    
                    'Set sheet color to green
                    'Sheets("TrailerList_Printable").Tab.ColorIndex = 4
                    
                    'Rename sheet
                    'Sheets("TrailerList_Printable").Name = Worksheets("TrailerList_Printable").Cells(33, 3)
                    
                    InputCount = InputCount + 1
                    
                ElseIf InputText <> PlanningText Then
                    InputCount = InputCount + 1
                    
                End If
          
            Next
            FilledColumns = 0
            'Cycle that counts filled cells in TrailerList_Printable
            For Z = 6 To 30
                If IsEmpty(Worksheets("TrailerList_Printable").Cells(Z, 2)) = False Then
                    FilledColumns = FilledColumns + 1
                End If
            Next
            
            'MsgBox (FilledColumns)
            
            'Cycle through TrailerList_Printable
            
            
            
              
            'MsgBox ("LoadExist: " & LoadExist & "InputRowCount" & InputRowCount)
              
            'LoadExist = 2
            
            'If x > InputRowCount And LoadExist = 2 Then
            '    MsgBox ("Load not found" & SplittedNow)
            'End If
          
        Next
        
        For k = 1 To FilledColumns
            'Counter reset
            SortCount = 1
            
            'Creating CheckList_Printable
            Worksheets("CheckList_Template").Copy After:=Worksheets("TrailerList_Printable")
            Sheets("CheckList_Template (2)").Name = "CheckList_Printable"
            
            'Stamp
            Worksheets("CheckList_Printable").Cells(16, 4) = Worksheets("TrailerList_Printable").Cells(7, 12)
            'Q-NR
            Worksheets("CheckList_Printable").Cells(7, 8) = Worksheets("TrailerList_Printable").Cells(2, 3)
            'Supplier
            Worksheets("CheckList_Printable").Cells(14, 4) = Worksheets("TrailerList_Printable").Cells(CheckListCounter, 2)
            'INET
            Worksheets("CheckList_Printable").Cells(14, 10) = Worksheets("TrailerList_Printable").Cells(CheckListCounter, 5)
            'TO
            Worksheets("CheckList_Printable").Cells(7, 2) = Worksheets("TrailerList_Printable").Cells(CheckListCounter, 3)
            'FDP
            Worksheets("CheckList_Printable").Cells(29, 6) = Worksheets("TrailerList_Printable").Cells(CheckListCounter, 4)
            'Country
            Worksheets("CheckList_Printable").Cells(18, 6) = Worksheets("TrailerList_Printable").Cells(CheckListCounter, 10)
            
            If Worksheets("CheckList_Printable").Cells(18, 6) = "PL" Then
                'PL OUT
                Worksheets("CheckList_Printable").Cells(32, 7) = "OUT: " & Worksheets("TrailerList_Printable").Cells(4, 8)
            End If
                        
            'ADR/T1?
            
            CheckTOCounter = 0
            
            For r = 4 To InputRowCount
            
                CheckTO = Worksheets("CheckList_Printable").Cells(7, 2)
                CheckTOControl = Worksheets("Input").Cells(r, 3)
                
                If CheckTO = CheckTOControl Then
                    If Worksheets("Input").Cells(r, 29) = "Y" Then
                        Worksheets("CheckList_Printable").Cells(7, 6) = "ADR"
                        Worksheets("CheckList_Printable").Cells(7, 6).Interior.Color = RGB(218, 99, 0)
                    ElseIf Worksheets("Input").Cells(r, 29) = "T" Then
                        Worksheets("CheckList_Printable").Cells(7, 6) = "T1"
                        Worksheets("CheckList_Printable").Cells(32, 7).Value = "CUSTOM GOODS"
                        Worksheets("CheckList_Printable").Cells(7, 6).Interior.Color = RGB(255, 57, 57)
                        Worksheets("CheckList_Printable").Cells(2, 2).Interior.Color = RGB(255, 57, 57)
                        Worksheets("CheckList_Printable").Cells(4, 2).Interior.Color = RGB(255, 57, 57)
                        Worksheets("CheckList_Printable").Cells(2, 2).Interior.Color = RGB(255, 57, 57)
                        Worksheets("CheckList_Printable").Cells(49, 2).Interior.Color = RGB(255, 57, 57)
                        CheckTOCounter = CheckTOCounter + 1
                    End If
                End If
                
                CheckTOCounter = CheckTOCounter + 1
                
            Next
                        
            For h = 1 To 70
                FDP = Worksheets("CheckList_Printable").Cells(29, 6)
                ControlFDP = Worksheets("SortingSheet").Cells(SortCount, 2)
                
                If FDP = ControlFDP Then
                
                    'Zone
                    Worksheets("CheckList_Printable").Cells(21, 6) = Worksheets("SortingSheet").Cells(h, 3)
                    
                    'Color (Not proud of that)
                    If Worksheets("SortingSheet").Cells(h, 4).Value = "Dark blue" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 25
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Dark purple" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 21
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Orange" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 46
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Dark gree" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 10
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Magenta" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 26
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Light Green" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 4
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Cyan" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 8
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Red" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 3
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Peach orange" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 45
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Grass green" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 50
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Pea green" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 43
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "White" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 2
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Lavender" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 39
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Sky blue" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.ColorIndex = 33
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Dark red" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.Color = RGB(176, 0, 0)
                                                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Dark dark green" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.Color = RGB(0, 122, 20)
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Black" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.Color = RGB(0, 0, 0)
                    
                    ElseIf Worksheets("SortingSheet").Cells(h, 4).Value = "Mustard" Then
                    Worksheets("CheckList_Printable").Cells(18, 2).Interior.Color = RGB(230, 170, 0)
                                                    
                    End If
                                                        
                End If
                                    
                SortCount = SortCount + 1
                                    
            Next

            CheckListCounter = CheckListCounter + 1
            
            'Immidietly after creating a fillin a sheet, print it out           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Worksheets("Planning").Cells(2, 8) = "YES" = True Then
                Sheets("CheckList_Printable").PrintOut
            End If
            
            'Set sheet color to green
            Sheets("CheckList_Printable").Tab.ColorIndex = 4
            
            'Rename sheet
            Sheets("CheckList_Printable").Name = Worksheets("CheckList_Printable").Cells(7, 2)
                
            Next
        
        'Immidietly after creating a fillin a sheet, print it out             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Worksheets("Planning").Cells(2, 8) = "YES" Then
            Sheets("TrailerList_Printable").PrintOut
        End If
        
        Sheets("TrailerList_Printable").Tab.ColorIndex = 4
            
        'Rename sheet
        Sheets("TrailerList_Printable").Name = Worksheets("TrailerList_Printable").Cells(33, 3)
                    
    Next
    

End Sub
