VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule_AdultSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pBookPath           As String
Private pBook               As Workbook
Private pUltArr()           As Variant
Private pFitArr()           As Variant
Private pRowSelector        As Integer
Private pTitleRanges()      As Range
Private pCoachRanges()      As Range
Private pMergedRanges()     As Range
Private pTimeRowRanges()    As Range
Private pTimeRanges()       As Range
Private pUltTitleRow1       As Integer
Private pUltTitleRow2       As Integer
Private pFitTitleRow1       As Integer
Private pFitTitleRow2       As Integer
Private pUltBottomRow       As Integer
Private pFitBottomRow       As Integer

Public Sub CreateSchedule(ByVal UltArr As Variant, ByVal FitArr As Variant)
    
    
    
    ' * ' DIM
    Dim RowSelect       As Integer
    Dim TimeRowItr, _
        RowItr1, _
        RowItr2, _
        ColItr1, _
        ColItr2, _
        FirstRow, _
        LastRow, _
        RowsUsed, _
        RowsPerTitle, _
        RowsPerCoach    As Integer
        
    
    Dim mTimeRow        As Schedule_TimeRow
    Dim mFirstCell, _
        mLastCell, _
        mTimeRange      As Range
    Dim sheet           As Worksheet
    
    
    ' * ' Set Arrays
    pUltArr = UltArr
    pFitArr = FitArr
    
    ' * ' Set Row Selector to 5
    pRowSelector = 6
    
    ' * ' Set Sheet
    Set sheet = pBook.Sheets("Ultimate")
    
    
    ' * ' Loop Through Each Time Row
    For TimeRowItr = 0 To UBound(pUltArr)
        Dim PasteArray()    As Variant
        
        ' * ' Set Time ROw
        Set mTimeRow = pUltArr(TimeRowItr)
        
        ' * ' Get Copy Friendly Array
        PasteArray = mTimeRow.ConvertToValidExcelArray
        
        ' * ' Populate Merged Ranges
        GetRanges sheet, PasteArray
        
        ' * ' Copy Values
        pTimeRowRanges(UBound(pTimeRowRanges) - 1).Value2 = PasteArray
        
        ' * ' Add Rows To pRowSelect
        pRowSelector = pRowSelector + UBound(PasteArray) + 1
        
        If PasteArray(0, 0) = "12:00 PM" Then
            pUltTitleRow1 = pRowSelector
        
        ElseIf pUltTitleRow1 <> 0 And pUltTitleRow2 = 0 Then
            pUltTitleRow2 = pRowSelector - 1
        End If
    Next
    
    pUltBottomRow = pRowSelector - 1
    
    ' * ' Set Row Selector to 5
    pRowSelector = 6
    
    ' * ' Set Sheet
    Set sheet = pBook.Sheets("Fitness")
    
    
    ' * ' Loop Through Each Time Row
    For TimeRowItr = 0 To UBound(pFitArr)
        
        ' * ' Set Time ROw
        Set mTimeRow = pFitArr(TimeRowItr)
        
        ' * ' Get Copy Friendly Array
        PasteArray = mTimeRow.ConvertToValidExcelArray
        
        ' * ' Populate Merged Ranges
        GetRanges sheet, PasteArray
        
        ' * ' Copy Values
        pTimeRowRanges(UBound(pTimeRowRanges) - 1).Value2 = PasteArray
        
        ' * ' Add Rows To pRowSelect
        pRowSelector = pRowSelector + UBound(PasteArray) + 1
        
        If PasteArray(0, 0) = "12:00 PM" Then
            pFitTitleRow1 = pRowSelector
        End If
        
    Next
    
    pFitBottomRow = pRowSelector - 1
    
    ' * ' Format ALL RANGES
    FormatRanges
    FormatSheets
    ' * ' Create Workbook
    BookTitle = MonthName(Month(DateTime.Now))
    pBookPath = g_objGWClient.SchedulePath & BookTitle & " Adult Schedule.pdf"
    
    pBook.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, FileName:=pBookPath
    
    pBook.Close True
End Sub
Private Sub FormatSheets()

    Dim sheet As Worksheet
    For Each sheet In pBook.Sheets
    
        With sheet
            .PageSetup.RightFooter = "Updated On: " & DateTime.Now
            .PageSetup.Zoom = False
            .PageSetup.Orientation = xlLandscape
            .PageSetup.TopMargin = Application.InchesToPoints(0.25)
            .PageSetup.LeftMargin = Application.InchesToPoints(0.25)
            .PageSetup.RightMargin = Application.InchesToPoints(0.25)
            .PageSetup.BottomMargin = Application.InchesToPoints(0.5)
            .PageSetup.FitToPagesTall = 1
            .PageSetup.FitToPagesWide = 1
            .Range("B1:H1").Columns.ColumnWidth = 26
        End With
        
        Select Case sheet.Name
            
            Case "Ultimate":
                With sheet
                    .Range("C1").Value = "Ultimate Schedule"
                    .Range("C3").Value = MonthName(Month(DateTime.Now))
                    .Range("C3").Font.Bold = True
                End With
                
                With sheet.Range(sheet.Cells(pUltTitleRow1, 7), sheet.Cells(pUltTitleRow2, 8))
                    .Merge
                    .Value = "Attention Members"
                    .Font.Bold = True
                    .Font.Size = 24
                    .Interior.Color = RGB(255, 255, 255)
                End With
                
                With sheet.Range(sheet.Cells(pUltTitleRow2 + 1, 7), sheet.Cells(pUltBottomRow, 8))
                    .Merge
                    .Font.Size = 18
                    .BorderAround xlContinuous, xlMedium
                    .Value = "To view the class schedule at any time, download our smart phone App, 'UFC GYM'. You'll find many helpful features that can make your daily workout routine even better. Scan In with your phone, get updated on gym specials, and rack up points for everytime you check-in. Then use those points to redeem select specials every month! Ask the Front Desk for more information!"
                    .WrapText = True
                    .Font.Italic = True
                    .Interior.Color = RGB(255, 255, 255)
                End With
                
                ' * ' Height Row
                With sheet
                    .Range(.Cells(6, 1), .Cells(pUltBottomRow, 1)).RowHeight = 780 / (pUltBottomRow - 5)
                End With
                
                
            Case "Fitness":
                 With sheet
                    .Range("C1").Value = "Fitness Schedule"
                    .Range("C3").Value = MonthName(Month(DateTime.Now))
                    .Range("C3").Font.Bold = True
                End With
                
                With sheet.Range(sheet.Cells(pFitTitleRow1, 7), sheet.Cells(pFitTitleRow1 + 1, 8))
                    .Merge
                    .Value = "Attention Members"
                    .Font.Bold = True
                    .Font.Size = 24
                    .Interior.Color = RGB(255, 255, 255)
                End With
                
                With sheet.Range(sheet.Cells(pFitTitleRow1 + 2, 7), sheet.Cells(pFitBottomRow, 8))
                    .Merge
                    .Font.Italic = True
                    .Font.Size = 18
                    .BorderAround xlContinuous, xlMedium
                    .Value = "To view the class schedule at any time, download our smart phone App, 'UFC GYM'. You'll find many helpful features that can make your daily workout routine even better. Scan In with your phone, get updated on gym specials, and rack up points for everytime you check-in. Then use those points to redeem select specials every month! Ask the Front Desk for more information!"
                    .WrapText = True
                    .Interior.Color = RGB(255, 255, 255)
                End With
                
                ' * ' Height Row
                With sheet
                    .Range(.Cells(6, 1), .Cells(pFitBottomRow, 1)).RowHeight = 780 / (pFitBottomRow - 5)
                End With
        End Select
    Next
End Sub
Private Sub FormatRanges()
    Dim Itr As Integer
    
    
    
    
    
    ' * ' Format Titles
    For Itr = 0 To UBound(pTitleRanges) - 1
        With pTitleRanges(Itr)
            .Merge
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Font.Size = 8
            .Font.Bold = True
        End With
    Next
    
    ' * ' Format Coaches
    For Itr = 0 To UBound(pCoachRanges) - 1
        With pCoachRanges(Itr)
            .Merge
            .Borders(xlEdgeBottom).LineStyle = xlDash
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Font.Size = 8
            .Font.Italic = True
        End With
    Next
    
    ' * ' Format TimeRange
    For Itr = 0 To UBound(pTimeRanges) - 1
        With pTimeRanges(Itr)
            .Merge
            .Font.Bold = True
            .BorderAround xlContinuous, xlMedium
        End With
    Next
    
    ' * ' Format Merged Rows
    For Itr = 0 To UBound(pMergedRanges) - 1
        With pMergedRanges(Itr)
            .Merge
            .Interior.Color = RGB(190, 190, 190)
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next
    
    
    ' * ' Format Time ROws
    For Itr = 0 To UBound(pTimeRowRanges) - 1
        With pTimeRowRanges(Itr)
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .BorderAround xlContinuous, xlMedium
            .WrapText = True
        End With
    Next
    
    
End Sub
Private Sub GetRanges(ByVal sheet As Worksheet, ByVal Arr As Variant)
    Dim TitleInt, _
        CoachInt            As Integer
        
    Dim RowItr1, _
        RowItr2, _
        RowItr3, _
        ColItr, _
        Itr                 As Integer
        
    Dim mTimeRowRange       As Range
    Dim mTimeRange          As Range
    
    ' * ' Time Range
    Set mTimeRange = sheet.Range(sheet.Cells(pRowSelector, 1), sheet.Cells(pRowSelector + UBound(Arr), 1))
    
    ' * ' Time Row Range
    Set mTimeRowRange = sheet.Range(sheet.Cells(pRowSelector, 1), sheet.Cells(pRowSelector + UBound(Arr), 8))
    
    
    ' * ' Add To Arrays
    Set pTimeRowRanges(UBound(pTimeRowRanges)) = mTimeRowRange
    Set pTimeRanges(UBound(pTimeRanges)) = mTimeRange
    
    ' * ' ReDim
    ReDim Preserve pTimeRowRanges(UBound(pTimeRowRanges) + 1)
    ReDim Preserve pTimeRanges(UBound(pTimeRanges) + 1)
    
    For ColItr = 1 To 7
        
        
        For RowItr1 = 0 To UBound(Arr, 1)
            ' * ' If First Val Of Col Is Empty Then Whole Column Is Empty
            If RowItr1 = 0 And IsEmpty(Arr(RowItr1, ColItr)) Then
                
                ' * ' EMPPPPTYYYYYYY RANGE
                Dim MergeRange As Range
                Set MergeRange = sheet.Cells(pRowSelector, ColItr + 1)
                
                Set MergeRange = MergeRange.Resize(UBound(Arr) + 1)
            
            
            
                ' * ' Add MergedRange To List
                Set pMergedRanges(UBound(pMergedRanges)) = MergeRange
                ReDim Preserve pMergedRanges(UBound(pMergedRanges) + 1)
                
                ' * ' Exit For
                Exit For
                             
            Else:
            
                Dim TitleRange, _
                    CoachRange      As Range
                
                
                ' * ' Set Title
                Set TitleRange = sheet.Cells(pRowSelector + RowItr1, ColItr + 1)
                
                ' * ' Add 1 To Row
                RowItr1 = RowItr1 + 1
                
                Do Until title
                        
                    Select Case IsEmpty(Arr(RowItr1, ColItr))
                        
                        Case True:
                            Set TitleRange = TitleRange.Resize(TitleRange.Rows.Count + 1)
                            RowItr1 = RowItr1 + 1
                            
                        Case False:
                            title = True
                            
                            Set CoachRange = sheet.Cells(pRowSelector + RowItr1, ColItr + 1)
                            
                            
                            Do Until coach
                            
                                If RowItr1 >= UBound(Arr, 1) Then
                                    coach = True
                                    RowItr1 = RowItr1 + 1
                                Else:
                                    RowItr1 = RowItr1 + 1
                                    Select Case IsEmpty(Arr(RowItr1, ColItr))
                                    
                                        Case True:
                                            Set CoachRange = CoachRange.Resize(CoachRange.Rows.Count + 1)
                                        
                                        Case False:
                                            coach = True
                                        
                                    End Select
                                End If
                            Loop
                        
                    End Select
                        
                Loop
                
                ' * ' Add Class Ranges
                Set pTitleRanges(UBound(pTitleRanges)) = TitleRange
                Set pCoachRanges(UBound(pTitleRanges)) = CoachRange
                
                ReDim Preserve pTitleRanges(UBound(pTitleRanges) + 1)
                ReDim Preserve pCoachRanges(UBound(pCoachRanges) + 1)
                
                title = False
                coach = False
                
                RowItr1 = RowItr1 - 1
            End If
            
        Next
        
    Next
                                         
            
    
End Sub
Private Sub Class_Initialize()
    ' * ' DIM
    Dim MainBook        As Workbook
    Dim NewBook         As Workbook
    Dim BookTitle       As String
    
    ' * ' Set MainBook to ActiveBook For COpying Purposes
    Set MainBook = ThisWorkbook
    
    ' * ' Create Workbook
    BookTitle = MonthName(Month(DateTime.Now))
    pBookPath = g_objGWClient.SchedulePath & "Excel Files\" & BookTitle & " Adult Schedule.xlsx"
    
    Application.DisplayAlerts = False
    
    Set NewBook = Workbooks.Add
        
        With NewBook
            .title = BookTitle
            .Subject = "Class Schedule"
            .SaveAs FileName:=pBookPath
        End With
        
        
        
    ' * ' Add Sheets
    MainBook.Sheets("AdultScheduleTemp").Copy Before:=Workbooks(NewBook.Name).Sheets(1)
    MainBook.Sheets("AdultScheduleTemp").Copy Before:=Workbooks(NewBook.Name).Sheets(1)
    
    
    ' * ' Rename
    NewBook.Sheets(1).Name = "Ultimate"
    NewBook.Sheets(2).Name = "Fitness"
    NewBook.Sheets(3).Delete
    
    Application.DisplayAlerts = True
    
    Set pBook = NewBook
    
    ' * ' Dim Ranges Arrays
    ReDim pTitleRanges(0)
    ReDim pCoachRanges(0)
    ReDim pMergedRanges(0)
    ReDim pTimeRanges(0)
    ReDim pTimeRowRanges(0)
    
    pUltLegendRow1 = 0
    pUltLegendRow2 = 0
    pFitLegendRow1 = 0
    pFitLegendRow2 = 0
    
End Sub
