VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule_YouthSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pBookPath           As String
Private pBook               As Workbook
Private pYouthArr()         As Variant
Private pRowSelector        As Integer
Private pTitleRanges()      As Range
Private pCoachRanges()      As Range
Private pMergedRanges()     As Range
Private pTimeRowRanges()    As Range
Private pTimeRanges()       As Range
Private pYouthBottomRow     As Integer
Private pWeekDayClasses     As Boolean
Private pFirstMondayClass   As Integer
Private pLastWeekdayClass   As Integer
Private pLastRow            As Integer
Private pEveningRow         As Integer


Public Sub CreateSchedule(ByVal YouthArr As Variant)
    
    Application.ScreenUpdating = False
    
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
    pYouthArr = YouthArr
    
    ' * ' Set Row Selector to 5
    pRowSelector = 6
    pEveningRow = 0
    
    ' * ' Set Sheet
    Set sheet = pBook.Sheets("Youth")
    
    
    ' * ' Loop Through Each Time Row
    For TimeRowItr = 0 To UBound(pYouthArr)
        Dim PasteArray()    As Variant
        
        ' * ' Set Time ROw
        Set mTimeRow = pYouthArr(TimeRowItr)
        
        If mTimeRow.ExcelArray(0) <= TimeSerial(12, 0, 0) And mTimeRow.NotEmpty Then
            pWeekDayClasses = True
        End If
        
        ' * ' Get Copy Friendly Array
        PasteArray = mTimeRow.ConvertToValidExcelArray
        
        ' * ' Populate Merged Ranges
        GetRanges sheet, PasteArray
        
        ' * ' Copy Values
        pTimeRowRanges(UBound(pTimeRowRanges) - 1).Value2 = PasteArray
        
        ' ' Check For Weekday Mornings
        If TimeRowItr = 1 Then
            pFirstMondayClass = pRowSelector
        End If
        
        ' * ' Add Rows To pRowSelect
        pRowSelector = pRowSelector + UBound(PasteArray) + 1
        
        If TimeValue(PasteArray(0, 0)) < TimeValue("12:00 PM") Then
            pLastWeekdayClass = pRowSelector
        End If
        
        If TimeValue(PasteArray(0, 0)) >= TimeValue("12:00 PM") And pEveningRow = 0 Then
            pEveningRow = pRowSelector
        End If
        
    Next
    
    pYouthBottomRow = pRowSelector
    ' * ' Format ALL RANGES
    FormatRanges
    FormatSheets
    
    ' * ' Create Workbook
    BookTitle = MonthName(Month(DateTime.Now))
    pBookPath = g_objGWClient.SchedulePath & BookTitle & " Youth Schedule.pdf"
    
    pBook.ExportAsFixedFormat xlTypePDF, FileName:=pBookPath, Quality:=xlQualityStandard
    
    
    pBook.Close True
    
    Application.ScreenUpdating = True
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
            
            Case "Youth":
                With sheet
                    .Range("C1").Value = "Youth Schedule"
                    .Range("C3").Value = MonthName(Month(DateTime.Now))
                    .Range("C3").Font.Bold = True
                End With
                
                If Not pWeekDayClasses Then
                    With sheet.Range(sheet.Cells(pFirstMondayClass, 3), sheet.Cells(pLastWeekdayClass - 1, 5))
                        .Merge
                        .Value = "No Classes"
                        .Interior.Color = RGB(255, 255, 255)
                        .BorderAround xlContinuous, xlMedium
                        .Font.Size = 72
                        .Font.Bold = True
                    End With
                End If
                
                
                ' * ' Un Merge Announcment Area
                With sheet.Range(sheet.Cells(pEveningRow, 7), sheet.Cells(pYouthBottomRow - 1, 8))
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                    .BorderAround xlContinuous, xlMedium
                    .Interior.Color = RGB(255, 255, 255)
                    .UnMerge
                End With
                
                Dim y As Integer
                
                For y = pEveningRow To pYouthBottomRow
                    
                    sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8)).Merge
                    
                    Select Case y - pEveningRow
                        
                        Case 1:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "Attention Parents!"
                                .Font.Underline = True
                                .Font.Bold = True
                                .Font.Size = 20
                            End With
                            
                        Case 2
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "***All students must be registered for all classs***"
                                .Font.Size = 14
                                .Font.Bold = True
                                .Font.Underline = True
                            End With
                            
                        Case 3:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "***ALL Students taking class, must be wearing UFC Apparel***"
                                .Font.Bold = True
                                .Font.Size = 14
                            End With
                        
                        Case 4:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "All Children Must Be Signed in At the Kid's Club Before Attending Class"
                                .Font.Size = 11
                            End With
                        
                        Case 5:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "Students that take BJJ (GI) Classes must have a UFC GI."
                                .Font.Size = 11
                            End With
                        
                        Case 6:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y + 1, 8))
                                .Merge
                                .Value = "***Students that are interested in taking our Advanced BJJ/Competition Class MUST SEE HEAD YOUTH BJJ INSTRUCTOR***"
                                .Font.Bold = True
                                .Font.Underline = True
                                .Font.Size = 14
                            End With
                            
                            y = y + 1
                        
                        Case 8:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "Students that take Striking Classes are required to have Hand Wraps, Gloves, workout shorts/leggings, and any UFC Gym T-Shirt."
                                .Font.Size = 9
                            End With
                            
                        Case 9:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "Students taking DUT/SAQ must wear shoes and UFC Athletic Wear."
                                .Font.Size = 9
                            End With
                            
                        Case 10:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "For MMA Fight Fit, Striking, And NO GI, we request that our students be in their UFC GYM Torrance Shirts. Shorts should not have pockets, as fingers can get caught and injured."
                                .Font.Size = 9
                            End With
                            
                        Case 11:
                            With sheet.Range(sheet.Cells(y, 7), sheet.Cells(y, 8))
                                .Value = "*SAQ(Speed, Agility, and Quickness) Will be scheduled under Youth DUT."
                                .Font.Size = 11
                            End With
                    End Select
                Next
                
                ' * ' Height Row
                If Not pWeekDayClasses Then
                    With sheet.Range(sheet.Cells(6, 2), sheet.Cells(pEveningRow, 6))
                        .Interior.Color = RGB(255, 255, 255)
                    End With
                End If
                
                With sheet
                    .Range(.Cells(6, 1), .Cells(pYouthBottomRow, 1)).RowHeight = 780 / (pYouthBottomRow - 5)
                End With
                
                ' * ' Format Back Sheet
                FormatYouthDescription pBook
                
            
        End Select
    Next
End Sub

Private Sub FormatYouthDescription(ByVal book As Workbook)
    Dim sheet As Worksheet
    
    
    With book
        Set sheet = .Sheets.Add
    End With
    
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
        
    With sheet
        .Name = "Youth Description"
        
        Dim x As Integer
        Dim bulletarr As Variant
        bulletarr = Array("o   Effective Self Defense", "o   Self Confidence", "o   Physical Fitness", "o   Discipline & Respect", "o   Balance & Coordination", "o   Focused Energy")
        
        For x = 1 To 16
        
            With sheet
                Select Case x
                    Case 1:
                        .Rows(x).RowHeight = 40
                        With .Cells(x, 1)
                            .Value = "Youth Program"
                            .Font.Bold = True
                            .Font.Size = 22
                            .VerticalAlignment = xlCenter
                            .Range("A" & x & ":B" & x).Merge
                            
                            .HorizontalAlignment = xlCenter
                            .WrapText = True
                        End With
                        
                    Case 2:
                    
                        .Rows(x).RowHeight = 15.75
                        .Range("A" & x & ":B" & x).Merge
                        With .Cells(x, 1)
                            .Font.Size = 9
                            .Value = "Ages 4 - 12"
                            .HorizontalAlignment = xlCenter
                            .Range("A" & x & ":B" & x).Merge
                            .WrapText = True
                        End With
                        
                    Case 3:
                        .Rows(x).RowHeight = 56.25
                        .Range("A" & x & ":B" & x).Merge
                        With .Cells(x, 1)
                            .Value = "Our UFC Kids and Youth Programs Incorporate personal empowerment, bully and stranger awareness, out and about safety, and emergency safety situations. Our kids & Youth Martial Arts system is a well-rounded and complete program that is taught in a rotating curriculum so that you can follow your child's progress and know exactly what their child is being taught before class. The techniques and philosophies build strong character, strong bodies, and greater respect for parents, teachers and others. We have incorporated a sports specific training program to enhance your young athlete's performance in any arena"
                            .Font.Size = 9
                            .WrapText = True
                        End With
                        
                    Case 4:
                        .Rows(x).RowHeight = 12.75
                        .Range("A" & x & ":B" & x).Merge
                        With .Cells(x, 1)
                            
                            .Value = "We use age specific methods to teach your children realistic awareness drills in a safe fun and interactive environment that emphasizes:"
                            .Font.Size = 9
                            .WrapText = True
                        End With
                        
                    Case 5, 6, 7, 8, 9, 10:
                        .Rows(x).RowHeight = 12
                        .Range("A" & x & ":B" & x).Merge
                        With .Cells(x, 1)
                            .Value = bulletarr(x - 5)
                            .WrapText = True
                            .Font.Size = 9
                        End With
                        
                    Case 11:
                        .Rows(x).RowHeight = 58.5
                        With .Cells(x, 1)
                            .Value = "Youth Brazilian Jiu-Jitsu - Our Brazilian Jiu-Jitsu focuses on basic foundation techniques along with proper application and position for joint locks and chokes. In a Brazilian Jiu-Jitsu class your child will learn how to defend or attack an opponent on the ground."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(x, 2)
                            .Value = "Competition BJJ(Advance Students)- Advance students will master the fundamental techniques and sequences, learning Takedowns, Submissions, Submission Escapes, Combo Attacks, Combo Escapes and BJJ technique drills."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        
                    Case 12, 14:
                        .Rows(x).RowHeight = 75
                         With .Cells(12, 1)
                            .Value = "Youth Striking - Youth Striking class teaches basic Kickboxing, Muay Thai, and western-Style Boxing skills and combinations. Youth Kickboxing classes include: cardio calisthenics, bag work, which includes punch, kick, and elbow strikes, partner drills, and pad work in addition to core strengthening exercises."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(12, 2)
                            .Value = "Youth Boxing-Youth Boxing class teaches basic boxing skills and techniques such as; stance, guard, movement, jab, cross, hook. It increases strength and improves endurance. Youth Boxing class will include cardio calisthenics, shadow boxing, bag work, partner drills, pad work and core strengthening exercises. These skills will give kids self-defense tools as well as self-confidence while keeping them physically fit. The 150 lb. heavy bags are the primary focus of the class."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(14, 1)
                            .Value = "Youth MMA - This class blends the hand skills of western boxing, the ground skills of Brazilian Jiu-Jitsu, takedowns and defense of Judo, Wrestling, and Muay Thai Kickboxing to broaden your child's skills."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(14, 2)
                            .Value = "SAQ- (Speed, Agility, and Quickness)- Whether you're training for strength, endurance, or a combination of both (rec league soccer, hockey, football, rugby, volleyball, or even ultimate Frisbee), the benefit of adding speed, agility, and quickness drills to your fitness routine can take your game to the next level. All athletes can benefit from improved balance, quicker feet, and a faster reaction time, and this is exactly what Speed, Agility, and Quickness (SAQ) drills help you achieve. While SAQ drills are often thought of as interchangeable, it is important to recognize how these components are related, as well as how they differ."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                    Case 13:
                        .Rows(x).RowHeight = 54.75
                        With .Cells(x, 1)
                            .Value = "Youth Judo - Judo, and Sport and method similar to wrestling developed in Japan in the 19th Century and meaning ""the gentle way"", utilizes throws and submissions."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(x, 2)
                            .Value = "Youth Fight Fit - Using high energy conditioning drills, coupled with techniques from other martial arts, MMA Fight Fit will improve your child's mental and physical endurances."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        
                    Case 15:
                        .Rows(x).RowHeight = 65.25
                        With .Cells(x, 1)
                            .Value = "Youth DUT-Youth DUT is UFC GYM's signature boot camp tailored for kids. Classes focus on dynamic full body movements that help build coordination, athletic ability and confidence in a safe environment. Classes vary ensuring each one is fun, competitive, engaging and challenging to the end."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                        With .Cells(x, 2)
                            .Value = "Youth Wrestling - This class teaches the basic grappling movements of the first olympic sport. It includes takedowns and using the Upper/Lower body with exercises and drills designed to create the ultimate wrestler."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                    Case 16:
                        .Rows(x).RowHeight = 39.75
                        .Range("A" & x & ":B" & x).Merge
                        .Range("A" & x & ":B" & x).BorderAround xlContinuous, xlMedium
                        With .Cells(x, 1)
                            .Value = "Lil' Grapplers (3 - 5) - The program integrates self-defense techniques, conditioning, and skill-oriented games aimed at improving your child's motore skills, agility and flexability."
                            .WrapText = True
                            .Font.Size = 9
                            .BorderAround xlContinuous, xlMedium
                        End With
                End Select
            End With
        Next
        
        .Columns(1).ColumnWidth = 66
        .Columns(2).ColumnWidth = 66
        
    End With
    
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
            .Borders(xlEdgeLeft).Weight = xlMedium
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
        
    Dim title, _
        coach               As Boolean
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
    pBookPath = g_objGWClient.SchedulePath & "Excel Files\" & BookTitle & " Youth Schedule.xlsx"
    
    Application.DisplayAlerts = False
    
    Set NewBook = Workbooks.Add
        
        With NewBook
            .title = BookTitle
            .Subject = "Class Schedule"
            .SaveAs FileName:=pBookPath
        End With
        
        
        
    ' * ' Add Sheets
    MainBook.Sheets("AdultScheduleTemp").Copy Before:=Workbooks(NewBook.Name).Sheets(1)
    
    
    ' * ' Rename
    NewBook.Sheets(1).Name = "Youth"
    NewBook.Sheets(2).Delete
    
    Application.DisplayAlerts = True
    
    Set pBook = NewBook
    
    ReDim pTitleRanges(0)
    ReDim pCoachRanges(0)
    ReDim pMergedRanges(0)
    ReDim pTimeRanges(0)
    ReDim pTimeRowRanges(0)
    
    pWeekDayClasses = False
    
End Sub
