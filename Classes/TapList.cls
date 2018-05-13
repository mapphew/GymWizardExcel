VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TapList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private pDtm_FromDate As Date
Private pDtm_ToDate As Date
Private pReportPath As String
Private pObj_Database As DAO.Database
Private pObj_Recordset1 As DAO.Recordset
Private pObj_Recordset2 As DAO.Recordset
Private pDtm_NextTapDate As Date

Public Property Get ReportPath() As String
    ReportPath = pReportPath
End Property
Private Sub TapWriter(ByVal aVar_Arr As Variant, ByVal aVar_Arr2 As Variant, ByVal aStr_Title As String, ByVal aDtm_Date As Date)

    ' ' Declarations
    Dim mObj_WrkBk As Workbook
    Dim mObj_WrkSht As Worksheet
    
    Dim mInt_Count, mInt_Row As Integer
    Dim mStr_Type As String
    
    ' ' Open Workbook
    Set mObj_WrkBk = Workbooks.Open(pReportPath)
    Set mObj_WrkSht = mObj_WrkBk.Sheets(1)
    
    With mObj_WrkBk.Sheets(1)
        
        .Range("C3").Value = aStr_Title
        
        ' * ' Pre Tap
        For mInt_Count = 1 To UBound(aVar_Arr2)
            mInt_Row = 30 + mInt_Count
            
            .Cells(mInt_Row, 2).Value = aVar_Arr2(mInt_Count, 1)
            .Cells(mInt_Row, 3).Value = aVar_Arr2(mInt_Count, 2)
            .Cells(mInt_Row, 4).Value = aVar_Arr2(mInt_Count, 3)
            .Cells(mInt_Row, 5).Value = "No CARD"
            
            With Range(.Cells(mInt_Row, 2), .Cells(mInt_Row, 8))
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
            
        Next
            
        For mInt_Count = UBound(aVar_Arr) To LBound(aVar_Arr) Step -1
                        
            Select Case aVar_Arr(mInt_Count, 5)
                
                Case 1
                    mInt_Row = 9
                
                Case 2
                    mInt_Row = 13
                    
                Case 3
                    mInt_Row = 17
                    
                Case 4, 5
                    mInt_Row = 21
                    
                Case 6
                    mInt_Row = 25
            End Select
            
            If mInt_Count <> UBound(aVar_Arr) Then
                If aVar_Arr(mInt_Count + 1, 5) = aVar_Arr(mInt_Count, 5) Then
                    .Rows(mInt_Row).EntireRow.Insert
                End If
            End If
            
            .Cells(mInt_Row, 2).Value = aVar_Arr(mInt_Count, 1)
            .Cells(mInt_Row, 3).Value = aVar_Arr(mInt_Count, 2)
            .Cells(mInt_Row, 4).Value = aVar_Arr(mInt_Count, 3)
            
            With .Cells(mInt_Row, 8)
                If Len(aVar_Arr(mInt_Count, 6)) > 25 Then
                    .Row.Height = 20
                    .WrapText = True
                End If
                
                .Value = aVar_Arr(mInt_Count, 6)
                .Font.Italic = True
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
            End With
            
            Select Case aVar_Arr(mInt_Count, 4)
            
                Case True
                    .Cells(mInt_Row, 5).Value = ChrW(&H2713)
            
            End Select
            
            With Range(.Cells(mInt_Row, 2), .Cells(mInt_Row, 8))
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
            
            With .Range("B9").EntireColumn
                .HorizontalAlignment = xlCenter
            End With
            
            With .Range("E9").EntireColumn
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
           
        Next
        
        
        
    End With
    
    ' ' Page Set UP
    With mObj_WrkSht
        .PageSetup.Order = xlDownThenOver
    End With
        
    ' * ' Pre Tap Time
    
    mObj_WrkBk.Close True
    
    
End Sub
Public Sub InitializeWithValues(ByVal FromDate As Date, ByVal ToDate As Date, ByVal NextTapDate As Date)

    ' ' Vars
    Dim pVar_Arr As Variant
    Dim pVar_Arr2 As Variant
    Dim FSO As New FileSystemObject
    
    ' ' Set Date
    pDtm_FromDate = FromDate
    pDtm_ToDate = ToDate
    pDtm_NextTapDate = NextTapDate
    
        
        
    mStr_Title = MonthName(Month(ToDate)) & " " & day(ToDate) & GetSuffix(day(ToDate)) & " Tap List"
    
    
    Call setPastDueNote
    Set pObj_Recordset1 = pObj_Database.OpenRecordset("SELECT * FROM [Invoices Coming Due] WHERE [Due Date] >= #" & pDtm_FromDate & _
                                                                                    "# AND [Due Date] <= #" & pDtm_ToDate & "# " & _
                                                                                    "AND [Paid Date] is Null AND [Due Status] <> 'Cancelled' " & _
                                                                                    " ORDER BY [Note], [Has ACH],[Has CC],[Last Name];")
    
    ' ' Loop
    pObj_Recordset1.MoveFirst
    
    
    With pObj_Recordset1
        
            ReDim pVar_Arr(1 To .RecordCount, 1 To 6)
            
            .MoveFirst
            
            Do Until .EOF
                If ![Note] = 4 Then
                    pVar_Arr(.AbsolutePosition + 1, 6) = "Client of " & ![Sold By]
                ElseIf ![Note] = 5 Then
                    pVar_Arr(.AbsolutePosition + 1, 6) = "Membership Sold By " & ![Sold By]
                Else:
                    pVar_Arr(.AbsolutePosition + 1, 6) = ""
                End If
                
                    
                pVar_Arr(.AbsolutePosition + 1, 1) = ![User ID]
                pVar_Arr(.AbsolutePosition + 1, 2) = Trim(![Last Name])
                pVar_Arr(.AbsolutePosition + 1, 3) = Trim(![First Name])
                pVar_Arr(.AbsolutePosition + 1, 4) = ![Has CC] Or ![Has ACH]
                pVar_Arr(.AbsolutePosition + 1, 5) = ![Note]
                .MoveNext
            Loop
        
        
        
    End With
    
    ' * ' Get PRE TAP
    Set pObj_Recordset1 = pObj_Database.OpenRecordset("SELECT * FROM [Invoices Coming Due] WHERE [Due Date] = #" & pDtm_NextTapDate & _
                                                                                    "# AND [Note] = 0 AND ([Has CC] = False AND [Has ACH] = False) AND [Paid Date] is Null AND [Due Status] <> 'Cancelled' " & _
                                                                                    " ORDER BY [Last Name];")
    
    With pObj_Recordset1
        
            ReDim pVar_Arr2(1 To .RecordCount, 1 To 6)
            
            .MoveFirst
            
            Do Until .EOF
                
                    
                pVar_Arr2(.AbsolutePosition + 1, 1) = ![User ID]
                pVar_Arr2(.AbsolutePosition + 1, 2) = Trim(![Last Name])
                pVar_Arr2(.AbsolutePosition + 1, 3) = Trim(![First Name])
                pVar_Arr2(.AbsolutePosition + 1, 4) = ![Has CC] Or ![Has ACH]
                pVar_Arr2(.AbsolutePosition + 1, 5) = ![Note]
                pVar_Arr2(.AbsolutePosition + 1, 6) = ""
                .MoveNext
            Loop
        
        
        
    End With
    
    ' ' Write
    pReportPath = g_objGWClient.TAPListPath & Month(pDtm_ToDate) & "." & day(pDtm_ToDate) & "TapList.xlsx"
    
    Sheets("TapListTemp").Copy
    Sheets(1).Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs pReportPath

    Application.DisplayAlerts = True
    
    Call TapWriter(pVar_Arr, pVar_Arr2, mStr_Title, pDtm_ToDate)
    
    
 
    
    ' * ' Format
    Dim book As Workbook
    Dim sheet As Worksheet
    
    Set book = Workbooks.Open(pReportPath)
    Set sheet = book.Sheets(1)
    
    
    With sheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .PrintTitleRows = "$5:$5"
    End With
    
    book.Close True
    
    
End Sub
Private Sub setPastDueNote()
    
    ' ' Declarations
    Dim mObj_Tap, mObj_PastDue As DAO.Recordset
    Dim mStr_Note As String
    
     ' ' Open Record Set for Invoices Due for TAP DATE
    Set pObj_Recordset1 = pObj_Database.OpenRecordset("SELECT * FROM [Invoices Coming Due] WHERE [Due Date] >= #" & pDtm_FromDate & "#" & _
                                                                                            " AND [Due Date] <= #" & pDtm_NextTapDate & "# AND [Paid Date] is Null;")
            
            
            With pObj_Recordset1
            
                ' ' Loop Through Records
                .MoveFirst
                
                Do Until .EOF
                    
                    ' * ' Check If Due on TAP
                    mStr_Note = 0
                        
                    If ![GSL Category] = "Service" Then
                        mStr_Note = 4
                    Else:
                        If ![Invoice Category] = "Post Dated Down" Then
                            mStr_Note = 5
                        ElseIf ![Due Status] = "Third Party Cancelled" Then
                            mStr_Note = 6
                        Else:
                            Set pObj_Recordset2 = pObj_Database.OpenRecordset("SELECT * FROM [Club Past Due] WHERE [User ID] = '" & ![User ID] & _
                                                                                                    "' AND Day([Due Date]) =" & day(![Due Date]) & _
                                                                                                    " ORDER BY [Due Date] ASC;")
                            
                            With pObj_Recordset2
                                
                                If .EOF Then
                                    mStr_Note = 0
                                
                                Else:
                                
                                    .MoveFirst
                                    
                                    Select Case DateDiff("m", pObj_Recordset2![Due Date], pObj_Recordset1![Due Date])
                                        
                                        Case 0
                                            mStr_Note = 1
                                            
                                        Case 1
                                            mStr_Note = 2
                                        
                                        Case 2
                                            mStr_Note = 3
                                        
                                    End Select
                                    
                                End If
                                
                            End With
                            
                        End If
                    End If
                    
                .Edit
                ![Note] = mStr_Note
                .Update
                .MoveNext
                
            Loop
            pObj_Recordset2.Close
            pObj_Recordset1.Close
            
        End With
End Sub
Private Sub Class_Initialize()
    Set pObj_Database = OpenDatabase(g_objGWClient.DatabasePath)
End Sub
