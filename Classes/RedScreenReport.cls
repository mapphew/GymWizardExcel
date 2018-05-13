VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedScreenReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pRecord     As DAO.Recordset
Private pVar_Arr    As Variant
Private pReportPath As String
Private pDate As Date
Private m_bCanUpdateExp As Boolean
Private m_bExpUpdateBypass As Boolean
Private p_objClient As ClubReadyClient
Private p_ClearedMembers

Public Sub Class_InitializeWithValues(ByVal dtReportDate As Date, ByVal oRecord As DAO.Recordset, objClient As ClubReadyClient)

    ' * ' Assign Values
    Set p_objClient = objClient
    Set pRecord = oRecord
    pDate = dtReportDate
    
    ' * ' Set Can UpDate Exp For Updating EXP
    m_bCanUpdateExp = True
    m_bExpUpdateBypass = False
    
    ' ' Write
    pReportPath = g_objGWClient.RedScreenPath & Month(dtReportDate) & "." & day(dtReportDate) & "RedScreens.xlsx"
    
    Sheets("RedScreensTemp").Copy
    Sheets(1).Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs pReportPath
    ActiveWorkbook.Close

    Application.DisplayAlerts = True
    
    ' * ' Analzye Entries
    Call Analyze
    
    ' * ' Add Entries To Excel Sheet
    Call WorkbookWrite
End Sub

Private Sub WorkbookWrite()

    ' * ' Dim
    Dim book As Workbook
    Dim sheet As Worksheet
    
    ' * ' Open Book
    Set book = Workbooks.Open(pReportPath)
    
    ' * ' Set Sheet
    Set sheet = book.Sheets(1)
    
    ' * ' Title
    sTitle = MonthName(Month(pDate)) & " " & day(pDate) & GetSuffix(day(pDate)) & " Redscreens"
    sheet.Range("G1").Value = sTitle
    
    ' * ' Sheet Name
    sheet.Name = "RedScreens"
    
    ' * ' Loop Array
    For x = 1 To UBound(pVar_Arr)
    
        Row = x + 2
        ' * ' Add Data
        For y = 1 To 8
            sheet.Cells(Row, y).Value = pVar_Arr(x, y)
        Next
        
        With sheet.Range("A" & Row & ":H" & Row)
            
            ' * ' Add Borders / Font Align and Color
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            
            ' * ' Color Cells Based On The Status
            Select Case pVar_Arr(x, 8)
                
                Case "Cleared":
                    .Interior.Color = RGB(0, 255, 0)
                
                Case "Frozen":
                    .Interior.Color = RGB(0, 0, 255)
                
                Case "Cross-Club":
                    .Interior.Color = RGB(178, 178, 178)
                
                Case "Past Due":
                    .Interior.Color = RGB(255, 255, 0)
                    
                Case "Review":
                    .Interior.Color = RGB(255, 0, 0)
                    
                    
            End Select
        End With
        
    Next
    
    ' * ' Save Report And Close
    book.Save
    book.Close
        
End Sub
Private Sub Analyze()
    Dim oDatabase       As DAO.Database
    Dim oRecentExport   As DAO.Recordset
    Dim oFrozenExport   As DAO.Recordset
    Dim pClient         As New ClubReadyClient
    Dim dtNewExp        As Date
    Dim strLocation     As String
    
    
    Set oDatabase = OpenDatabase(g_objGWClient.DatabasePath)
    
    pRecord.MoveLast
    
    ' * ' Redim PVar
    ReDim pVar_Arr(1 To pRecord.RecordCount, 1 To 8)
    
    ' * ' Loop
    With pRecord
        .MoveFirst
        
        ' * ' Get The Location For this Report
        strLocation = ![Club Name]
        
        sStatus = ""
        
        Do Until .EOF
            If ![Home Club] <> strLocation Then
                sStatus = "Cross-Club"
            Else:
                ' * ' Get Member From Recent Record
                Set oRecentExport = oDatabase.OpenRecordset("SELECT * FROM [Member Export] WHERE [Member ID] ='" & ![UserID] & "';")
                Set oFrozenExport = oDatabase.OpenRecordset("SELECT * FROM [Frozen Members] WHERE [Member ID] ='" & ![UserID] & "';")
                      
                ' * ' Some Locations have a Franchise location nearby and members home club location might match
                ' * ' i.e Knox Members Would also come up as "Torrance" even though it wasnt the signature torrance
                ' * ' Checking Member Export for member will verify  if they are from the current location
                If oRecentExport.RecordCount = 0 Then
                    sStatus = "Cross-Club"
                    
                ' * ' Check Frozen Members
                ElseIf oFrozenExport.RecordCount > 0 Then
                    sStatus = "Frozen"
                    
                ' * ' Past Due Member
                ElseIf oRecentExport![Current Past Due] > 0 Then
                    sStatus = "Past Due"
                
                ' * ' If a member makes a payment or has their Expiration updated on that day, the report classifies it as denied still
                ' * ' Checking member export for current expiration verifies if they are now cleared
                ElseIf oRecentExport![Membership Expires] > pDate Then
                    sStatus = "Cleared"
                    
                ' * ' TAP Date might have changed or member was taken out of collections and didnt get the expiration updated
                ' * ' verifying that a member has a upcoming payment means they are active and need an updated expiration
                ElseIf Not IsNull(oRecentExport![Next Draft Date]) Then '' Member Expired but has a payment up coming
                    
                    ' * ' Get New Exp
                    dtNewExp = oRecentExport![Next Draft Date]
                    sStatus = "Update Expiration"
                    
                    
                ' * ' Members who dont have upcoming payments can be in collections, PIF, or Cancelled
                ElseIf IsNull(oRecentExport![Next Draft Date]) Then
                    
                    ' * ' Check Last Payment Date
                    ' * ' Members who cancel get 60 days from the last payment
                    If DateAdd("m", 2, oRecentExport![Last Payment Date]) > Now Then
                    
                        dtNewExp = DateAdd("m", 2, oRecentExport![Last Payment Date])
                        sStatus = "Cancelled / Within 60 Days"
                        
                        
                    Else:
                        ' * ' Review this Member
                        sStatus = "Review"
                    End If
                End If
            End If
            
            If (InStr(sStatus, "Expiration Needs Updating") Or InStr(sStatus, "Cancelled / Within 60 Days")) And m_bCanUpdateExp Then
                ' * ' Attempt Update
                m_bCanUpdateExp = pClient.UpdateMemberExpiration(![UserID], oRecentExport![Member Since], dtNewExp, oRecentExport![Membership Expires])
                
                If m_bCanUpdateExp = False Then
                        Select Case MsgBox("The LogIn Credentials provided do not have permissions to update member expirations. Would you like to provide proper log in credentials?", vbYesNo Or vbExclamation, Application.Name)
                    
                            Case vbYes
                                Dim frmLogIn As New frmLogIn
                                frmLogIn.Show
                                
                                Set pClient = frmLogIn.ClubReadyClient
                                Call pClient.UpdateMemberExpiration(![UserID], oRecentExport![Member Since], dtNewExp, oRecentExport![Membership Expires])
                                
                                m_bCanUpdateExp = True
                            Case vbNo
                                m_bCanUpdateExp = False
                        End Select
                End If
                
                ' * ' Change Status
                If m_bCanUpdateExp Then
                    sStatus = "Cleared"
                End If
              
            End If
            
            ' * ' Check if Review is Required
            If sStatus = "Review" Then
                Dim frmReviewCheck As New FrmCheck
                Call frmReviewCheck.PopulateMemberInfo(Trim(![First Name]) & " " & Trim(![Last Name]), ![UserID], sStatus, ![Check In Date / Time])
                sStatus = frmReviewCheck.NewStatus
            End If
            
            
            
            pVar_Arr(.AbsolutePosition + 1, 1) = ![Club Name]
            pVar_Arr(.AbsolutePosition + 1, 2) = ![UserID]
            pVar_Arr(.AbsolutePosition + 1, 3) = Trim(![Last Name])
            pVar_Arr(.AbsolutePosition + 1, 4) = Trim(![First Name])
            pVar_Arr(.AbsolutePosition + 1, 5) = ![Check In Date / Time]
            pVar_Arr(.AbsolutePosition + 1, 6) = ![Home Club]
            pVar_Arr(.AbsolutePosition + 1, 7) = ![Membership Type]
            pVar_Arr(.AbsolutePosition + 1, 8) = sStatus
            
            sStatus = ""
            Set objMember = Nothing
            
            ' * ' Next Record
            .MoveNext
            
        Loop
        
    End With
         
    ' * ' All Entries Are Now Analyzed
                    
                
                    
End Sub
