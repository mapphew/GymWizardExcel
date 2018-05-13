VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClubReadyClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sStoreID          As String
Private m_objXMLClient      As XMLHTTP60
Private m_sSessionCookie    As String
Private m_bCanAccessReports As Boolean
Private m_bCanChangeMemExp  As Boolean

Private Sub Class_Initialize()
    ' * ' Check For Globals
    If g_objGWClient Is Nothing Then
        Set g_objGWClient = New GymWizardClient
    End If
    
    Set m_objXMLClient = New XMLHTTP60
    
End Sub

Private Property Get SessionCookie() As String

    SessionCookie = m_sSessionCookie

End Property


Public Property Get StoreID() As String

    StoreID = m_sStoreID

End Property

Private Property Get XMLClient() As XMLHTTP60

    Set XMLClient = m_objXMLClient

End Property
' ----------------------------------------------------------------
' Procedure Name: OnTimePaymentHistory
' Purpose: Retrieves the payment history of an agreement
' Procedure Kind: Function
' Procedure Access: Public
' Parameter s_AgreementNumber (String):
' Return Type: Variant()
' Author: UFCGym
' Date: 2/16/2018
' ----------------------------------------------------------------
Public Function OnTimePaymentHistory(ByVal s_AgreementNumber As String, ByVal s_MemberId As String) As Integer
    
    Dim s_MyUrl As String
    Dim o_HTMLDoc As New HTMLDocument
    Dim o_IMGCol As IHTMLElementCollection
    Dim o_IMGItr As HTMLImg
    Dim o_TR    As HTMLTableRow
    Dim o_Div   As HTMLDivElement
    Dim o_ScriptItr As HTMLScriptElement
    Dim o_ScriptCol    As IHTMLElementCollection
    Dim s_JScriptVar As String
    
    
    
    s_MyUrl = "http://clubready.com/admin/manageagreement.asp?id=" & s_MemberId & "&cpid=" & s_AgreementNumber & "&t=1"
    
    With XMLClient
        .Open "GET", s_MyUrl
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Set-Cookie", SessionCookie
        .send
        
    End With
    
    o_HTMLDoc.body.innerHTML = XMLClient.responseText
    
    Set o_ScriptCol = o_HTMLDoc.getElementsByTagName("script")
    
    For Each o_ScriptItr In o_ScriptCol
        
        ' * ' get Right one
        If InStr(o_ScriptItr.innerHTML, "ethecont") Then
            a_split = Split(o_ScriptItr.innerHTML, ";")
            
            Dim l_Itr As Long
            For l_Itr = 0 To UBound(a_split)
                If InStr(a_split(l_Itr), "ethecont") Then
                    a_split = Split(a_split(l_Itr), "=")
                    Exit For
                End If
            Next
            
            s_JScriptVar = Trim(Replace(a_split(1), "'", ""))
            
            Exit For
        End If
        
        
    Next
    
    ' * ' Now Get The Agreement
    Payload = "cpid=" & s_JScriptVar
    
    With XMLClient
        .Open "POST", "http://clubready.com/common/agreement-management/ajax_agreementdetail.asp", False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Set-Cookie", SessionCookie
        .send Payload
    End With
    
    o_HTMLDoc.body.innerHTML = XMLClient.responseText
    
    Set o_IMGCol = o_HTMLDoc.getElementsByTagName("img")
    
    For Each o_IMGItr In o_IMGCol
        Debug.Print (o_IMGItr.getAttribute("src"))
        
        If InStr(o_IMGItr.getAttribute("src"), "check16.png") Then
            Set o_Div = o_IMGItr.parentElement.parentElement
            
            Exit For
        End If
    Next
    
    OnTimePaymentHistory = CInt(Replace(o_Div.Children(3).innerText, "%", ""))
    
End Function
Private Function ParseJScriptResponse(ByVal sValue As String) As Variant
    
    ' * ' DIm
    Dim iStart, _
        iEnd            As Integer
    Dim Itr             As Integer
    Dim Vals            As Variant
    Dim newString As String
    
    Dim myChar As Char
    
    iStart = -1
    iEnd = -1
    
    Vals = Split(sValue, ";")
    
    newString = Right(Vals(0), Len(Vals(0)) - InStr(Vals(0), "cache") - 5)

    Vals = Split(newString, "&")
    
    ParseJScriptResponse = Vals
End Function
Public Function GetMemberAgreements(ByVal m_sMemberID As String) As Variant
    Dim m_sMyUrl As String
    Dim m_oDoc As New HTMLDocument
    Dim m_oAgreementRow As HTMLTableRow
    Dim m_oAgreementHTMLTable As HTMLTable
    Dim m_iAgrCount As Integer
    ''Dim m_oSpan As HTMLSpanElement
    ' ' Set Url
    m_sMyUrl = "http://clubready.com/admin/contractpurchases.asp?id=" & m_sMemberID
    
    ' * ' GET HTML
    XMLClient.Open "GET", m_sMyUrl
    XMLClient.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    XMLClient.setRequestHeader "Set-Cookie", SessionCookie
    XMLClient.send
    
    ' * ' Get Document
    m_oDoc.body.innerHTML = XMLClient.responseText
    
    Set m_oAgreementRow = m_oDoc.getElementById("pgtr")
    
    ' * ' GET TABLE WITH AGREEMENTS
    Set m_oAgreementHTMLTable = m_oAgreementRow.getElementsByTagName("td")(1).Children(6)
    
    ' * ' Get Agreement Count
    m_iAgrCount = m_oAgreementHTMLTable.Children(0).Children.Length - 1
    
    ' * ' Create Variant To Hold Rows
    Dim m_vAgrRows() As Variant
    Dim m_vAgrMaster() As Variant
    ReDim m_vAgrMaster(1 To m_iAgrCount)
    ReDim m_vAgrRows(1 To m_iAgrCount)
    
    ' * ' Populate Variant
    For x = 1 To m_iAgrCount
        Set m_vAgrRows(x) = m_oAgreementHTMLTable.Children(0).Children(x)
    Next
    
    ' * ' Process HTML Data TO Get Agreement Data
    For x = 1 To m_iAgrCount
        Dim m_vAgr(1 To 6) As Variant
        m_sAgrSP1 = "None"
        m_sAgrSP2 = "None"
        
        ' * ' Get Data
        m_sAgrDate = DateValue(m_vAgrRows(x).Children(0).innerText)
        
        m_sAgrStatus = m_vAgrRows(x).Children(1).innerText
        If InStr(m_sAgrStatus, "Cancelled") > 0 Then
            m_sAgrStatus = "Cancelled"
        ElseIf InStr(m_sAgrStatus, "Fully Paid") > 0 Then
            m_sAgrStatus = "Fully Paid"
        ElseIf InStr(m_sAgrStatus, "Active EFT") > 0 Then
            If InStr(m_sAgrStatus, "Frozen") > 0 Then
                m_sAgrStatus = "Frozen"
            Else:
                m_sAgrStatus = "Active"
            End If
        ElseIf InStr(m_sAgrStatus, "Third party since") > 0 Then
            m_sAgrStatus = "Third Party"
        ElseIf InStr(m_sAgrStatus, "Third party settled") > 0 Then
            m_sAgrStatus = "Third Party Settled"
        Else:
            Debug.Print ("Error! " & m_sAgrStatus & " Is Not a Recognized Status!")
            
        End If
        
        If InStr(m_vAgrRows(x).Children(2).innerText, "&") > 0 Then '' Agr Is SPlit
            m_sAgrSP1 = Trim(Split(m_vAgrRows(x).Children(2).innerText, "&")(0))
            m_sAgrSP2 = Trim(Split(m_vAgrRows(x).Children(2).innerText, "&")(1))
        Else:
            m_sAgrSP1 = Trim(m_vAgrRows(x).Children(2).innerText)
        End If
        
        
        m_oSpan = Trim(Split(m_vAgrRows(x).Children(3).Children(0).Children(1).innerHTML, "<A")(0))
        
        m_sAgrNum = Replace(Replace(Replace(Split(m_oSpan, ")")(0), ")", ""), "(", ""), "#", "")
        m_sAgrTitle = Split(Trim(Split(m_oSpan, ")")(1)), " x ")(1)
        
        m_vAgr(1) = m_sAgrDate
        m_vAgr(2) = m_sAgrStatus
        m_vAgr(3) = m_sAgrSP1
        m_vAgr(4) = m_sAgrSP2
        m_vAgr(5) = m_sAgrNum
        m_vAgr(6) = m_sAgrTitle
        
        ' * ' Add To Master
        m_vAgrMaster(x) = m_vAgr
    Next
    
    GetMemberAgreements = m_vAgrMaster
End Function



Public Function DownloadReports(Optional ByVal rGSD As Report_Gross_Sales_Detail = Nothing, _
                                Optional ByVal rICD As Report_Invoices_Coming_Due = Nothing, _
                                Optional ByVal rCPD As Report_Club_Past_Due = Nothing) As Boolean
    
    Dim vVar1(0 To 2)   As Variant
    Dim vVar2()         As Variant
    Dim rHandler        As New Report_DownloadHandler
    
    ' * ' First Var
    Set vVar1(0) = rGSD
    Set vVar1(1) = rICD
    Set vVar1(2) = rCPD
    
    ' * ' DownloadVar
    ReDim vVar2(0)
    
    ' * ' Loop
    For x = 0 To UBound(vVar1)
        ' * ' Add Reports That Aren't Nothing
        If Not vVar1(x) Is Nothing Then
            Set vVar2(UBound(vVar2)) = vVar1(x)
            ReDim Preserve vVar2(UBound(vVar2) + 1)
        End If
    Next
    
    ' * ' Cut Off End Index
    ReDim Preserve vVar2(UBound(vVar2) - 1)
    
    Call rHandler.BulkReportDownload(vVar2)
    
    
    
End Function
' ----------------------------------------------------------------
' Procedure Name: DownloadReport
' Purpose: Downloads given report and uploads it to Access DB
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sReportType (String): Type of Report to be download
' Parameter dFromDate (Date):
' Parameter dToDate (Date):
' Parameter sMemberID (String):
' Return Type: Boolean
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Public Function DownloadReport(Optional ByVal rRep As Variant = Nothing, Optional ByVal sUrl As String, Optional ByVal sFileName As String, Optional ByVal ReportType As String, Optional ByVal ClearTable As Boolean = True) As Boolean
    
    On Error GoTo DownloadReport_Error
    Dim sReportType     As String
    Dim sReportPath     As String
    Dim sReportUrl      As String
    Dim sDatabasePath   As String
    Dim objStream       As New ADODB.Stream
    Dim objDBConnection As New ADODB.Connection
    Dim objWorkbook     As Workbook
    Dim objWorksheet    As Worksheet
    Dim iLastRow        As Integer
    
    '' ------------------------------- ''
    '' ----------DEPRICATED----------- ''
    '' ------------------------------- ''
    ' * ' Check If Has report permissions
    ''If m_bCanAccessReports = False Then
        ''DownloadReport = False
        ''Exit Function
    ''End If
    
    
    If Not rRep Is Nothing Then
        sReportType = Replace(Replace(TypeName(rRep), "Report_", ""), "_", " ")
        sReportPath = g_objGWClient.TempFolderPath & rRep.FileName
        sReportUrl = rRep.DownloadUrl
    Else:
        sReportType = ReportType
        sReportPath = g_objGWClient.TempFolderPath & sFileName
        sReportUrl = sUrl
    End If
    
    
    ' ' Download Report
    XMLClient.Open "GET", sReportUrl
    XMLClient.send
    
    ' * ' Get Stream
    objStream.Open
    objStream.Type = 1
    objStream.Write XMLClient.responseBody
    objStream.SaveToFile sReportPath, adSaveCreateOverWrite
    objStream.Close
    
    ' * ' Disable Alerts To Avoid Error Message for Outdated Excel extensions
    Application.DisplayAlerts = False
    
    ' * ' Open Report Workbook
    Set objWorkbook = Workbooks.Open(sReportPath)
    Set objWorksheet = objWorkbook.Sheets(1)
    
    ' * ' Set Worksheet name to Sheet1 for easier DB uploading
    objWorksheet.Name = "Sheet1"
    
    ' * ' Now we do some worksheet clean up based on Report Type
    Select Case sReportType
        Case "Attendance":
            objWorksheet.Range("A1").EntireRow.Delete
        
        Case "Member Export", "Frozen Members":
            With objWorksheet
                .Range("AH6").EntireColumn.Delete
                .Range("C6").EntireColumn.Delete
                .Range("A6").EntireColumn.Delete
                .Range("A1:A5").EntireRow.Delete
            End With
            
                
        Case "Agreements Log":
            ' * ' Get Last Row
            iLastRow = objWorksheet.Cells(500, "C").End(xlUp).Row
            
            objWorksheet.Range("A" & iLastRow).Rows.EntireRow.Delete
            objWorksheet.Range("A1").Rows.EntireRow.Delete
            
            objWorksheet.Range("AR1").Columns.EntireColumn.Delete
            objWorksheet.Range("AN1").Columns.EntireColumn.Delete
            
        Case "Gross Sales Detail":
            objWorksheet.Range("A1").Rows.EntireRow.Delete
            objWorksheet.Range("AH1").Columns.EntireColumn.Delete
        
        Case "Refunds":
        
            ' * ' Get Last Row
            Dim rowItr As Long
            
            For rowItr = 2 To 1000
                
                If objWorksheet.Range("C" & rowItr).Value = "" Then
                    iLastRow = rowItr
                    rowItr = 1001
                End If
            
            Next
            
            objWorksheet.Range("A" & iLastRow).Rows.EntireRow.Delete
            objWorksheet.Range("A1").Rows.EntireRow.Delete
            
            
        Case "Club Past Due":
            objWorksheet.Range("A1").Rows.EntireRow.Delete
            
            objWorksheet.Range("B1").Columns.EntireColumn.NumberFormat = "@"
            objWorksheet.Range("N1:Q1").Columns.EntireColumn.NumberFormat = "@"
            objWorksheet.Range("K1:L1").Columns.EntireColumn.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
            
            objWorksheet.Columns("L").Select
            
            Selection.Replace What:="", Replacement:="0", LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                        ReplaceFormat:=False
        Case "Invoices Coming Due":
            objWorksheet.Range("A1").Rows.EntireRow.Delete
            objWorksheet.Range("B1:C1").Columns.EntireColumn.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
            objWorksheet.Range("A1").Columns.EntireColumn.NumberFormat = "@"
            objWorksheet.Range("G1:K1").Columns.EntireColumn.NumberFormat = "@"
            objWorksheet.Range("L1:M1").Columns.EntireColumn.NumberFormat = "@"
            objWorksheet.Columns("L:M").Select
            
            Selection.Replace What:="Y", Replacement:="YES", LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                        ReplaceFormat:=False
            Selection.Replace What:="", Replacement:="NO", LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                        ReplaceFormat:=False
            
    End Select
    
    ' * ' Close Workbook
    objWorkbook.Close True
    
    ' * ' Get Database Path
    sDatabasePath = g_objGWClient.DatabasePath
        
    ' * ' Turn alerts Back on
    Application.DisplayAlerts = True
    
    ' * ' Upload workbook data to Database
    objDBConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sReportPath & ";Extended Properties=Excel 8.0;" _
        & "Persist Security Info=False"
    
    
    If ClearTable Then
        objDBConnection.Execute "DELETE * FROM [" & sReportType & "] IN '" & sDatabasePath & "'"
    End If
    
    objDBConnection.Execute "INSERT INTO [" & sReportType & "] IN '" & sDatabasePath & "' SELECT * FROM [Sheet1$]"
        
    objDBConnection.Close
    
    ' * ' Delete Report
    ''DeleteFile sReportPath
    
    ' * ' After Upload Shinanigans
    Select Case sReportType
        
        Case "Refunds":
            Dim objGWDB As New GWDatabase
            objGWDB.ProcessRefundData
                
    End Select
    
    
    
    
    DownloadReport = True
    
    
    On Error GoTo 0
    Exit Function

DownloadReport_Error:

    DownloadReport = False

End Function
Public Function CreatePrintableSchedule(ByVal YouthBool As Boolean, ByVal AdultBool As Boolean)
    
    Dim MasterScheduleArray(2) As Variant
    Dim o_Database As DAO.Database
    Dim o_Recordset As DAO.Recordset
    Dim t_TimeItr As Date
    
    Dim Monday(), _
        Tuesday(), _
        Wednesday(), _
        Thursday(), _
        Friday(), _
        Saturday(), _
        Sunday()            As Variant
        
    Dim InsertArray         As Variant
    Dim Youth, _
        Ultimate, _
        Fitness             As Variant
    ReDim InsertArray(0)
    
    
    ' * ' TODO Create DB Clear Here
    Set o_Database = OpenDatabase(g_objGWClient.DatabasePath)
    
    o_Database.Execute "DELETE FROM [Scheduled Classes];"
    
    ' * ' Get and Parse Schedules
    Call ProcessScheduleHTML(GetScheduleHTML("cat|41"), "Fitness")
    Call ProcessScheduleHTML(GetScheduleHTML("cat|43"), "Ultimate")
    Call ProcessScheduleHTML(GetScheduleHTML("26706"), "Ultimate", "No-GI")
    Call ProcessScheduleHTML(GetScheduleHTML("cat|44"), "Youth")
    Call ProcessScheduleHTML(GetScheduleHTML("cat|42"), "Ultimate")
    
    ' * ------------------------------ START ARRAY * -------------------------------------'
    ' * ' Format In DB
    Dim o_DB As New GWDatabase
    o_DB.FormatClasses
    
    
    Dim l_Itr As Long
    For l_Itr = 0 To 2
    
        ' * ' Earliest Class is at 5:00AM and 9:00 PM is Latest
        t_TimeItr = TimeValue("5:00AM")
        RowTotal = 0
        
        Do While t_TimeItr <= TimeValue("10:00PM")
            
            Select Case l_Itr
                Case 0: s_ClassType = "Youth"
                Case 1: s_ClassType = "Ultimate"
                Case 2: s_ClassType = "Fitness"
            End Select
            
            ' * ' Get Classes That Match Time and Type
            Set o_Recordset = o_Database.OpenRecordset("SELECT * FROM [Scheduled Classes] WHERE [Type] = '" & s_ClassType & "' AND [Start Time] =#" & t_TimeItr & "#;")
            
            ' * ' Create Arrays For Each Day
            Dim TimeArray()         As Variant
                
            ' * ' ReDim
            ReDim Monday(0)
            ReDim Tuesday(0)
            ReDim Wednesday(0)
            ReDim Thursday(0)
            ReDim Friday(0)
            ReDim Saturday(0)
            ReDim Sunday(0)
            
            
            ' * ' Loop Through All Classes And Match Time
            With o_Recordset
                ' * ' Check If Empty
                If Not .BOF And Not .EOF Then
                    .MoveFirst
                    
                    Do Until .EOF
                        
                        Dim o_NewClass As New Schedule_Class
                        o_NewClass.CreateClass ![Formatted Title], ![Start Time], ![Coach1], ![Type]
                        
                        ' * ' Get Day of the week and put in var
                        Select Case ![Weekday]
                            Case 1:
                                Monday = AddClass(Monday, o_NewClass)
                                
                            Case 2:
                                Tuesday = AddClass(Tuesday, o_NewClass)
                                
                            Case 3:
                                Wednesday = AddClass(Wednesday, o_NewClass)
                                
                            Case 4:
                                Thursday = AddClass(Thursday, o_NewClass)
                                
                            Case 5:
                                Friday = AddClass(Friday, o_NewClass)
                                
                            Case 6:
                                Saturday = AddClass(Saturday, o_NewClass)
                                
                            Case 7:
                                Sunday = AddClass(Sunday, o_NewClass)
                        End Select
                        
                    ' * ' Next Record
                    .MoveNext
                    Loop
                
                End If
                
            End With
            
            ' * ' Array used to store all classes for that time slot
            TimeArray = Array(t_TimeItr, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday)
               
            Dim TimeArrayRow As New Schedule_TimeRow
            TimeArrayRow.Class_InitializeWithValues (TimeArray)
            
            If TimeArrayRow.NotEmpty Then
                
                ' * ' Add Values To InsertArray
                ReDim Preserve InsertArray(UBound(InsertArray) + 1)
                Set InsertArray(UBound(InsertArray)) = TimeArrayRow
                    
            End If
               
            Set TimeArrayRow = Nothing
            
            ' * ' Add 30 Minutes To Time
            t_TimeItr = DateAdd("n", 30, t_TimeItr)
            
            
            
        Loop
    
        ' * ' Clean The Array Up
        Dim v_NewArray As Variant
        ReDim v_NewArray(UBound(InsertArray) - 1)
        
        Dim l_VarItr As Long
        For l_VarItr = 1 To UBound(InsertArray)
        
            Set v_NewArray(l_VarItr - 1) = InsertArray(l_VarItr)
            
        Next
        
        ' * ' Add to Master
        MasterScheduleArray(l_Itr) = v_NewArray
        Set v_NewArray = Nothing
        ReDim InsertArray(0)
    Next
        
    
    
    ' * ' Create Schedule Object For Adult and Youth
    Dim AdultSched As New Schedule_AdultSchedule
    Dim YouthSched  As New Schedule_YouthSchedule
    
    If YouthBool Then
        YouthSched.CreateSchedule MasterScheduleArray(0)
    End If
    If AdultBool Then
        AdultSched.CreateSchedule MasterScheduleArray(1), MasterScheduleArray(2)
    End If
    
    
End Function

Public Function GetExcelScheduleArray(ByVal s_ClassType As String) As Variant
    
    Set o_Database = OpenDatabase(g_objGWClient.DatabasePath)
                      
        
    ' * ' Earliest Class is at 5:00AM and 9:00 PM is Latest
    t_TimeItr = TimeValue("5:00AM")
    RowTotal = 0
        
    Do While t_TimeItr <= TimeValue("10:00PM")
            
        ' * ' Get Classes That Match Time and Type
        Set o_Recordset = o_Database.OpenRecordset("SELECT * FROM [Scheduled Classes] WHERE [Type] = '" & s_ClassType & "' AND [Start Time] =#" & t_TimeItr & "#;")
            
        ' * ' Create Arrays For Each Day
        Dim TimeArray()         As Variant
            
        ' * ' ReDim
        ReDim Monday(0)
        ReDim Tuesday(0)
        ReDim Wednesday(0)
        ReDim Thursday(0)
        ReDim Friday(0)
        ReDim Saturday(0)
        ReDim Sunday(0)
            
            
        ' * ' Loop Through All Classes And Match Time
        With o_Recordset
            ' * ' Check If Empty
            If Not .BOF And Not .EOF Then
                .MoveFirst
                
                Do Until .EOF
                    
                    Dim o_NewClass As New Schedule_Class
                    o_NewClass.CreateClass ![Formatted Title], ![Start Time], ![Coach1], ![Type]
                    
                    ' * ' Get Day of the week and put in var
                    Select Case ![Weekday]
                        Case 1:
                            Monday = AddClass(Monday, o_NewClass)
                            
                        Case 2:
                            Tuesday = AddClass(Tuesday, o_NewClass)
                            
                        Case 3:
                            Wednesday = AddClass(Wednesday, o_NewClass)
                            
                        Case 4:
                            Thursday = AddClass(Thursday, o_NewClass)
                            
                        Case 5:
                            Friday = AddClass(Friday, o_NewClass)
                            
                        Case 6:
                            Saturday = AddClass(Saturday, o_NewClass)
                            
                        Case 7:
                            Sunday = AddClass(Sunday, o_NewClass)
                    End Select
                    
                ' * ' Next Record
                .MoveNext
                Loop
            
            End If
            
        End With
        
        ' * ' Array used to store all classes for that time slot
        TimeArray = Array(t_TimeItr, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday)
           
        Dim TimeArrayRow As New Schedule_TimeRow
        TimeArrayRow.Class_InitializeWithValues (TimeArray)
        
        If TimeArrayRow.NotEmpty Then
            
            ' * ' Add Values To InsertArray
            ReDim Preserve InsertArray(UBound(InsertArray) + 1)
            Set InsertArray(UBound(InsertArray)) = TimeArrayRow
                
        End If
               
        Set TimeArrayRow = Nothing
            
        ' * ' Add 30 Minutes To Time
        t_TimeItr = DateAdd("n", 30, t_TimeItr)
            
    Loop
    
    
        
    GetScheduleExcelArray = InsertArray
  
    
End Function
Private Function AddClass(ByVal Arr As Variant, ByVal Class As Schedule_Class) As Variant

    If UBound(Arr) = 0 Then
        If IsEmpty(Arr(0)) Then
            Arr(0) = Array(Class.Name, Class.coach)
        Else:
            ReDim Preserve Arr(UBound(Arr) + 1)
            Arr(UBound(Arr)) = Array(Class.Name, Class.coach)
        End If
    Else:
        ReDim Preserve Arr(UBound(Arr) + 1)
        Arr(UBound(Arr)) = Array(Class.Name, Class.coach)
    End If
    
    AddClass = Arr
End Function
Private Function ProcessScheduleHTML(ByVal o_Doc As HTMLDocument, ByVal s_ClassType As String, Optional ByVal s_Variance As String = "Default")
   
   Dim o_SchedTable As HTMLTable
   Dim o_RowCol As IHTMLElementCollection
   Dim o_TableRow As HTMLTableRow
   Dim l_ChildNum As Long
    
    ' * ' OPEN DB For Later Use
    Dim o_Database As DAO.Database
    Dim o_Recordset As DAO.Recordset
    
    Set o_Database = OpenDatabase(g_objGWClient.DatabasePath)
    
    
    
    
   ' * ' Now Get The Main Table
   Set o_SchedTable = o_Doc.getElementsByTagName("table")(0)
   
   ' * ' Get ALL Rows from Main Table
   Set o_RowCol = o_SchedTable.getElementsByTagName("tr")

    ' * ' Get The Rows For The Classes
    Set o_ClassCol = o_SchedTable.Children(0).Children
    
    ' * ' First Child is always the Header Row
    For l_ChildNum = 1 To o_ClassCol.Length - 1
    
        ' * ' Set The Row That will be iterated through for classes
        Set o_TableRow = o_ClassCol(l_ChildNum)
        
        
        Dim l_DayNum As Long
        Dim o_DayCell As HTMLTableCell
        
        ' * ' Every Row Has 8 children the first child is the time cell
        ' * ' the children at index 1-7 correlate to the day of the week; 1 being monday
        For l_DayNum = 1 To 7
            
            ' * ' DayCell contains all classes for that time and day
            Set o_DayCell = o_TableRow.Children(l_DayNum)
            
            Dim l_ClassNum As Long
            Dim o_ClassTable As HTMLTableCell
            For l_ClassNum = 1 To o_DayCell.Children.Length
                
                ' * ' Get Class ID
                s_ClassID = Replace(o_DayCell.Children(l_ClassNum - 1).ID, "class_", "")
                
                
                ' * ' Each class is an HTMLTable
                Set o_ClassTable = o_DayCell.Children(l_ClassNum - 1).Children(0).Children(0).Children(0)
                
                ' * ' Get Title and Start Time
                If o_ClassTable.Children.Length = 3 Then
                    s_TitleTimeRaw = o_ClassTable.Children(0).innerHTML
                    ' * ' Get Coach
                    s_Coach1 = o_ClassTable.Children(2).innerText
                Else:
                    s_TitleTimeRaw = o_ClassTable.Children(1).innerHTML
                    
                    ' * ' Get Coach
                    s_Coach1 = o_ClassTable.Children(3).innerText
                End If
                
                a_split = Split(s_TitleTimeRaw, "</B>")
                
                If UBound(a_split) >= 2 Then
                    s_Title = Replace(a_split(1), "<B>", "")
                    s_Time = Replace(Split(Replace(a_split(2), "</B>", ""), "-")(0), "<BR>", "")
                Else:
                    s_Title = Replace(a_split(0), "<B>", "")
                    s_Time = Replace(Split(Replace(a_split(1), "</B>", ""), "-")(0), "<BR>", "")
                End If
                
                s_Coach2 = "None"
                
                ' * ' Insert Data into
                Set o_Recordset = o_Database.OpenRecordset("SELECT * FROM [Scheduled Classes] WHERE [Class ID] = '" & s_ClassID & "';")
                
                With o_Recordset
                    If .BOF And .EOF Then
                        .AddNew
                    Else:
                        .Edit
                    End If
                    
                    ![Class ID] = s_ClassID
                    ![Start Time] = DateTime.TimeValue(s_Time)
                    ![Weekday] = l_DayNum
                    ![Raw Title] = s_Title
                    ![Coach1] = s_Coach1
                    ![Coach2] = s_Coach2
                    ![Special] = s_Variance
                    ![Type] = s_ClassType
                    
                    .Update
                End With
                    
                Set o_Recordset = Nothing
                    
            Next
            
        Next
        
    Next
        
End Function
Private Function GetScheduleHTML(ByVal ID As String) As HTMLDocument

    Dim doc As New HTMLDocument
    Dim Payload, _
        UpdateURl, _
        AdminUrl, _
        GetUrl      As String
    Dim vari As Variant
    
    ' * ' URLS
    AdminUrl = "https://clubready.com/common/scheduling/ajax_showadminclassweek.asp"
    UpdateURl = "https://clubready.com/common/scheduling/ajax_updateclassweek.asp"
    
    
    ' * ' Create FormData
    Payload = "id=" & Encode(ID) & "&dy=" & Encode(DateTime.Now) & "&r=404126"
    
    ' * ' Post
    With XMLClient
        .Open "POST", AdminUrl, False
        .send Payload
        
        .Open "POST", UpdateURl, False
        .send Payload
    End With
    
    ' * ' Response
    vari = ParseJScriptResponse(XMLClient.responseText)
    
    GetUrl = "http://clubready.com/common/scheduling/cached/ajax_classweekgrid.asp?cache=" & vari(0) & "&" & Encode(vari(1)) & "&" & Encode(vari(2)) & "&cid=" & Encode(ID) & "&" & vari(4)
    
    With XMLClient
        .Open "GET", GetUrl
        .send
    End With
        
    
    ' * ' Set DOc
    doc.body.innerHTML = XMLClient.responseText
    
    ' * ' Return
    Set GetScheduleHTML = doc
        
End Function
' ----------------------------------------------------------------
' Procedure Name: ProcessRefundData
' Purpose: Analyzes all Refunds and puts invoices into categories
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Private Sub ProcessRefundData()
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: GetSessionValues
' Purpose: Gets StoreID and Staff Type from Dashboard HTML Doc
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter objHTMLDocument (HTMLDocument): HTML Doc to Parse
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Private Sub GetSessionValues(objHTMLDocument As HTMLDocument)
    
    Call GetStoreID(objHTMLDocument)
    ''Call GetStaffType(objHTMLDocument)
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: GetStaffType
' Purpose: Get Staff type to verify permissions later
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter objHTMLDocument (HTMLDocument): Dashboard HTML DOC to parse
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Private Sub GetStaffType(objHTMLDocument As HTMLDocument)

    Dim objParentElement    As HTMLDivElement
    Dim objChildElement     As HTMLDivElement
    
    Set objParentElement = objHTMLDocument.getElementsByClassName("fr")(1)
    Set objChildElement = objParentElement.getElementsByTagName("div")(1)
    
    ' * ' Set Permission Values
    Select Case Trim(objChildElement.innerHTML)
        
        Case "Div 1 Corp Operations Manager"
            m_bCanAccessReports = True
            m_bCanChangeMemExp = True
            
        Case "Div 1 Corp Asst Ops Manager"
            m_bCanAccessReports = True
            m_bCanChangeMemExp = False
            
        Case "Div 1 Corp Service Rep"
            m_bCanAccessReports = False
            m_bCanChangeMemExp = False
    
    End Select
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: GetStoreID
' Purpose: Parses HTML DOCUMENT of CR dashboard to get Store ID
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter objHTMLDocument (HTMLDocument): Dashboard HTMLDOC to parse
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Private Sub GetStoreID(objHTMLDocument As HTMLDocument)

    Dim objElements As IHTMLElementCollection
    Dim objElement  As HTMLScriptElement
    
    Set objElements = objHTMLDocument.getElementsByTagName("script")
    
    ' * ' Check each Script element and find the one with StoreID Var in it
    For Each objElement In objElements
        i_Check = InStr(objElement.innerHTML, "dashStoreID")
        
        If i_Check > 0 Then
            
            ' * ' Extract the StoreID string from innterHTML
            For Counter = i_Check To Len(objElement.innerHTML)
                If Mid(objElement.innerHTML, Counter, 1) = ";" Then
                    Str1 = Mid(objElement.innerHTML, i_Check, Counter - i_Check)
                    str2 = Trim(Replace(Mid(Str1, InStr(Str1, "'") + 1, InStrRev(Str1, "'") - InStr(Str1, "'") + 1), "'", ""))
                    m_sStoreID = str2
                    Counter = Len(objElement.innerHTML)
                End If
                    
            Next
        End If
    Next
    

End Sub

' ----------------------------------------------------------------
' Procedure Name: LogIn
' Purpose: Takes given log in information and attempts to log into clubready
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sUserName (String):
' Parameter sPassword (String):
' Return Type: Boolean
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Public Function LogIn(sUserName As String, sPassword As String) As Boolean

    Dim bResult     As Boolean
    Dim sUrl        As String
    
    
    ' ' Set Url
    sUrl = "http://clubready.com/Security/Login?nl=" & _
            "&uid=" & sUserName & _
            "&pw=" & sPassword & _
            "&subdom=clubready&defaulthome=&storeID=&inst=&nosub=&redirectto=http://www.clubready.com/ufcgym.asp"
    
    ' ' Log In
    XMLClient.Open "GET", sUrl, False
    XMLClient.send
    
    Do While XMLClient.readyState <> 4
        DoEvents
    Loop
    
    Dim objDoc          As New HTMLDocument
    
    ' * ' Check If LogIn was Succesful
    If InStr(XMLClient.responseText, "Invalid Username") > 0 Then
        bResult = False
    ElseIf InStr(XMLClient.responseText, "You Can Access Several Locations") Then
        Dim frmLocations As New frmLocationSelect
        Dim oDoc As New HTMLDocument
        oDoc.body.innerHTML = XMLClient.responseText
        
        
        frmLocations.SelectLocation oDoc
        
        XMLClient.Open "GET", "http://clubready.com/Security/GoLocation/" & frmLocations.StoreID
        XMLClient.send
        bResult = True
        
    Else:
        objDoc.body.innerHTML = XMLClient.responseText
        bResult = True
        
    End If
    
    Dim objElements     As IHTMLElementCollection
    Dim sHeaders        As String
    Dim iStartCookie    As Integer
    
    
    ' * ' Set Document To HTML Response from client
    objDoc.body.innerHTML = XMLClient.responseText
    
    ' * ' Set Session ID / Staff Type
    Call GetSessionValues(objDoc)
    
    ' * ' Get Session Cookie
    sHeaders = XMLClient.getAllResponseHeaders
    iStartCookie = InStr(getHeader, "s3=")
    m_sSessionCookie = Mid(getHeader, startCookie + 3, 100)
    
    LogIn = bResult

End Function

' ----------------------------------------------------------------
' Procedure Name: UpdateMemberExpiration
' Purpose: Sets a members expiration to a given date
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sMemberID (String): Member Number
' Parameter dMemberSince (Date): Member start date
' Parameter dNewExp (Date): Date to be set
' Parameter dOldExp (Date): Date currently set
' Author: Matthew Rodriguez
' Date: 10/13/2017
' ----------------------------------------------------------------
Public Function UpdateMemberExpiration(ByVal sMemberID As String, ByVal dMemberSince As Date, ByVal dNewExp As Date, ByVal dOldExp As Date) As Boolean

    On Error GoTo UpdateMemberExpiration_Error
    Dim sUrl        As String
    Dim sFormData   As String
    
    sUrl = "http://clubready.com/admin/updatememberstatus.asp"
    
    ' * ' Form Data
    sFormData = "userId=" & sMemberID & _
                "&memberSince=" & Format(dMemberSince, "m/d/yyyy") & _
                "&memberSince_old=" & Format(dMemberSince, "m/d/yyyy") & _
                "&membershipExpires=" & Format(dNewExp, "m/d/yyyy") & _
                "&membershipExpires_old=" & Format(dOldExp, "m/d/yyyy") & _
                "&prospectTypeId=0&custNote=&r=552735"
               
    ' * ' Format
    sFormData = Replace(sFormData, "/", "%2F")
    
    
    XMLClient.Open "POST", sUrl
    XMLClient.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    XMLClient.setRequestHeader "Set-Cookie", SessionCookie
    XMLClient.send sFormData
    
    
    On Error GoTo 0
    UpdateMemberExpiration = True
    Exit Function

UpdateMemberExpiration_Error:

    UpdateMemberExpiration = False

End Function

Public Function AddMemberTask(sMemberID As String, sStaffID As String, sTaskTitle As String, dTaskDate As Date, iPriority As TaskPriority, iPurpose As TaskPurpose, iColor As TaskColor)
    
    Dim sUrl As String
    Dim sFormData As String
    
    sUrl = "https://clubready.com/common/user-tasks/ajax_addcusttask.asp"
    
    ' * ' Form Data
    sFormData = "todotitle=" & sTaskTitle & _
                "&tododue=7&tododueentry=" & dTaskDate & _
                "&todoassignedto=" & sStaffID & _
                "&color=" & iColor & _
                "&todopurpose=" & iPurpose & _
                "&todopriority=" & iPriority & _
                "&id=" & sMemberID
                
    ' * ' Format
    sFormData = Replace(Replace(sFormData, "/", "%2F"), " ", "+")

    ' * ' Send
    XMLClient.Open "POST", sUrl
    XMLClient.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    XMLClient.setRequestHeader "Set-Cookie", SessionCookie
    XMLClient.send sFormData
     
    AddMemberTask = (InStr(XMLClient.responseText, "Added Successfully") <> 0)
    
End Function
