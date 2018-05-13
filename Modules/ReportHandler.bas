Attribute VB_Name = "ReportHandler"
Sub CreateTapList(ByVal dDtm As Date)
    Dim Report      As New Report
    Dim cr          As New ClubReadyClient
    Dim TapList     As New TapList
    Dim FromDate, _
        ToDate      As Date
    
    
    
    ' * ' Get Date Range From Last TAP
    If Not day(dDtm) = 1 Then
        FromDate = DateAdd("d", -4, dDtm)
        ToDate = dDtm
    Else:
        FromDate = dDtm
        ToDate = dDtm
        
    End If
    
    ' * ' Download Reports
    Report.DownloadReport "Invoices Coming Due", FromDate, ToDate
    Report.DownloadReport "Club Past Due"
    
    ' * ' Make TapList
    TapList.InitializeWithValues FromDate, ToDate
    
    
End Sub

Sub CreateReportWithTemplate(ByVal ReportType As String, ByVal OutputPath As String)
    
    Select Case ReportType
        Case "Tap List":
    End Select
    
End Sub
