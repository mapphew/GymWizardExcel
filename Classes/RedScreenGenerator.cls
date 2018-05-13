VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedScreenGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pDatabase As DAO.Database


Public Sub Generate(ByVal dtFromDate As Date, ByVal dtToDate As Date)

    Dim objDatabase As DAO.Database
    Dim oRecordSet As DAO.Recordset
    Dim objClient As New ClubReadyClient
    Dim objMemExp As New Report_Member_Export
    Dim objAtt As New Report_Attendance
    Dim objFrozen As New Report_Frozen_Members
    
    ' * ' Open DB
    Set objDatabase = OpenDatabase(g_objGWClient.DatabasePath)
    
    g_objGWClient.GetLogInInfo
    
    ' * ' Download Attendance
    objClient.LogIn g_objGWClient.UserName, g_objGWClient.Password
    
    objMemExp.InitWithVals objClient.StoreID
    objAtt.InitWithVals objClient.StoreID, dtFromDate, dtToDate
    objFrozen.InitWithVals objClient.StoreID
    
    objClient.DownloadReport objMemExp
    objClient.DownloadReport objAtt
    objClient.DownloadReport objFrozen
    
    
    Dim dtItr As Date
    dtItr = dtFromDate
    
    Do While dtItr <= dtToDate
        
        ' * ' Populate Record Set With RedScreens
        Set oRecordSet = objDatabase.OpenRecordset("SELECT * FROM [Attendance] WHERE FormatDateTime([Check In Date / Time], 2) = #" & dtItr & "# AND [Denied] = 'Yes';")
        
        ' * ' Create Report
        Dim rsReport As New RedScreenReport
        rsReport.Class_InitializeWithValues dtItr, oRecordSet, objClient
        
        ' * ' Add Day To itr
        dtItr = DateAdd("d", 1, dtItr)
    Loop
    
End Sub


Private Sub Class_Initialize()
    ' * ' Check For Globals
    If g_objGWClient Is Nothing Then
        Set g_objGWClient = New GymWizardClient
    End If
End Sub
