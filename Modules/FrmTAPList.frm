VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTAPList 
   Caption         =   "Tap List Generator"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmTAPList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTAPList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub DayNumList_Change()
    ColorTap FrmTAPList.MonthNumList.Value
    DateCheck
End Sub

Private Sub GenButton_Click()
    GenButton.Enabled = False
    
    Dim dDate As Date
    Dim dFromDate, _
        dToDate, _
        dNextTapDate    As Date
        
    
    dDate = DateSerial(YearTextBox.Value, MonthNumList.Value, DayNumList.Value)
    Application.ScreenUpdating = False
    
    LoadLabel.Caption = "Generating Please Wait...."
    LoadLabel.Visible = True
    
    ' * ' Get Date Range
    Select Case day(dDate)
        Case 1:
            dFromDate = dDate
        Case 28:
            dFromDate = DateAdd("d", -2, dDate)
        Case 5, 29:
            dFromDate = DateAdd("d", -3, dDate)
        Case 31:
            dFromDate = DateAdd("d", -5, dDate)
        Case Else:
            dFromDate = DateAdd("d", -4, dDate)
    End Select
    
    dToDate = dDate
    
    ' * ' Next Tap Date
    Select Case day(dToDate)
        Case 1:
            dNextTapDate = DateAdd("d", 4, dToDate)
            
        Case 5, 10, 15, 20:
            dNextTapDate = DateAdd("d", 5, dToDate)
            
        Case Else:
            dNextTapDate = DateSerial(Year(DateAdd("m", 1, dToDate)), Month(DateAdd("m", 1, dToDate)), 1)
            
    End Select
    
    ' * ' Download Reports
    Dim r_ICD As New Report_Invoices_Coming_Due
    Dim r_CPD As New Report_Club_Past_Due
    Dim r_TAPList As New TapList
    Dim o_Client As New ClubReadyClient
    
    o_Client.LogIn "rmatt", "9672"
    
    r_ICD.InitWithVals o_Client.StoreID, dFromDate, dNextTapDate
    r_CPD.InitWithVals o_Client.StoreID
    
    o_Client.DownloadReport r_ICD
    o_Client.DownloadReport r_CPD
    
    r_TAPList.InitializeWithValues dFromDate, dToDate, dNextTapDate
    PathLabel.Caption = r_TAPList.ReportPath
    
    LoadLabel.Caption = "Complete!"
    GenButton.Enabled = True
    
    Application.ScreenUpdating = True
    OpenButton.Visible = True
    
    
    
    
    
End Sub

Private Sub MonthNumList_Change()
    PopulateDay FrmTAPList.MonthNumList.Value
    DateCheck
    
End Sub


Private Sub OpenButton_Click()
    Dim pPath As String
    pPath = PathLabel.Caption
    
    Dim book As Workbook
    Set book = Workbooks.Open(pPath)
    
    Unload FrmTAPList
End Sub

Public Sub UserForm_Activate()

    ' * ' Set Values
    FrmTAPList.MonthNumList.List = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    FrmTAPList.MonthNumList.ListIndex = Month(DateTime.Now) - 1
    
    FrmTAPList.DayNumList.ListIndex = day(DateTime.Now) - 1
    FrmTAPList.YearTextBox.Value = Year(Now)
    
    
End Sub

Private Sub PopulateDay(ByVal MonthNum As Integer)

    Dim LastDay As Integer
    Dim Year As Integer
    Dim x As Integer
    
    
    Year = FrmTAPList.YearTextBox.Value
    LastDay = day(DateAdd("d", -1, DateSerial(Year, MonthNum + 1, 1)))
    
    FrmTAPList.DayNumList.List = Array()
    
    For x = 1 To LastDay
        FrmTAPList.DayNumList.AddItem (x)
    Next
    
    ColorTap MonthNum
End Sub

Private Sub ColorTap(ByVal MonthNum As Integer)
    Dim Year As Integer
    Dim LastDay As Integer
    
    Year = FrmTAPList.YearTextBox.Value
    LastDay = day(DateAdd("d", -1, DateSerial(Year, MonthNum + 1, 1)))
    
    Select Case FrmTAPList.DayNumList.Value
        
        Case 1, 5, 10, 15, 20, 25:
            FrmTAPList.DayNumList.ForeColor = vbBlack
            FrmTAPList.WarningLabel.Visible = False
            FrmTAPList.GenButton.Visible = True
        Case LastDay:
            FrmTAPList.DayNumList.ForeColor = vbBlack
            FrmTAPList.WarningLabel.Visible = False
            FrmTAPList.GenButton.Visible = True
            
        Case Else:
            FrmTAPList.DayNumList.ForeColor = vbRed
            WarningLabelGen 1
    End Select
    
End Sub
Private Sub DateCheck()

    If DateSerial(FrmTAPList.YearTextBox.Value, FrmTAPList.MonthNumList, FrmTAPList.DayNumList.Value) > DateTime.Now Then
        WarningLabelGen 2
    End If
    
End Sub
Private Sub WarningLabelGen(ByVal errnum As Integer)
    Dim warningtext As String
    
    FrmTAPList.GenButton.Visible = False
    
    Select Case errnum
        Case 1:
            warningtext = "*** The Selected Date Is Not A Valid TAP Date"
        Case 2:
            warningtext = "Date is In The Future"
    End Select
          
    FrmTAPList.WarningLabel.Caption = warningtext
End Sub

Private Sub UserForm_Deactivate()
    Unload FrmTAPList
End Sub

Public Sub UserForm_Initialize()
    FrmTAPList.MonthNumList.List = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    FrmTAPList.MonthNumList.ListIndex = Month(DateTime.Now) - 1
    
    FrmTAPList.DayNumList.ListIndex = day(DateTime.Now) - 1
End Sub
