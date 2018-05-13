VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRedScreens 
   Caption         =   "Red Screen Date Selection"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10140
   OleObjectBlob   =   "frmRedScreens.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRedScreens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GenButton_Click()
    
    ' * ' Now Get Dates For The Reports
    Dim dtFromDate As Date
    Dim dtToDate As Date
    
    With frmRedScreens
        dtFromDate = DateSerial(YearTextBox1.Value, MonthNumList1.Value, DayNumList1.Value)
        dtToDate = DateSerial(YearTextBox2.Value, MonthNumList2.Value, DayNumList2.Value)
    End With
    
    ' * ' Create Generator
    Dim rsGen As New RedScreenGenerator
    rsGen.Generate dtFromDate, dtToDate
    
End Sub

Public Sub UserForm_Initialize()
    frmRedScreens.MonthNumList1.List = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    frmRedScreens.MonthNumList2.List = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    
    frmRedScreens.MonthNumList1.ListIndex = Month(DateTime.Now) - 1
    frmRedScreens.MonthNumList2.ListIndex = Month(DateTime.Now) - 1
    
    frmRedScreens.DayNumList1.ListIndex = day(DateTime.Now) - 2
    frmRedScreens.DayNumList2.ListIndex = day(DateTime.Now) - 2
    
    frmRedScreens.YearTextBox1.Value = Year(Now)
    frmRedScreens.YearTextBox2.Value = Year(Now)
End Sub

Private Sub MonthNumList1_Change()

    PopulateDay frmRedScreens.MonthNumList1.Value, 1
    
End Sub
Private Sub MonthNumList2_Change()

    PopulateDay frmRedScreens.MonthNumList2.Value, 2
    
End Sub
Private Sub PopulateDay(ByVal MonthNum As Integer, ByVal listNum As Integer)

    Dim LastDay As Integer
    Dim Year As Integer
    Dim x As Integer
    
    If listNum = 1 Then
        Year = frmRedScreens.YearTextBox1.Value
        LastDay = day(DateAdd("d", -1, DateSerial(Year, MonthNum + 1, 1)))
        
        frmRedScreens.DayNumList1.List = Array()
        For x = 1 To LastDay
            frmRedScreens.DayNumList1.AddItem (x)
        Next
    Else
        Year = frmRedScreens.YearTextBox2.Value
        LastDay = day(DateAdd("d", -1, DateSerial(Year, MonthNum + 1, 1)))
        
        frmRedScreens.DayNumList2.List = Array()
        For x = 1 To LastDay
            frmRedScreens.DayNumList2.AddItem (x)
        Next
    End If
    
    
End Sub
