VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Global_UFCDBPath As String
'---------------------------------------------------------------------------------------
' Method : Workbook_Open
' Author : mgrod
' Date   : 4/20/2016
' Purpose: Creates Proper Files In Report File Path For Use
'---------------------------------------------------------------------------------------
Private Sub Workbook_Open()
        
    ' * ' Set GW Client
    Set g_objGWClient = New GymWizardClient
    
    ' * ' Run StartupCheck
    g_objGWClient.StartUpCheck
    
    ' ' Set Active Worksheet
    ThisWorkbook.Sheets(1).Activate
    
    ' ' Set Today's Date
    Range("MONTH_INT").Value = Month(DateTime.Now)
    Range("DAY_INT").Value = day(DateTime.Now)
    Range("YEAR_INT").Value = Year(DateTime.Now)
    
    
    
    
End Sub



Public Sub VarTest()
    Dim newArr() As Variant
    newArr = GetVar
    
End Sub

Public Function GetVar() As Variant
    
    Dim testVar() As Variant
    
    ReDim testVar(2)
    
    testVar(0) = "test1"
    testVar(1) = "test2"
    testVar(2) = "test3"
    
    GetVar = testVar
End Function
