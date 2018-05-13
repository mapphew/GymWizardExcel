VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_sDownloadUrl As String
Private p_sFileName     As String
Private p_bClearTable   As Boolean

Public Property Get ClearTable()
    ClearTable = p_bClearTable
End Property
Public Property Get FileName() As String
    FileName = p_sFileName
End Property

Public Property Get DownloadUrl() As String
    DownloadUrl = p_sDownloadUrl
End Property



Public Sub InitWithVals(ByVal sStoreID As String, ByVal dtmFromDate As Date, ByVal dtmToDate As Date, Optional ByVal sFileName As String = "Attendance.xls", Optional ByVal bClearTable As Boolean = True)
                                
    
    ' * ' Create URL
    p_sDownloadUrl = "https://direct.clubready.com/reports/excel/83/?storeID=" & sStoreID & _
                        "&fromdate=" & dtmFromDate & _
                        "&todate=" & dtmToDate & "&chainID=&divisionID=&districtID=&genderfilter=&agefilter=&membershiptypeID=&customtypes=&dayofweek=&customTextFieldId=644&prospectTypeId="
    
    ' * ' File Name
    p_sFileName = sFileName

    p_bClearTable = bClearTable
    
End Sub

