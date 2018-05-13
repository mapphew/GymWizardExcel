VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Refunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_sDownloadUrl  As String
Private p_sFileName     As String
Private p_bClearTable   As Boolean

Public Property Get FileName() As String
    FileName = p_sFileName
End Property

Public Property Get DownloadUrl() As String
    DownloadUrl = p_sDownloadUrl
End Property

Public Property Get ClearTable()
    ClearTable = p_bClearTable
End Property

Public Sub InitWithVals(ByVal dFromDate As Date, ByVal dToDate As Date, ByVal sStoreID As String, Optional ByVal bClearTable As Boolean = True)
                                
    
    ' * ' Create URL
    p_sDownloadUrl = "https://direct.clubready.com/Reporting/RefundsReturns/RunRefundsReturns?" & _
                    "startDate=" & dFromDate & _
                    "&endDate=" & dToDate & _
                    "&divisionId=&districtId=&storeId=" & sStoreID & _
                    "&reportType="
                        
    ' * ' File Name
    p_sFileName = "Ref.xls"
    
    p_bClearTable = bClearTable

End Sub


