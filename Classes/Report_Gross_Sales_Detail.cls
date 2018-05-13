VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Gross_Sales_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private p_sDownloadUrl  As String
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

Public Sub InitWithVals(ByVal dFromDate As Date, ByVal dToDate As Date, ByVal sStoreID As String, Optional ByVal sFileName As String = "GSD.xls", Optional ByVal bClearTable As Boolean = True)
                                
    
    ' * ' Create URL
    p_sDownloadUrl = "http://direct.clubready.com/reports/excel/47/?" & _
                            "storeid=" & sStoreID & _
                            "&userid=&fromdate=" & Format(dFromDate, "mm/dd/yyyy") & _
                            "&todate=" & Format(dToDate, "mm/dd/yyyy") & "&chainid=&transaction=&saletype=&productcategoryid=&packageid="
                        
    ' * ' File Name
    p_sFileName = sFileName
    
    p_bClearTable = bClearTable

End Sub

