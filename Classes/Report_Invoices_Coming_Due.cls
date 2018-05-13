VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Invoices_Coming_Due"
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



Public Sub InitWithVals(ByVal sStoreID As String, ByVal dFromDate As Date, _
                                ByVal dToDate As Date, _
                                Optional sFileName As String = "ICD.xls", _
                                Optional bClearTable As Boolean = True, _
                                Optional icd_Collections As icdCollections = IncludeCollections, _
                                Optional icd_PayType As icdPayType = AnyPayType, _
                                Optional icd_PastDowns As icdInclusion = Include, _
                                Optional icd_AllPastInv As icdAllPastInvoices = IncludeAll, _
                                Optional icd_InvoiceType As icdInvoiceType = AllInvoices, _
                                Optional icd_CancelledInv As icdInclusion = Include)
                                
    
    ' * ' Create URL
    p_sDownloadUrl = Replace(Replace("http://direct.clubready.com/reports/excel/32/?" & _
                        "fromdate=" & dFromDate & _
                        "&todate=" & dToDate & _
                        "&nocollections=" & icd_Collections & _
                        "&paytype=" & icd_PayType & _
                        "&includepastdowns=" & icd_PastDowns & _
                        "&includeallpastinvoices=" & icd_AllPastInv & _
                        "&storeID=" & sStoreID & _
                        "&districtID=&divisionID=&chainId=&invoicetype=" & icd_InvoiceType & _
                        "&excludeCancelledInvoices=" & icd_CancelledInv, "1004", ""), "-1", "True")
                        
    ' * ' File Name
    p_sFileName = sFileName

    p_bClearTable = bClearTable
End Sub
                                

