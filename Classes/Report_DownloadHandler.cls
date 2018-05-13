VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DownloadHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function BulkReportDownload(ByVal mReports As Variant, ByVal frm As UserFormProcessingScreen)

    
    ' * ' DIM
    Dim m_iReportCount  As Long
    Dim m_iReportNr     As Long
    Dim m_AsyncHandler As New C_AsyncHandler
    
    ' * ' Get Total
    m_iReportCount = UBound(mReports) + 1
    Call m_AsyncHandler.VBAMultiThread(m_iReportCount, frm, , mReports)
    
End Function
