VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Member_Export"
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



Public Sub InitWithVals(ByVal sStoreID As String, Optional ByVal sFileName As String = "MemExp.xlsx", Optional ByVal bClearTable As Boolean = True)
                                
    
    ' * ' Create URL
    p_sDownloadUrl = "http://direct.clubready.com/common/reportserverproxy/ssrsproxy.asp?rs%3AFormat=EXCELOPENXML&rspType=xlsx&rpt=Member+Export&pth=Ssrs.Web.Members&fileName=MembersExport&SearchCategory=2&StoreId=" & _
                        sStoreID & "&ExtraFields=1&SearchType=0&Gender=1&OtherFilter=13"
                        
    ' * ' File Name
    p_sFileName = sFileName

    p_bClearTable = bClearTable
    
End Sub

