VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Member_Agreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sAgreementName As String
Private m_dtmAgreementSignDate As Date
Private m_sSalesPerson1 As String
Private m_sSalesPerson2 As String
Private m_bIsPaidInFull As Boolean
Private m_sAgreementType As String
Private m_sAgreementStatus As String
Private m_sAgreementNumber As String
Private m_dtAgreementExpirationDate As Date

Public Property Get AgreementExpirationDate() As Date: AgreementExpirationDate = m_dtAgreementExpirationDate: End Property

Public Property Get AgreementName() As String: AgreementName = m_sAgreementName: End Property

Public Property Get AgreementNumber() As String: AgreementNumber = m_sAgreementNumber: End Property

Public Property Get AgreementSignDate() As Date: AgreementSignDate = m_sAgreementSignDate: End Property

Public Property Get AgreementStatus() As String: AgreementStatus = m_sAgreementStatus: End Property

Public Property Get AgreementType() As String: AgreementType = m_sAgreementType: End Property

Public Property Get IsPaidInFull() As Boolean: IsPaidInFull = m_bIsPaidInFull: End Property

Public Property Get SalesPerson1() As String: SalesPerson1 = m_sSalesPerson1: End Property

Public Property Get SalesPerson2() As String: SalesPerson2 = m_sSalesPerson2: End Property

Public Sub Class_InitializeWithValues(ByVal sAgreementName As String, ByVal dtmAgreementSignDate As Date, ByVal sSalesPerson1 As String, ByVal sSalesPerson2 As String, ByVal sAgreementNumber As String, ByVal sAgreementStatus As String)
    
    ' * '
    m_sAgreementName = sAgreementName
    m_sSalesPerson1 = sSalesPerson1
    m_sSalesPerson2 = sSalesPerson2
    m_dtmAgreementSignDate = dtmAgreementSignDate
    m_sAgreementStatus = sAgreementStatus
    m_sAgreementNumber = sAgreementNumber
    
    
    ' * ' Check Type Of Package
    If InStr(sAgreementName, "PT/MMA") Or InStr(sAgreementName, "Per Week") Then
        m_sAgreementType = "Serivce"
    Else:
        m_sAgreementType = "Membership"
    End If
    
    If m_sAgreementType = "Membership" Then
        
        If InStr(m_sAgreementName, "Month") Or InStr(m_sAgreementName, "YR") Or InStr(m_sAgreementName, "YEAR") Then
            m_bIsPaidInFull = True
            GetPIFDuration
        Else:
            m_bIsPaidInFull = False
        End If
    Else:
    End If
            
    
    
End Sub

Private Sub GetPIFDuration()
    Dim sSplitArray() As String
    Dim sNewArray() As String
    Dim lYears As Long
    Dim lMonths As Long
    Dim lItr As Long
    
    lYears = 0
    lMonths = 0
    
    If InStr(m_sAgreementName, "YR") Then
        sSplitArray = Split(m_sAgreementName, " ")
        
        For lItr = 0 To UBound(sSplitArray)
            If sSplitArray(lItr) = "YR" Then
                lYears = lYears + sSplitArray(lItr - 1)
            ElseIf InStr(sSplitArray(lItr), "YR") Then
                lYears = lYears + Replace(sSplitArray(lItr), "YR", "")
                
            End If
        Next
    End If
    
    If InStr(m_sAgreementName, "Month") And Not InStr(m_sAgreementName, "Months") Then
        sSplitArray = Split(m_sAgreementName, " ")
        
        For lItr = 0 To UBound(sSplitArray)
            If sSplitArray(lItr) = "Month" Then
                lMonths = lMonths + sSplitArray(lItr - 1)
            ElseIf InStr(sSplitArray(lItr), "Month") Then
                lMonths = lMonths + Replace(sSplitArray(lItr), "Month", "")
                
            End If
        Next
    End If
    
    If InStr(m_sAgreementName, "Months") Then
        sSplitArray = Split(m_sAgreementName, " ")
        
        For lItr = 0 To UBound(sSplitArray)
            If sSplitArray(lItr) = "Months" Then
                lMonths = lMonths + sSplitArray(lItr - 1)
            ElseIf InStr(sSplitArray(lItr), "Months") Then
                lMonths = lMonths + Replace(sSplitArray(lItr), "YMonths", "")
                
            End If
        Next
    End If
    
    lMonths = lMonths + (lYears * 12)
    
    m_dtAgreementExpirationDate = DateAdd("m", lMonths, m_dtmAgreementSignDate)
        

End Sub
