VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sMemberNumber As String
Private m_objMembershipAgreement As Class_Member_Agreement
Private m_objCRClient As ClubReadyClient
Private m_Agreements() As Variant
Private m_dtMemberExpirationDate As Date
Private m_bIsActiveMember As Boolean

Private m_bIsIn3rdParty As Boolean

Public Property Get Agreements() As Variant(): Agreements = m_Agreements: End Property

Private Property Get CRClient() As ClubReadyClient: Set CRClient = m_objCRClient: End Property

Private Property Set CRClient(ByVal objNewValue As ClubReadyClient): Set m_objCRClient = objNewValue: End Property

Public Property Get IsActiveMember() As Boolean: IsActiveMember = m_bIsActiveMember: End Property

Public Property Get IsIn3rdParty() As Boolean: IsIn3rdParty = m_bIsIn3rdParty: End Property

Public Property Get MemberExpirationDate() As Date: MemberExpirationDate = m_dtMemberExpirationDate: End Property

Public Property Get MemberNumber() As String: MemberNumber = m_sMemberNumber: End Property

Public Property Let MemberNumber(ByVal sNewValue As String): m_sMemberNumber = sNewValue: End Property

Public Property Get MembershipAgreement() As Class_Member_Agreement: MembershipAgreement = m_objMembershipAgreement: End Property

Public Sub Class_InitializeWithValues(ByVal objCRClient As ClubReadyClient, ByVal sMemberID)
    
    Set CRClient = objCRClient
    MemberNumber = sMemberID
    
    m_bIsActiveMember = False
    m_bIsIn3rdParty = False
    
    ' * ' Get Membership
    Call GetAgreements
    
    
    
End Sub

Private Sub GetAgreements()
    
    ' * '
    Dim sAgreements As Variant
    Dim lItr As Long
    
    sAgreements = CRClient.GetMemberAgreements(MemberNumber)
    ReDim Preserve m_Agreements(1 To UBound(sAgreements))
    
    For lItr = 1 To UBound(sAgreements)
        Dim objAgreement As New Class_Member_Agreement
        Dim sTitle As String
        Dim dtSignDate As Date
        Dim sSalesPerson1 As String
        Dim sSalesPerson2 As String
        Dim sAgreementNumber As String
        Dim sAgreementStatus As String
        
        dtSignDate = sAgreements(lItr)(1)
        sAgreementStatus = sAgreements(lItr)(2)
        sSalesPerson1 = sAgreements(lItr)(3)
        sSalesPerson2 = sAgreements(lItr)(4)
        sAgreementNumber = sAgreements(lItr)(5)
        sTitle = sAgreements(lItr)(6)
    
        objAgreement.Class_InitializeWithValues sTitle, dtSignDate, sSalesPerson1, sSalesPerson2, sAgreementNumber, sAgreementStatus
        
        Set m_Agreements(lItr) = objAgreement
        
        Set objAgreement = Nothing
        
    Next
    
    ' * ' Get Active Membership
    
    For lItr = 1 To UBound(m_Agreements)
        Set objAgreement = m_Agreements(lItr)
        
        If objAgreement.AgreementType = "Membership" Then
        
            ' * ' Check For PIF
            If objAgreement.IsPaidInFull And objAgreement.AgreementStatus <> "Cancelled" Then
                Set m_objMembershipAgreement = objAgreement
                
                ' * ' Now GET Expiration of Contract and see if its in the future
                If m_objMembershipAgreement.AgreementExpirationDate > Now Then
                   m_bIsActiveMember = True
                   m_dtMemberExpirationDate = m_objMembershipAgreement.AgreementExpirationDate
                End If
                Exit For
            End If
            
            ' * ' Check IF EFT Membership that's Active
            If Not objAgreement.IsPaidInFull And objAgreement.AgreementStatus = "Active EFT" Then
                Set m_objMembershipAgreement = objAgreement
                m_bIsActiveMember = True
                
                Exit For
            End If
            
            ' * ' Check for Collections
            If objAgreement.AgreementStatus = "Third Party" Then
                m_bIsIn3rdParty = True
                Set m_objMembershipAgreement = objAgreement
                Exit For
            End If
            
        End If
    Next
    
        

End Sub
