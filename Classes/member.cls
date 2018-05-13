VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pUserID, pFirstName, pLastName, pMemberType, pStatus, pHomeClub As String
Private pStartDate, pExpiration, pEndDate, pLastPayDate, pNextDraftDate As Variant
Private pBalance, pLastPayAmount, pDueNext30 As Currency

Public Property Let LastName(ByVal Value As String)
    pLastName = Value
End Property

Public Property Let FirstName(ByVal Value As String)
    pFirstName = Value
End Property

Public Property Let UserID(ByVal Value As String)
    pUserID = Value
End Property

Public Property Let MembershipType(ByVal Value As String)
    pMemberType = Value
End Property

Public Property Let HomeClub(ByVal Value As String)
    pHomeClub = Value
End Property

Public Property Let Status(ByVal Value As String)
    pStatus = Value
End Property

Private Sub Process()
    
    If pStatus = "Unknown" Then
        ' ' Get Status
        If pBalance > 0 Then
            pStatus = "Past Due"
            
        ElseIf pExpiration > DateTime.Now Then
            pStatus = "Active"
        
        ElseIf IsNull(pNextDraftDate) Then
            If Not IsNull(pEndDate) Then
                pStatus = "Cancelled"
            Else:
                pStatus = "Review"
            End If
        Else:
            pStatus = "Review"
        End If
    End If
    
    
    
    
End Sub

Public Sub InitializeWithRecord(ByVal Record As Variant)
    
    
    ' ' Asign Variables
    With Record
        pFirstName = ![FirstName]
        pLastName = ![LastName]
        pUserID = ![UserID]
        pStartDate = ![startDate]
        pExpiration = ![Expiration]
        pEndDate = ![EndDate]
        pMemberType = ![MemberType]
        pBalance = ![Balance]
        pLastPayDate = ![LastPayDate]
        pLastPayAmount = ![LastPayAmount]
        pDueNext30 = ![DueNext30]
        pNextDraftDate = ![NextDraftDate]
        pStatus = ![Status]
    End With
    
    
    Call Process
    
End Sub
