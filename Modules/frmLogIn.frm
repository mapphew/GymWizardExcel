VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogIn 
   Caption         =   "Log In"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmLogIn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_objClubReadyClient As ClubReadyClient


Private Sub btnLogIn_Click()
    Set m_objClubReadyClient = New ClubReadyClient
    Me.Hide
    If Not m_objClubReadyClient.LogIn(Me.txtBoxUserName, Me.txtBoxPassword) Then
        Me.lblErrorMessage = "Invalid User Name"
        Me.Show
    Else:
    End If
    
End Sub

Public Property Get ClubReadyClient() As ClubReadyClient: Set ClubReadyClient = m_objClubReadyClient: End Property
