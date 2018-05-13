VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCheck 
   Caption         =   "Status Check"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "FrmCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "SkipMZToolsTabIndexReview"
End
Attribute VB_Name = "FrmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' File   : FrmCheck
' Author : mgrod
' Date   : 4/21/2016
' Purpose: Set the status of a RedScreen that wasn't automatically set
'---------------------------------------------------------------------------------------
Option Explicit

Private m_sNewStatus As String
Public Property Get NewStatus() As String: NewStatus = m_sNewStatus: End Property

Public Sub PopulateMemberInfo(ByVal strName As String, ByVal strID As String, ByVal strStatus As String, ByVal dtCheckIn As Date)
    
    Me.lbl_name2 = strName
    Me.lbl_status2 = strStatus
    Me.lbl_userid2 = strID
    Me.lbl_time2 = dtCheckIn
    Me.Show
    
End Sub

Private Sub btnNo_Click()
    m_sNewStatus = "Review"
    Me.Hide
End Sub

Private Sub btnYes_Click()
    m_sNewStatus = "Cleared"
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Me.Hide
End Sub
