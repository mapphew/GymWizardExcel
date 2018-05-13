VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSchedules 
   Caption         =   "Schedule Generator"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmSchedules.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GenButton_Click()
    Dim Youth, _
        Adult       As Boolean
    Dim Shell As New Shell
    
    Adult = FrmSchedules.CheckBoxAdult.Value
    Youth = FrmSchedules.CheckBoxYouth.Value
    
    Dim o_Client As New ClubReadyClient
    o_Client.LogIn g_objGWClient.UserName, g_objGWClient.Password
    o_Client.CreatePrintableSchedule Youth, Adult
    
    Select Case MsgBox("Open Schedule Folder?", vbYesNo, "Schedules Completed!")
        
        Case vbYes:
            Shell.Open g_objGWClient.SchedulePath
            Unload FrmSchedules
        Case vbNo:
            Unload FrmSchedules
    End Select
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()
    Unload FrmSchedules
End Sub

Private Sub UserForm_Initialize()
' * ' Check For Globals
    If g_objGWClient Is Nothing Then
        Set g_objGWClient = New GymWizardClient
    End If
End Sub
