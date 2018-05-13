VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogInSetUp 
   Caption         =   "Log In Setup"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4470
   OleObjectBlob   =   "frmLogInSetUp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogInSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSubmit_Click()
    
    ' * ' Check That Passwords Match
    If LCase(txtBoxPassword1.Value) <> LCase(txtBoxPassword2.Value) Then
        lblErrorMessage.Caption = "Password's Don't Match!"
        lblErrorMessage.ForeColor = RGB(255, 0, 0)
    Else:
        Dim objDatabase         As DAO.Database
        Dim objRecordset1, _
            objRecordset2       As DAO.Recordset
            
        Set objDatabase = OpenDatabase(g_objGWClient.DatabasePath)
        Set objRecordset1 = objDatabase.OpenRecordset("SELECT * FROM [User Info] WHERE [Variable] = 'Admin_UserName';")
        Set objRecordset2 = objDatabase.OpenRecordset("SELECT * FROM [User Info] WHERE [Variable] = 'Admin_Password';")
        
        objRecordset1.Edit
        objRecordset1![Value] = LCase(txtBoxUserName.Value)
        
        objRecordset2.Edit
        objRecordset2![Value] = LCase(txtBoxPassword1.Value)
        
        objRecordset1.Update
        objRecordset2.Update
        
        MsgBox "Information Saved Successfully!", vbOKOnly, Application.Name
        
        Unload frmLogInSetUp
    End If
    
End Sub

Private Sub UserForm_Initialize()
    ' * ' Check For Globals
    If g_objGWClient Is Nothing Then
        Set g_objGWClient = New GymWizardClient
    End If
    
    
End Sub
