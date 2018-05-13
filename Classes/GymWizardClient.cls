VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GymWizardClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_sDatabasePath As String
Private m_sTempFolderPath As String
Private m_sGWFolderPath As String
Private m_sTapListFolderPath As String
Private m_sRedScreenFolderPath As String
Private m_sScheduleFolderPath As String
Private m_sUserName As String
Private m_sPassword As String

Private Sub Class_Initialize()
    
    ' * ' Set Path Vars
    m_sGWFolderPath = Environ("USERPROFILE") & "\My Documents\Gym Wizard\"
    m_sDatabasePath = GWFolderPath & "GWDatabase.accdb"
    m_sTempFolderPath = Environ("temp") & "\"
    m_sTapListFolderPath = GWFolderPath & "TAP Lists\"
    m_sScheduleFolderPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\Class Schedules\"
    m_sRedScreenFolderPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\Red Screens\"
    
    
      
End Sub
Public Sub GetLogInInfo()

    ' * ' Set Log In Credentials
    Dim objDatabase1 As DAO.Database
    Dim Recordset1, _
        Recordset2     As DAO.Recordset
        
    Set objDatabase1 = OpenDatabase(m_sDatabasePath)
    Set Recordset1 = objDatabase1.OpenRecordset("SELECT * FROM [User Info] WHERE [Variable] = 'Admin_UserName';")
    Set Recordset2 = objDatabase1.OpenRecordset("SELECT * FROM [User Info] WHERE [Variable] = 'Admin_Password';")
    
    m_sUserName = Recordset1![Value]
    m_sPassword = Recordset2![Value]
    
    Set Recordset1 = Nothing
    Set Recordset2 = Nothing
    Set objDatabase1 = Nothing
End Sub
' ----------------------------------------------------------------
' Procedure Name: StartUpCheck
' Purpose: Checks For and/or creates Required Files and Folders
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Public Sub StartUpCheck()
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    
    ' * ' Folders
    If Not FolderExists(m_sGWFolderPath) Then
        FSO.CreateFolder m_sGWFolderPath
    End If
    
    If Not FolderExists(m_sTapListFolderPath) Then
        FSO.CreateFolder m_sTapListFolderPath
    End If
    
    If Not FolderExists(m_sScheduleFolderPath) Then
        FSO.CreateFolder m_sScheduleFolderPath
    End If
    
    If Not FolderExists(m_sRedScreenFolderPath) Then
        FSO.CreateFolder m_sRedScreenFolderPath
    End If
        
    ' * ' DATABASE
    Dim objDatabase As GWDatabase
    Set objDatabase = New GWDatabase
    
    ' * ' Check For First Time Log In
    If Not objDatabase.SetUpCheck Then
        Dim frmSetUp As New frmLogInSetUp
        Load frmSetUp
    End If
    
    
    
End Sub

Public Property Get TAPListPath() As String

    TAPListPath = m_sTapListFolderPath

End Property

Public Property Get SchedulePath() As String

    SchedulePath = m_sScheduleFolderPath

End Property
Public Property Get RedScreenPath() As String
    RedScreenPath = m_sRedScreenFolderPath
End Property

Public Property Get DatabasePath() As String

    DatabasePath = m_sDatabasePath

End Property

Public Property Get GWFolderPath() As String

    GWFolderPath = m_sGWFolderPath

End Property

Public Property Get TempFolderPath() As String

    TempFolderPath = m_sTempFolderPath

End Property


Private Function AgreementCheck(ByVal m_oClient As ClubReadyClient)

End Function

Public Property Get UserName() As String

    UserName = m_sUserName

End Property

Public Property Get Password() As String

    Password = m_sPassword

End Property
