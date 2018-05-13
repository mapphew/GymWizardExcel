Attribute VB_Name = "DashboardHandler"
Public Sub DailyChecks()

    ' ' Variables
    Dim RSArray(1 To 31) As Boolean
    Dim TapArray(1 To 31) As Boolean
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    ' ' Get Month On DashBoard
    mInt_MonthNum = Range("MONTH_INT")
    mStr_MonthName = MonthName(mInt_MonthNum)
    mInt_Year = Range("YEAR_INT")
    
    ' ' Get Last Day of Month
    mDtm_Date = DateTime.DateSerial(mInt_Year, mInt_MonthNum, 1)
    mDtm_Date = DateTime.DateAdd("m", 1, mDtm_Date)
    mDtm_Date = DateTime.DateAdd("d", -1, mDtm_Date)
    mInt_LastDay = day(mDtm_Date)
    
    ' ' Set Directory Paths
    mStr_RSDir = ThisWorkbook.Path & "\Reports\RedScreens\" & mInt_Year & "\" & mStr_MonthName
    mStr_TapDir = ThisWorkbook.Path & "\Reports\TapList\" & mInt_Year & "\" & mStr_MonthName
    
    ' ' Check For Reports
    For x = 1 To mInt_LastDay
    
        ' ' Annual Assesment Check
        If x = mInt_LastDay Then
            TapArray(x) = FSO.FileExists(mStr_TapDir & "\" & mInt_MonthNum & "." & x & "TAPList.xlsx")
        End If
        
        ' ' RedScreens
        RSArray(x) = FSO.FileExists(mStr_RSDir & "\" & mInt_MonthNum & "." & x & "RedScreens.xlsx")
        
        ' ' Tap List
        Select Case x
            
            Case 1, 5, 10, 15, 20, 25:
                TapArray(x) = FSO.FileExists(mStr_TapDir & "\" & mInt_MonthNum & "." & x & "TAPList.xlsx")
                                        
        End Select
        
    Next
    
    ' ' Check Mark Time
    For x = 1 To mInt_LastDay
    
        ' ' RedScreens
        If RSArray(x) Then
            Range("RS_R_" & x).Value = ChrW(&H2713)
        Else:
            Range("RS_R_" & x).Value = ""
        End If
        
        ' ' Tap
        If TapArray(x) Then
            Range("TAP_R_" & x).Value = ChrW(&H2713)
        Else:
            Range("TAP_R_" & x).Value = ""
        End If
    
    Next
    
    Call InstallCheck
End Sub

Sub InstallCheck()
'=======================================================================================
' Method   : Install Check
' Author   : Matthew Rodriguez
' Created  : 1/27/2017
' Purpose  : Checks for critical folders to insure workbook is installed correctly
' Called by: user
' Arguments:
' Comments :
' Changes-------------------------------------------------------------------------------
' Date            Programmer                Change
'
'=======================================================================================
    
    Dim FSO As New FileSystemObject
    Dim mStr_DocDir, _
        mStr_AppDir As String
    
    ' ' My Documents
    mStr_DocDir = Environ("USERPROFILE") & "\My Documents"
    
    ' ' Check For Gym Wizard File
    mStr_AppDir = mStr_DocDir & "\Gym Wizard\"
    
    If Not FSO.FolderExists(mStr_AppDir) Then
        Call Install
    End If
    
    Dim db As New UFCDB
    Set db = Nothing
    
    
    
End Sub





Function CreateBackUpFolders(ByVal Path As String) As Boolean
    Dim FSO As New FileSystemObject
    
    On Error GoTo Handler
        If Not FSO.FolderExists(Path & "\Gym Wizard\RedScreens") Then
            FSO.CreateFolder (Path & "\Gym Wizard\RedScreens")
        End If
        If Not FSO.FolderExists(Path & "\Gym Wizard\Tap Lists") Then
            FSO.CreateFolder (Path & "\Gym Wizard\Tap Lists")
        End If
        If Not FSO.FolderExists(Path & "\Gym Wizard\Schedules") Then
            FSO.CreateFolder (Path & "\Gym Wizard\Schedules")
        End If
        
    On Error GoTo 0
        CreateBackUpFolders = True
        
Handler:
    On Error GoTo 0
    CreateBackUpFolders = False
    
End Function


