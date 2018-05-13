Attribute VB_Name = "Reports"
Option Explicit

Function GetReportArray(ByVal aObj_Workbook As Workbook, _
                                aStr_FileType As String) As Variant
'=======================================================================================
' Method   : GetReportArray
' Author   : Matthew Rodriguez
' Created  : 4/24/2016
' Purpose  : Retrieves data from given workbook and creates an array
' Called by:
' Arguments:
' Comments :
' Changes-------------------------------------------------------------------------------
' Date            Programmer                Change
'
'=======================================================================================
    
    ' ' Declarations
    Dim mObj_Range                            As Range
    Dim mInt_EndCol, mInt_StartRow  As Integer
    Dim mLng_EndRow                        As Long
    
    
    ' ' Set StartCol and EndCol
    mInt_StartRow = CInt(GetINI(aStr_FileType, "StartRow"))
    mInt_EndCol = CInt(GetINI(aStr_FileType, "EndCol"))
    
    
    With aObj_Workbook.Sheets(1)
    
        ' ' Loop Through Data to get Final Row
        For mLng_EndRow = mInt_StartRow To .Rows.Count
        
            If .Cells(mLng_EndRow, 2).Value = "" Then
            
                Exit For
                
            End If
            
        Next
        
        
        ' ' Create Range and Array
        Set mObj_Range = .Range(.Cells(mInt_StartRow, 1), .Cells(mLng_EndRow - 1, mInt_EndCol))
        
    End With
    
    ' ' Return
    GetReportArray = mObj_Range.Value
    
    
End Function

Function ReportSave(dDate As Date, sType As String) As Workbook

    ' ' Declarations
    Dim FSO             As Object
    Dim book, WB        As Workbook
    Dim TarSheet        As Worksheet
    Dim tempPath, _
            sPath, _
            sFileName, _
            sDir, _
            mName           As String
            
        
    Dim sourceRow, _
        sourceStartRow, _
        sourceEndRow As Integer
    
    'Set Variables
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    tempPath = ThisWorkbook.Path & "\Templates\"
    mName = MonthName(Month(dDate))
    
    
    'Get Book Type
    Select Case sType
               
        'Check Ins
        Case "RedScreens"
        
            sFileName = Month(dDate) & "." & day(dDate) & "RedScreens.xlsx"
            sDir = ThisWorkbook.Path & "\Reports\RedScreens\" & Year(dDate) & "\" & mName & "\"
            sPath = sDir & sFileName
            
            
            Call FSO.CopyFile(tempPath & "RedScreenTemp.xlsx", _
                                sPath)
              
        Case "TapList"
        
            sFileName = Month(dDate) & "." & day(dDate) & "TapListTemp.xlsx"
            sDir = ThisWorkbook.Path & "\Reports\TapList\" & Year(dDate) & "\" & MonthName(Month(dDate)) & "\"
            sPath = sDir & sFileName
            
            For Each WB In Application.Workbooks
                If WB.FullName = sFileName Then
                    WB.Close
                End If
            Next
                        
            Call FSO.CopyFile(tempPath & "TapListTemp.xlsx", _
                                sPath)
            
    End Select
    
    Set book = Workbooks.Open(sPath)
    Set ReportSave = book
    
End Function

