Attribute VB_Name = "Utilities"
Function GetSuffix(ByVal day As Integer) As String
    Select Case day
        
        Case 1, 21, 31
            GetSuffix = "st"
        Case 2, 22
            GetSuffix = "nd"
        Case 3, 3
            GetSuffix = "rd"
        Case Else:
            GetSuffix = "th"
    
    End Select
End Function
Public Function ParseJScriptResponse(ByVal sValue As String) As Variant
    
    ' * ' DIm
    Dim iStart, _
        iEnd            As Integer
    Dim Itr             As Integer
    Dim Vals            As Variant
    Dim newString As String
    
    Dim myChar As Char
    
    iStart = -1
    iEnd = -1
    
    Vals = Split(sValue, ";")
    
    newString = Right(Vals(0), Len(Vals(0)) - InStr(Vals(0), "cache") - 5)

    Vals = Split(newString, "&")
    
    ParseJScriptResponse = Vals
End Function
Public Function Encode(ByVal Value As String) As String
    
    Encode = Replace(Replace(Value, "|", "%7C"), "/", "%2F")
    
End Function

' ----------------------------------------------------------------
' Procedure Name: FileExists
' Purpose: Checks that file exists at given path
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePath (String):
' Return Type: Boolean
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Public Function FileExists(ByVal sFilePath As String) As Boolean
    
    FileExists = (Dir(sFilePath) <> "")
    
End Function

' ----------------------------------------------------------------
' Procedure Name: FolderExists
' Purpose: Check if Folder Exists
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sFilePath (String): Path
' Return Type: Boolean
' Author: mgrod
' Date: 1/13/2018
' ----------------------------------------------------------------
Public Function FolderExists(ByVal sFilePath As String) As Boolean

    Dim bResult As Boolean
    
    bResult = (Dir(sFilePath, vbDirectory) <> "")
    
    

    FolderExists = bResult

End Function

' ----------------------------------------------------------------
' Procedure Name: DeleteFile
' Purpose: Deletes file at given path
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter sFilePath (String):
' Author: Matthew Rodriguez
' Date: 10/14/2017
' ----------------------------------------------------------------
Public Sub DeleteFile(ByVal sFilePath As String)
    
    ' * ' Check file Exists
    If FileExists(sFilePath) Then
        SetAttr sFilePath, vbNormal ' Make sure file is not Read Only
        Kill sFilePath ' Delete
    End If
    
End Sub

Function IsInArray(stringToBeFound As String, Arr As Variant) As Boolean
    IsInArray = (UBound(Filter(Arr, stringToBeFound)) > -1)
End Function

Function IsInArrayList(stringToBeFound As String, Arr As Object) As Boolean
    Set Arr = CreateObject("Collection.ArrayList")
    
    For x = 0 To Arr.Count - 1
        If Arr(x) = stringToBeFound Then
            IsInArrayList = True
            Exit Function
        End If
    Next
    
    IsInArrayList = False
    
End Function

Function FormatPhoneNumber(ByVal NumberStr As String, Format As Integer) As String

    Dim FinalNumber As String
    Dim sArray() As String
    Dim sNumArray() As String
    ReDim sNumArray(0)
    
    If Format = 1 Then
        
        ReDim sArray(Len(NumberStr) - 1)
        For i = 1 To Len(NumberStr)
            sArray(i - 1) = Mid$(NumberStr, i, 1)
        Next
        
        
        ' * ' loop
        For x = 0 To UBound(sArray)
            
            Select Case sArray(x)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
                
                    sNumArray(UBound(sNumArray)) = sArray(x)
                    ReDim Preserve sNumArray(UBound(sNumArray) + 1)
                    
            End Select
        
        Next
        
        ReDim Preserve sNumArray(UBound(sNumArray) - 1)
        
        FinalNumber = "(" & sNumArray(0) & sNumArray(1) & sNumArray(2) & ")+" & sNumArray(3) & sNumArray(4) & sNumArray(5) & "-" & sNumArray(6) & sNumArray(7) & sNumArray(8) & sNumArray(9)
        
        
    End If
    
    FormatPhoneNumber = FinalNumber
    
End Function
