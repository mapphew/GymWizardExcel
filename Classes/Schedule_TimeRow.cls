VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule_TimeRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pExcelArray                 As Variant
Private pTimeArray                  As Variant
Private pMaximumClassesForTimeSlot  As Integer
Private pNotEmpty                   As Boolean
Private pWeekClassCount(1 To 7)     As Integer
Public Property Get NotEmpty() As Boolean
    NotEmpty = pNotEmpty
End Property
Public Property Get ExcelArray() As Variant
    ExcelArray = pTimeArray
 End Property
Private Sub SetMaximumClassCount()

    ' * ' Dim Variant
    Dim mArray(), _
        mSubArray()   As Variant
    Dim Itr         As Integer
    
    ' * ' Set Array
    mArray = pTimeArray
    
    ' * ' Iterate Through Values in Array
    For Itr = 1 To UBound(mArray)
        mSubArray = mArray(Itr)
        
        If Not IsEmpty(mSubArray(0)) Then
            pNotEmpty = True
            Exit Sub
        End If
        
    Next
    
    ' * ' Sets pNotEmpty to False As Default
    pNotEmpty = False
    
End Sub
Public Function ConvertToValidExcelArray() As Variant()

    ' * ' DIM Vars
    Dim mArray(), _
        mNewArray(), _
        mSubArray()     As Variant
    
    Dim mCoach, _
        mTitle          As String
    Dim Itr1, _
        Itr2, _
        DayItr, _
        ClassItr1, _
        ClassItr2, _
        Count, _
        Max, _
        MaxRows, _
        mRowsPerTitle, _
        mRowsPerCoach, _
        mRowsPerClass   As Integer
    
    Dim mClass          As Schedule_Class
    ' * ' Set Instance of Array
    mArray = pTimeArray
    
    ' * ' Get Week Class Count
    For Itr1 = 1 To 7
        
        If IsEmpty(mArray(Itr1)(0)) Then
            pWeekClassCount(Itr1) = 0
        Else:
            pWeekClassCount(Itr1) = UBound(mArray(Itr1)) + 1
        End If
    
    Next
    
    ' * ' Max Classes FOr That Week
    pMaximumClassesForTimeSlot = Application.WorksheetFunction.Max(pWeekClassCount)
    Max = pMaximumClassesForTimeSlot
    MaxRows = Max * 2
    
    ' * ' Dim NewArray Based On # of Classes
    ' * ' New Array Will be in an EXCEl friendly Format
    ReDim Preserve mNewArray(0 To pMaximumClassesForTimeSlot * 2 - 1, 0 To 7)
    
           
        
    For Itr1 = 1 To 7
        
        Count = pWeekClassCount(Itr1)
        
        ' * ' Check If ClassCount For That Day is 0
        If Count > 0 Then
            
            ' * ' Get Rows Per Title/Coach
            If Count = Max Then
                mRowsPerTitle = 1
                mRowsPerCoach = 1
                
            ElseIf Count < Max And MaxRows Mod Count = 0 Then
                mRowsPerTitle = Max / Count
                mRowsPerCoach = Max / Count
                
            Else:
                mRowsPerTitle = 1
                mRowsPerCoach = (MaxRows / Count) - mRowsPerTitle
            End If
            
            ' * ' Total Rows Per Class
            mRowsPerClass = mRowsPerTitle + mRowsPerCoach
            
            
            ' * ' Now Populate Excel Array Using new PerRow Vals
            For Itr2 = 0 To MaxRows - 1 Step mRowsPerClass
                                    
                    ' * ' Get Class Title NAd Coach Name
                    mTitle = mArray(Itr1)(Itr2 / mRowsPerClass)(0)
                    mCoach = mArray(Itr1)(Itr2 / mRowsPerClass)(1)
                    
                    ' * ' Add To New Array At given index
                    mNewArray(Itr2, Itr1) = mTitle
                    mNewArray(Itr2 + mRowsPerTitle, Itr1) = mCoach
                                    
                
            Next
                
                    
        End If
            
    Next
        
    ' * ' Now Adjust Rows To Accomodate Merged Cells
    mNewArray(0, 0) = Format(mArray(0), "Medium Time")
    
    ConvertToValidExcelArray = mNewArray
End Function

Private Function GetClassTitle(ByVal Value As String) As String

    ' ' Create Title
    Dim mTitle As String
    
    Dim sTitle As String
    sTitle = Replace(Value, " ", "")
    
    Select Case sTitle
        
        
        '' Ultimate = Array("DailyUltimateTraining™", "TRX", "Striking", "BrazilianJiu-Jitsu(Gi)Level1", "BoxingConditioning", "FightFit", "BrazilianJiu-Jitsu(Gi)Level2(90)", _
                        "BrazilianJiu-Jitsu(NoGi)", "BrazilianJiu-Jitsu(NoGi)", "BoxingSkills&Techniques", "Women'sBJJ(No-Gi)", "Women'sBrazilianJiu-Jitsu(Gi)", "MMA", "Judo", "FamilyBrazilianJiu-Jitsu(Gi)", _
                        "KickboxingConditioning", "BJJ", "BrazilianJiu-Jitsu(Gi)Level3(90)", "Women'sonlyBJJ", "Women'sBJJ(No-Gi)", "FamilyDailyUltimateTraining™", "Women'sBrazilianJiu-Jitsu(Gi)", "Women'sBJJ(No-Gi)", "Women'sBrazilianJiu-Jitsu(Gi)")
                        
                        
        '' Fitness = Array("RestorativeYoga", "HighOctaneConditioning", "BODYPUMP™", "HOTFUSION®", "UFCFIT", "MATRX", "KillerCore(30)", "YogaFlow", _
                                "KillerCycle", "Zumba®", "Hi-DEFinition", "MuscleMakeover", "PiYo®", "KickAssKick", "POUND:Rockout.Workout.™", "P90X®Live!", _
                                "CircuitBreaker", "Insanity®Live!", "Barre-lessBarre", "CardioFunkDanceParty", "HardCoreAbs(30)", "BelowtheBelt")
        
        Case "HighOctaneConditioning", "HighOctaneConditio": mTitle = "High Octane Conditioning"
        Case "BODYPUMP™", "HOTFUSION®", "UFCFIT", "MATRX", "Zumba®", "Hi-DEFinition", "PiYo®": mTitle = Value
        Case "KillerCore(30)": mTitle = "Killer Core (30min)"
        Case "YogaFlow": mTitle = "Yoga Flow"
        Case "KillerCycle": mTitle = "Killer Cycle"
        Case "MuscleMakeover": mTitle = "Muscle Makeover"
        Case "POUND:Rockout.Workout.™", "POUND:Rockout.Work": mTitle = "POUND: Rockout Workout™"
        Case "P90X®Live!": mTitle = "P90X® Live!"
        Case "CircuitBreaker": mTitle = "Circuit Breaker"
        Case "Insanity®Live!": mTitle = "Insanity® Live!"
        Case "Barre-lessBarre": mTitle = "Barre-less Barre"
        Case "CardioFunkDanceParty", "CardioFunkDancePa": mTitle = "Cardio Funk Dance Party"
        Case "HardCoreAbs(30)": mTitle = "Hard Core Abs (30min)"
        Case "BelowtheBelt": mTitle = "Below the Belt"
        Case "KickAssKick": mTitle = "KickAss Kick"
        Case "EnergizingYoga": mTitle = "Energizing Yoga"
        Case "BelowTheBelt(30)": mTitle = "Below The Belt"
        
        Case "FightFitFightFitSampler-BTC": mTitle = "[A] Fight Fit"
        Case "BrazilianJiu-Jitsu(Gi)Level1", "BrazilianJiu-Jitsu": mTitle = "[C] [1] BJJ"
        Case "DailyUltimateTraining™", "DailyUltimateTrain": mTitle = "[B] DUT"
        Case "TRX": mTitle = "[A] TRX"
        Case "Striking":
            mTitle = "[A] [1] Striking"
        Case "FamilyDailyUltimat": mTitle = "[B] Family DUT"
        Case "BoxingConditioning": mTitle = "[A] Boxing Conditioning"
        Case "FightFit": mTitle = "[A] MMA Fight Fit"
        Case "BrazilianJiu-Jitsu(Gi)Level2(90)": mTitle = "[C] [2] BJJ (90min)"
        Case "BrazilianJiu-Jitsu(NoGi)": mTitle = "[C] [1 & 2] BJJ (No-Gi)"
        Case "BoxingSkills&Techniques", "BoxingSkills&Tech": mTitle = "[A] Boxing Style & Technique"
        Case "Women'sBJJ(No-Gi)", "Women'sBrazilianJiWomen'sBJJ(No-Gi)": mTitle = "[C] Women's BJJ (No-Gi)"
        Case "Women'sBrazilianJiu-Jitsu(Gi)", "Women'sBrazilianJiWomen'sonlyBJJ": mTitle = "[C] Women's BJJ (Gi)"
        Case "MMA": mTitle = "[O] MMA"
        Case "Judo": mTitle = "[O] Judo"
        Case "FamilyBrazilianJiu-Jitsu(Gi)", "FamilyBrazilianJiu": mTitle = "Family BJJ"
        Case "KickboxingConditioning", "KickboxingCondition": mTitle = "[A] Kickboxing Conditioning"
        
        Case "Youth Wrestling[YouthWrestling]": mTitle = "Youth Wrestling"
        Case "YouthStriking": mTitle = "Youth Striking"
        Case "YouthBrazilianJiu-": mTitle = "Youth BJJ"
        Case "YouthBoxing": mTitle = "Youth Boxing"
        Case "YouthBJJ(NoGI)": mTitle = "Youth BJJ (No - GI)"
        Case "YouthDailyUltimateSAQ": mTitle = "S.A.Q"
        Case "Zumba® for KIDS": mTitle = "Youth Zumba"
        
        
        
        Case Else:
            
            If InStr(Value, "KillerCycle") Then
                mTitle = "Killer Cycle"
            ElseIf InStr(Value, "SPECIAL") Then
                mTitle = "Special Class"
            ElseIf Value = "GROOV3" Then
                mTitle = "GROOV3"
            ElseIf InStr(Value, "Zumba") Then
                mTitle = "Zumba"
            ElseIf InStr(Value, "PeopleWaiting") Then
                Value = Replace(Value, "PeopleWaiting", "")
                mTitle = GetClassTitle(Mid(Value, 2))
            Else:
                Debug.Print ("The Following ClassString Does Not Have a Case: " & sTitle)
            End If
            
            
    End Select
    
    GetClassTitle = mTitle
End Function
Public Sub Class_InitializeWithValues(ByVal TimeRow As Variant)

    ' * ' Set Row
    pTimeArray = TimeRow
     
    ' * ' Set MaxClassCount
    SetMaximumClassCount
     
    
End Sub
