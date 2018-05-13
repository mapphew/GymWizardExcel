VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pClassName As String
Private pStartTime As Date
Private pEndTime As Date
Private pDuration As Date
Private pCoach As String
Private pWeekDay As String
Private pClassType As String

Public Property Get ClassType() As String
    ClassType = pClassType
End Property

Public Property Get Name() As String
    Name = pClassName
End Property

Public Property Get StartTime() As Date
    StartTime = pStartTime
End Property

Public Property Get EndTime() As Date
    EndTime = pEndTime
End Property

Public Property Get Duration() As Date
    Duration = pDuration
End Property

Public Property Get coach() As String
    coach = pCoach
End Property

Public Property Get Weekday() As String
    Weekday = pWeekDay
End Property

Property Let Weekday(ByVal Value As String)
    pWeekDay = Value
End Property


Public Sub CreateClass(ByVal title As String, ByVal StartTime As Date, ByVal coach As String, ByVal ClassType As String)

    ' ' Set Vars
    pClassName = title
    pStartTime = StartTime
    pEndTime = EndTime
    pCoach = coach
    pClassType = ClassType
    
End Sub

