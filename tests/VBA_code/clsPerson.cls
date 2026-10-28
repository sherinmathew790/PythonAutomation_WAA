VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The class Person contains the data for a single person as found in each of the rows in the weekly sheets.
' The name of the Person is not included, but the ID is.
Option Explicit

Private p_ID As Long
Private p_days As Integer
Private p_city As String
Private p_unit As Division
Private p_grade As String
Private p_subGrade As String
Private p_since As Date
Private p_year As Integer
Private p_month As Integer
Private p_week As Integer
Private p_day As Integer
Private p_weekAssigned As Integer
Private p_monthAssigned As Integer
Private p_Identity As Long

Private p_intervalList As ArrayList
Private p_currentInterval As clsInterval
Private p_intervalCount As Integer

Private Sub Class_Initialize()
    p_ID = 0
    p_city = Empty
    p_unit = unknown
    p_grade = Empty
    p_subGrade = Empty
    p_since = Empty
    Set p_intervalList = New ArrayList
End Sub

Property Get Identity() As Long
    Identity = p_Identity
End Property
Property Let Identity(value As Long)
    p_Identity = value
End Property

Property Get ID() As Long
    ID = p_ID
End Property

Property Let ID(value As Long)
    p_ID = value
End Property

Property Get Days() As Integer
    Days = p_days
End Property
Property Let Days(value As Integer)
    p_days = value
End Property

Property Get City() As String
    City = p_city
End Property
Property Let City(value As String)
    p_city = value
End Property

Property Get Unit() As Division
    Unit = p_unit
End Property

Property Let Unit(value As Division)
    p_unit = value
End Property

Property Get Grade() As String
    Grade = p_grade
End Property

Property Let Grade(value As String)
    p_grade = value
End Property

Property Get SubGrade() As String
    SubGrade = p_subGrade
End Property

Property Let SubGrade(value As String)
    p_subGrade = value
End Property

Property Get Since() As Date
    Since = p_since
End Property

Property Let Since(value As Date)
    p_since = value
End Property

Property Get yearNr() As Integer
    yearNr = p_year
End Property

Property Let yearNr(value As Integer)
    p_year = value
End Property

Property Get monthNr() As Integer
    monthNr = p_month
End Property

Property Let monthNr(value As Integer)
    p_month = value
End Property

Property Get weekNr() As Integer
    weekNr = p_week
End Property

Property Let weekNr(value As Integer)
    p_week = value
End Property

Property Get DayNr() As Integer
    DayNr = p_day
End Property

Property Let DayNr(value As Integer)
    p_day = value
End Property

Property Get WeekAssigned() As Integer
    WeekAssigned = p_weekAssigned
End Property

Property Let WeekAssigned(value As Integer)
    p_weekAssigned = value
End Property

Property Get MonthAssigned() As Integer
    MonthAssigned = p_monthAssigned
End Property

Property Let MonthAssigned(value As Integer)
    p_monthAssigned = value
End Property

Property Get IntervalList() As ArrayList
    Set IntervalList = p_intervalList
End Property

Property Set IntervalList(value As ArrayList)
    Set p_intervalList = value
End Property

Property Get CurrentInterval() As clsInterval
    Set CurrentInterval = p_currentInterval
End Property

Property Set CurrentInterval(value As clsInterval)
    Set p_currentInterval = value
End Property


' The person entered the bench. Store the weekNr, MonthNr, Since and Days of the entry in a new Interval
Function StoreNewBenchJoin(newSince As Date, sheetDate As Date, newJoinWeekNr As Integer, newJoinMonthNr As Integer) As Boolean
    FuncName = "StoreNewBenchJoin"
    On Error GoTo Catch
    
    Dim newInterval As clsInterval
    
    Set newInterval = New clsInterval
    If newInterval.StoreNewBenchJoin(newJoinWeekNr, newJoinMonthNr, Year(newSince)) <> True Then GoTo Catch
    p_intervalList.Add newInterval
    Set p_currentInterval = newInterval
    
    Since = newSince
    Days = DateDiff("d", Since, sheetDate)

    StoreNewBenchJoin = True
    GoTo Bye
Catch:
    StoreNewBenchJoin = False
    
Bye:
End Function

' Add the exit week, month and Days in the current interval and update the Since for this person
Function StoreNewBenchExit(newSince As Date, sheetDate As Date) As Boolean
    FuncName = "StorePreviousBenchInterval"
    On Error GoTo Catch
    Dim test As String
            
    If p_currentInterval Is Nothing Then
        StoreNewBenchExit = True
        GoTo Bye
    End If
    ' store the data of the interval that is now closed (= has both Join and Exit data)
    If p_currentInterval.StoreNewBenchExit(WeekAssigned, MonthAssigned, Days) <> True Then GoTo Catch
   
    StoreNewBenchExit = True
    GoTo Bye
Catch:
    StoreNewBenchExit = False
    
Bye:
End Function


' when a person is assigned they disappear from the weekly sheets. As a result the last interval has no
' data for an exit (which is only added when they appear again on the bench). Close the interval by adding
' an exit during the last sheet.
Public Function CloseLastInterval() As Boolean
    FuncName = "CloseLastInterval"
    On Error GoTo Catch
    
    If p_currentInterval Is Nothing Then
        CloseLastInterval = True
        GoTo Bye
    End If
   If (CurrentInterval.StoreNewBenchExit(weekNr, monthNr, Days)) <> True Then
        GoTo Catch
    End If
    
    CloseLastInterval = True
    GoTo Bye
Catch:
    CloseLastInterval = False

Bye:
End Function


' copy the dictionaries for the intervals to the target
Function CopyIntervalsTo(ByRef targetPerson As Variant) As Boolean
    FuncName = "CopyIntervalsTo"
    On Error GoTo Catch
    
    Set targetPerson.CurrentInterval = Me.CurrentInterval
    Set targetPerson.IntervalList = Me.IntervalList.Clone
    
    CopyIntervalsTo = True
    
    GoTo Bye
Catch:
    CopyIntervalsTo = False
    
Bye:
End Function


' Sums the days in all the intervals together with the days in .Days.
' Sets .Days to this sum and return the sum
Function DaysTotal() As Integer
    Dim InterDays As Integer
    Dim inter As clsInterval
    
    For Each inter In IntervalList
        InterDays = InterDays + inter.IntervalDays
    Next
    
    DaysTotal = InterDays
End Function

