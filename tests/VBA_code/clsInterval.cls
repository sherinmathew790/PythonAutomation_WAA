VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The class Interval contains a period that a person was on the bench.
' It has timestamps for joining and exiting the bench.
' When processing is completed, each interval will be closed. That is to say
' it has a valid timestamp for when the period started (Join) and a
' valid timestamp for when the period ended (Exit), as well as the duration
' for the period.
Option Explicit

Private t_joinWeek As Integer
Private t_joinMonth As Integer
Private t_joinYear As Integer
Private t_exitWeek As Integer
Private t_exitMonth As Integer
Private t_days As Integer

Public Property Get JoinWeekNr()
    JoinWeekNr = t_joinWeek
End Property

Public Property Get JoinMonthNr()
    JoinMonthNr = t_joinMonth
End Property

Public Property Get JoinYear()
    JoinYear = t_joinYear
End Property

Public Property Get ExitWeekNr()
    ExitWeekNr = t_exitWeek
End Property

Public Property Get ExitMonthNr()
    ExitMonthNr = t_exitMonth
End Property

Public Property Get IntervalDays() As Integer
    IntervalDays = t_days
End Property


Function StoreNewBenchJoin(newJoinWeekNr As Integer, newJoinMonthNr As Integer, newYear As Integer) As Boolean
    FuncName = "clsInterval.StoreNewBenchJoin"
    On Error GoTo Catch
        
    t_joinWeek = newJoinWeekNr
    t_joinMonth = newJoinMonthNr
    t_joinYear = newYear
    
    StoreNewBenchJoin = True
    
    GoTo Bye
Catch:
    StoreNewBenchJoin = False
    
Bye:
End Function


Function StoreNewBenchExit(newExitWeekNr As Integer, newExitMonthNr As Integer, duration As Integer) As Boolean
    FuncName = "clsInterval.StoreNewBenchExit"
    On Error GoTo Catch
            
    t_exitWeek = newExitWeekNr
    t_exitMonth = newExitMonthNr
    t_days = duration
    
    StoreNewBenchExit = True
    GoTo Bye
Catch:
    StoreNewBenchExit = False
    
Bye:
End Function

