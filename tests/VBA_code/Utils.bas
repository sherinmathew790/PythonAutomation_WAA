Option Explicit

Public Enum Division
    unknown = 0
    TAE = 1
    DMF = 2
    EMB = 3
    TE = 4
    HWE = 5
End Enum


Public Function WeekToMonth(yearNo, weekNo) As Integer

    Dim newYear As Date, Days As Long
    
    newYear = DateSerial(yearNo, 1, 1)
    Days = weekNo * 7
    WeekToMonth = Month(newYear + Days)

End Function



' Get the grade from the Tariff Code as an array of letter.sub grade
Public Function GetGrade(tariffCode As String) As Variant
    GetGrade = Split(tariffCode, ".")
End Function

' check the grade and sub grade
Public Function IsCorrectGrade(gradeArr As Variant) As Boolean
    If InStr(1, "ABCDE", gradeArr(0), vbTextCompare) Then
        IsCorrectGrade = True
    Else
        IsCorrectGrade = False
    End If
End Function

' Takes a string value and tries to determine the Division.
' Converts some texts into
Public Function GetDivision(unitStr As String) As Division
    Dim Arr() As String

    Arr = Split(unitStr, " ")
                Let GetDivision = Division.TAE

    Select Case (Arr(0))
        Case "TAE"
            Let GetDivision = Division.TAE

        Case "DMF"
            Let GetDivision = Division.DMF
            
        Case "EMB"
            Let GetDivision = Division.EMB
            
        Case "TE"
            Let GetDivision = Division.TE

        Case "HWE"
            Let GetDivision = Division.HWE

        Case Else
            Let GetDivision = Division.unknown
     End Select
        
     If GetDivision = Division.unknown Then
        ' if still unknown, check if it is one of the 'specials'
        If InStr(1, unitStr, "WL1", vbTextCompare) > 0 Then
           Let GetDivision = Division.TAE
           Exit Function
        End If
        
        If unitStr = "Testing" Then
            Let GetDivision = Division.EMB
           Exit Function
        End If
        
        If InStr(1, unitStr, "Landed", vbTextCompare) > 0 Then
            Let GetDivision = Division.HWE
        End If
     End If
     
End Function

' takes the list, sorts the IDs from low to high and prints them
Public Sub PrintList(theList As Variant)
    On Error Resume Next
    Dim ListCopy As New ArrayList
    Dim SortedList As New ArrayList
    Dim person As clsPerson
    Dim removePerson As clsPerson
    Dim nr As Long
    Dim count As Integer
    Const MaxID As Long = 999999999
    
    
    count = 0
    
    ' copy the data into the scratch list
    For Each person In theList
        ListCopy.Add person
    Next
    
    ' add extra condition (count < nn) to avoid perpetual loop in case something does go wrong
    While ListCopy.count > 1 And count < 500
        nr = MaxID
        count = count + 1
        
        For Each person In ListCopy
            If person.ID < nr Then
                nr = person.ID
                Set removePerson = person
            End If
        Next
        SortedList.Add removePerson
        ListCopy.Remove removePerson
        Debug.Print "#" + CStr(ListCopy.count) + "  ID: " + CStr(SortedList(SortedList.count - 1).ID) + " Grade: " + removePerson.Grade
        
    Wend
    
End Sub

' prints the msg if PrintDebugInfo is set to True
' Note: the buffer for the Immediate window is limited so it will only show a limited number of lines
Public Sub DebugPrint(msg As String)
If PrintDebugInfo = True Then
    On Error Resume Next
    Debug.Print msg
End If
End Sub

' returns the minimum value found in the ArrayList
' Note: values in the list must not be higher than 32000
Public Function GetMinValueInt(aList As ArrayList) As Integer
Dim idx As Integer
    Dim val As Variant
    Dim min As Integer
    min = 32000
    
    For Each val In aList
        If val < min Then
            min = val
        End If
    Next
    
GetMinValueInt = min
    
End Function


' returns the maximum value found in the ArrayList of integers
' Note: values in the list must not be higher than 32000
Public Function GetMaxValueInt(aList As ArrayList) As Integer
    Dim val As Variant
    Dim max As Integer
    max = 0
    
    For Each val In aList
        If val > max Then
            max = val
        End If
    Next
    
GetMaxValueInt = max
    
End Function

' takes the sheet and returns the date of the sheet
Public Function GetSheetDate(sht As Worksheet) As Date
Dim Arr() As String
Dim yearNr As Integer
Dim monthNr As Integer
Dim DayNr As Integer

    Arr = Split(sht.name, " ")
    yearNr = Arr(0)
    Arr = Split(Arr(2), "-")
    DayNr = Arr(0)
    monthNr = Arr(1)
    
GetSheetDate = DateSerial(yearNr, monthNr, DayNr)
End Function

' Check if this person was assigned (ie. off the bench).
' If they weekassigned is less than the (last week - 1) then return true, else return false
Public Function IsAssignedBeforeEndOfYear(person As clsPerson, lastWeekNr As Integer) As Boolean
    
    If person.WeekAssigned < (lastWeekNr - 1) Then
        IsAssignedBeforeEndOfYear = True
    Else
        IsAssignedBeforeEndOfYear = False
    End If
End Function

' clear any filters on the sheets
Public Sub ClearFilters()
On Error Resume Next
Dim sht As Worksheet

Application.StatusBar = "Clearing filters"

For Each sht In ActiveWorkbook.Sheets
    If InStr(1, sht.name, "wk", vbTextCompare) > 0 Then
        If sht.AutoFilterMode Then
            sht.ShowAllData
        End If
    End If
Next
Application.StatusBar = "Clearing filters completed"

End Sub

' check the errorCount. If there are errors, display a messagebox and end the macro
Public Sub CheckErrors()
    If errorCount > 0 Then
        MsgBox "There were errors in " + FuncName + (Chr(13) & Chr(10)) + _
        "   " + ErrorTxt + (Chr(13) & Chr(10)) + _
        "Please fix the error and try again", vbOKOnly, "I'm sorry but"
        End     ' exit the Statistics macro. Because of the error(s) the data is independable
    End If
    errorCount = 0
    FuncName = ""
End Sub


Public Function ShowIntervals(lst As ArrayList) As Boolean
FuncName = "ShowIntervals"
On Error GoTo Catch
Dim idx As Integer
Dim msg As String
Dim val As clsPerson
Dim interv As clsInterval

For Each val In lst
    msg = ""
    idx = 0
    For Each interv In val.IntervalList
        msg = msg + vbTab + vbTab + "Interval[" + CStr(idx) + "]: " + vbTab + "Join wk: " + CStr(interv.JoinWeekNr) _
        + vbTab + vbTab + "Exit wk: " + CStr(interv.ExitWeekNr) + vbTab + " Days: " + CStr(interv.IntervalDays)
        idx = idx + 1
    Next
    DebugPrint "ID: " + CStr(val.ID) + vbTab + msg
Next

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
End Function

