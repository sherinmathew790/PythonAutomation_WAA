Option Explicit

' The 12 monthlyPersonList ArrayLists should contain the exact same unique persons as the PersonsYearList
Function CheckUniqueMonthPersonsAgainstYearPersons(aMonthList() As ArrayList, aYearList As ArrayList) As Boolean
On Error GoTo Catch
    Dim dict As New Scripting.Dictionary
    Dim dict2 As New Scripting.Dictionary
    Dim idx As Integer
    Dim val As Variant
    
    ' get the uniqe persons over all months in the monthList
    For idx = 1 To 12
        For Each val In aMonthList(idx)
            If dict.Exists(val.ID) <> True Then
                dict.Add val.ID, val
            End If
        Next
    Next
    
    ' get the unique persons in the yearList
    For Each val In aYearList
            If dict2.Exists(val.ID) <> True Then
                dict2.Add val.ID, val
            End If
    Next
    
    ' the number of persons in dict2 should be the same as the number of persons in the yearList
    ' (because the yearList is supposed to contain only unique persons)
    If dict2.count <> aYearList.count Then GoTo Catch
    
    
    ' the number of persons in the dictionaries should be the same:
    If dict.count <> dict2.count Then GoTo Catch
    
    ' both should contain the same persons
    For Each val In aYearList
        If dict.Exists(val.ID) <> True Then GoTo Catch
    Next
    
GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
    FuncName = "CheckUniqueMonthPersonsAgainstYearPersons"
Bye:
End Function


