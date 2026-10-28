Option Explicit

' The version number is displayed for reference in case updates are done to the code
' UAT_TEST Declared golbal variable
Public GlobalPersonsYearList As ArrayList
Public Const Version As String = "v1.7"
Public errorCount As Integer
Public ReviewYear As String
Public Const PrintDebugInfo As Boolean = True  ' if true, debug info will be printed to the Immediate window
Public DummyDate As Date
Public ErrorTxt As String
Public FuncName As String

' latestWeek is the weeknr of the latest available sheet. Full year = 52, but we might also be analysing
' sometime during the year.
Public latestweek As Integer
' WeeksRemoved is the minimum number of weeks a person has to be absent from consecutive sheets to be deemed Removed.
' In other words, if the person has been absent for more than WeeksRemoved weeks, they are assumed to be on assignment
Public Const WeeksRemoved As Integer = 2

'Workforce:
Public DMF_Nr As Integer
Public EMB_Nr As Integer
Public HWE_Nr As Integer
Public TAE_Nr As Integer
Public TE_Nr As Integer

' This sub is the main entry point of the VBA application
Sub Statistics_Main()

    Dim idx As Integer
    Dim chartObj As ChartObject
    Dim sht As Worksheet
    Dim chartSheet As Worksheet
    Dim Arr() As String
    Dim yearNr As Integer
    Dim monthNr As Integer
    Dim weekNr As Integer
    Dim DayNr As Integer
    Dim totalErrors As Integer
    Dim b As Variant
    Dim wkList As Variant
    Dim found As Boolean
    Dim allSheets(275) As Worksheet
    Dim sheetCount As Integer
    
    ' Available persons
    Dim WeeklySheetList As New ArrayList             ' list with the latest sheet for each week; the starting point for all following calculations
    Dim PersonListPerWeek As New ArrayList           ' list with available persons in a week (based on latest sheet for each week)
    Dim AllPersonWeeksList As New ArrayList          ' list with all lists per week with available persons
    Dim PersonMonthsList(1 To 12) As New ArrayList   ' list of available persons for each month
    Dim personsYearList As New ArrayList             ' list with all unique available persons per year
    Dim BenchExitsYearList As New ArrayList          ' list with all persons that exited the bench (possibly multiple times) in a year
    
    errorCount = 0
    DummyDate = DateSerial(2000, 1, 1)
    Set chartSheet = Sheets("Statistics")
    chartSheet.Activate
    ' write version number to sheet
    chartSheet.Cells.Item(9, "Q").value = Version
    ' get the number of employees for each Division
    ReadWorkforce
    
    ' remove old graphs
    For Each chartObj In chartSheet.ChartObjects
        chartObj.Delete
    Next
    
    ' Check for sheets of the chosen year
    ' determine which year we are looking at
    ReviewYear = Sheets("Statistics").Cells.Item(4, "G").value
    ' look for the sheets
    idx = 0
    found = False
    For Each sht In ActiveWorkbook.Sheets
        Arr = Split(sht.Name, " ")
        If InStr(1, Arr(0), ReviewYear) > 0 Then
            Set allSheets(idx) = sht
            idx = idx + 1
            found = True
        End If
    Next
    If found <> True Then
        MsgBox "There are no sheets for " + ReviewYear + ".", vbOKOnly, "I'm sorry but"
        End     ' exit this macro
    End If
    sheetCount = idx - 1

    DebugPrint vbCrLf + "------- NEW NEW NEW -------"
    DebugPrint "Version: " + Version
    DebugPrint "Assignment threshold: " + CStr(WeeksRemoved) + " weeks"
    ' init arraylist for each month
    For idx = 1 To 12
        Set PersonMonthsList(idx) = New ArrayList
    Next

   ' get the list with the latest sheet for each week, for all weeks
    Set WeeklySheetList = GetLastSheetPerWeek(allSheets, sheetCount)
    CheckErrors
    
    ' assemble a List with the persons for each week
    For Each sht In WeeklySheetList
        If InStr(1, sht.Name, ReviewYear) > 0 And Not InStr(1, sht.Name, "backup", vbTextCompare) > 0 Then
            ' get year, month,week and day from sheet name
            Arr = Split(sht.Name, " ")
            yearNr = Arr(0)
            weekNr = Split(Arr(1), "k")(1)
            Arr = Split(Arr(2), "-")
            DayNr = Arr(0)
            monthNr = Arr(1)
            
            Set PersonListPerWeek = BuildPersonListPerWeek(sht, yearNr, monthNr, weekNr, DayNr)
            AllPersonWeeksList.Add PersonListPerWeek
            Set PersonMonthsList(monthNr) = CreateMonthlyList(PersonMonthsList, monthNr, PersonListPerWeek)
'            DebugPrint "processed Sheet: " + sht.name + ", count: " + CStr(ListPerWeek.count)
        End If
    Next
    
    CheckErrors
    latestweek = weekNr
    Application.StatusBar = "Creating year lists"
    DebugPrint "total weeks: " + CStr(AllPersonWeeksList.count)
    Set personsYearList = CreatePersonsYearList(AllPersonWeeksList)
    CheckErrors
    DebugPrint "Errors in PersonsYearList: " + CStr(errorCount) + " of " + CStr(personsYearList.count) + " persons"
    
    Set BenchExitsYearList = CreateBenchExitYearList(personsYearList)
    CheckErrors
    DebugPrint "Errors in BenchExitsYearList: " + CStr(errorCount) + " of " + CStr(BenchExitsYearList.count) + " persons"
    
    ' Sanity checks
    Application.StatusBar = "Running sanity checks"
    If CheckUniqueMonthPersonsAgainstYearPersons(PersonMonthsList, personsYearList) > 0 Then CheckErrors
'    If ShowIntervals(personsYearList) > 0 Then CheckErrors
'    If ShowIntervals(BenchExitsYearList) > 0 Then CheckErrors
    
    ' show graphs
    Application.StatusBar = "Creating graphs"
    
    ' UAT_TEST Store the personsYearList in the global variable
    Set GlobalPersonsYearList = personsYearList

    ' the following graphs use the number of persons
    If ShowYearByDivision(chartSheet, personsYearList, 100, 150) > 0 Then CheckErrors
    If ShowYearByGrade(chartSheet, personsYearList, 600, 150) > 0 Then CheckErrors
    
    If ShowMonthlyTotalByDivision(chartSheet, PersonMonthsList, 100, 400) > 0 Then CheckErrors
    If ShowMonthlyTotalByGrade(chartSheet, PersonMonthsList, 600, 400) > 0 Then CheckErrors
    
    If ShowTotalPercentageByDivision(chartSheet, personsYearList, 100, 650) > 0 Then CheckErrors
    If ShowMonthlyPercentageByDivision(chartSheet, PersonMonthsList, 600, 650) > 0 Then CheckErrors
   
    If ShowTotalDaysByDivision(chartSheet, personsYearList, 100, 900) > 0 Then CheckErrors
    If ShowTotalDaysByGrade(chartSheet, personsYearList, 600, 900) > 0 Then CheckErrors
   
    ' the following graphs use the bench intervals inside the analysis period
    If ShowYearAvgByDivision(chartSheet, BenchExitsYearList, 100, 1200) > 0 Then CheckErrors
    If ShowYearAvgByGrade(chartSheet, BenchExitsYearList, 600, 1200) > 0 Then CheckErrors
    
    If ShowYearAvgBySubGrade(chartSheet, BenchExitsYearList, "A", 100, 1450) > 0 Then CheckErrors
    If ShowYearAvgBySubGrade(chartSheet, BenchExitsYearList, "B", 600, 1450) > 0 Then CheckErrors
    If ShowYearAvgBySubGrade(chartSheet, BenchExitsYearList, "C", 100, 1700) > 0 Then CheckErrors
    If ShowYearAvgBySubGrade(chartSheet, BenchExitsYearList, "D", 600, 1700) > 0 Then CheckErrors
    If ShowYearAvgBySubGrade(chartSheet, BenchExitsYearList, "E", 100, 1950) > 0 Then CheckErrors
    
    If ShowDurationByWeeks(chartSheet, BenchExitsYearList, 100, 2200) > 0 Then CheckErrors
    If ShowDurationByWeeksByGrade(chartSheet, BenchExitsYearList, 100, 2450) > 0 Then CheckErrors
    
    If ShowMonthlyBenchExitsByDivision(chartSheet, BenchExitsYearList, 100, 2700) > 0 Then CheckErrors
    If ShowMonthlyBenchExitsByGrade(chartSheet, BenchExitsYearList, 600, 2700) > 0 Then CheckErrors
    
    If ShowMonthlyBenchEntriesByDivision(chartSheet, personsYearList, 100, 2950) > 0 Then CheckErrors
    If ShowMonthlyBenchEntriesByGrade(chartSheet, personsYearList, 600, 2950) > 0 Then CheckErrors
    
    If ShowMonthlyRelativeExitsByDivision(chartSheet, BenchExitsYearList, PersonMonthsList, 100, 3200) > 0 Then CheckErrors
    If ShowMonthlyRelativeExitsByGrade(chartSheet, BenchExitsYearList, PersonMonthsList, 600, 3200) > 0 Then CheckErrors
    
    
    If ShowMonthlyRatioByDivision(chartSheet, personsYearList, BenchExitsYearList, 100, 3450) > 0 Then CheckErrors
    If ShowMonthlyRatioByGrade(chartSheet, personsYearList, BenchExitsYearList, 600, 3450) > 0 Then CheckErrors
   
    If ShowMonthlyDeltaByDivision(chartSheet, personsYearList, BenchExitsYearList, 100, 3800) > 0 Then CheckErrors
   
    
Application.StatusBar = "Graphs Completed"
    
End Sub
'UAT_TEST get division counts for Unit
Public Function CalculateDivisionCounts(persYearList As ArrayList) As Dictionary
    Dim YTAE As Integer
    Dim YDMF As Integer
    Dim YEMB As Integer
    Dim YTE As Integer
    Dim YHWE As Integer
    Dim unknown As Integer
    Dim person As clsPerson
    Dim result As Dictionary

    ' Initialize variables
    YTAE = 0
    YDMF = 0
    YEMB = 0
    YTE = 0
    YHWE = 0
    unknown = 0

    ' Create a dictionary to store the results
    Set result = New Dictionary

    For Each person In persYearList
        Select Case person.Unit
            Case division.TAE
                YTAE = YTAE + 1
            Case division.DMF
                YDMF = YDMF + 1
            Case division.EMB
                YEMB = YEMB + 1
            Case division.TE
                YTE = YTE + 1
            Case division.HWE
                YHWE = YHWE + 1
            Case Else
                ' unknown unit
                unknown = unknown + 1
        End Select
    Next

    ' Add the values to the dictionary
    result.Add "YTAE", YTAE
    result.Add "YDMF", YDMF
    result.Add "YEMB", YEMB
    result.Add "YTE", YTE
    result.Add "YHWE", YHWE
    result.Add "unknown", unknown

    ' Return the dictionary
    Set CalculateDivisionCounts = result
End Function
' UAT_TEST Access the global variable GlobalPersonsYearList
Function GetGlobalPersonsYearList() As ArrayList
    Set GetGlobalPersonsYearList = GlobalPersonsYearList
End Function
' Go through the list of sheets and add the last one of every week to the ListOfWeeklySheets
' Note: the number of sheets for any given week is not fixed, f.e. some weeks have 1 sheet, other weeks have 3
' returns ListOfWeeklySheets
Private Function GetLastSheetPerWeek(allSheets() As Worksheet, sheetCount As Integer) As ArrayList
    FuncName = "GetLastSheetPerWeek"
    On Error GoTo Catch
    
    Dim idx As Integer
    Dim var As Variant
    Dim sht As Worksheet
    Dim shtToSave As Worksheet
    Dim Arr() As String
    Dim dayArr() As String
    Dim yearNr As Integer
    Dim monthNr As Integer
    Dim weekNr As Integer, oldWeekNr As Integer
    Dim DayNr As Integer, oldDayNr As Integer
    Dim ListOfWeeklySheets As New ArrayList

'    DebugPrint "Creating last sheet of every week"
    
    For idx = 0 To sheetCount
        Set sht = allSheets(idx)
        If InStr(1, sht.Name, ReviewYear) > 0 And Not InStr(1, sht.Name, "backup") > 0 Then
            ' get year, month and week from sheet name
            Arr = Split(sht.Name, " ")
            yearNr = Arr(0)
            weekNr = Split(Arr(1), "k")(1)
            monthNr = WeekToMonth(yearNr, weekNr)
            
            ' setup for first sheet
            If oldWeekNr = 0 Then
                oldWeekNr = weekNr
            End If
            
            ' when switching to a new week, add the previous one to the list
            If weekNr > oldWeekNr And (Not shtToSave Is Nothing) Then
                ListOfWeeklySheets.Add shtToSave
                oldWeekNr = weekNr
            End If
            Set shtToSave = sht
        End If
    Next
      
    ' and add the last sheet
    ListOfWeeklySheets.Add shtToSave
    GoTo Bye
    
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
    
Bye:
    Set GetLastSheetPerWeek = ListOfWeeklySheets
   
End Function

' Returns a list with the available persons for the given week
' NOTE: if the value in the column Days is negative, the entry is ignored.
'       the negative numbers indicate when the availability supposedly starts, so they
'       are not available yet
Function BuildPersonListPerWeek(sht As Worksheet, yearNr As Integer, monthNr As Integer, weekNr As Integer, DayNr As Integer) As ArrayList
FuncName = "BuildPersonListPerWeek"
On Error GoTo Catch

Dim person As clsPerson
Dim rowNr As Variant
Dim idx As Integer, beginning As Integer
Dim val As String
Dim dateFormat As String
Dim Days As Integer
Dim ListOfPersons As New ArrayList
Dim Grade As Variant
Dim div As division
Dim identCounter As Integer
    
identCounter = 0
' find beginning of table
For beginning = 1 To 100
    If sht.Cells.Item(beginning, 1).value = "1" Then Exit For
Next

' just assume we found the header of the table
For idx = beginning To 20000
    If sht.Cells.Item(idx, 2).value = "Removed from planboard" Or sht.Cells.Item(idx, 2).value = Empty Then
         ' we have reached the end of the rows with available persons, so we are done here
         Exit For
    End If
     
    Days = sht.Cells.Item(idx, 3).value
    If Days >= 0 Then
        div = GetDivision(sht.Cells.Item(idx, 8).value)
        Grade = GetGrade(sht.Cells.Item(idx, 10).value)
        ' Check if the grade,sub grade and Division fall within the acceptable range
        If IsCorrectGrade(Grade) And div <> division.unknown Then
            ' create and add a new person to the dictionary
            Set person = New clsPerson
            person.Identity = weekNr * 100 + identCounter
            identCounter = identCounter + 1
            person.yearNr = yearNr
            person.monthNr = monthNr
            person.weekNr = weekNr
            person.DayNr = DayNr
            person.Since = DummyDate
            person.Days = Days
            person.ID = sht.Cells.Item(idx, 4).value
            person.City = sht.Cells.Item(idx, 7).value
            Let person.Unit = div
            person.Grade = Grade(0)
            If Grade(1) > 5 Then Grade(1) = 5
            person.SubGrade = Grade(1)
            ListOfPersons.Add person
        End If
   End If
        
    
GoTo Bye
Catch:
    ' increment the counter but continue
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
    DebugPrint "Error in " + sht.Name + " at number " + CStr(idx - beginning + 1)
    
Bye:
Next

' DebugPrint "Created list of persons for week " + CStr(WeekNr) + ", count = " + CStr(ListOfPersons.count)
 Application.StatusBar = "Created list of persons for week " + CStr(weekNr) + ", count = " + CStr(ListOfPersons.count)
  
 Set BuildPersonListPerWeek = ListOfPersons

End Function

' create a list of unique persons for the given month
Function CreateMonthlyList(MonthsList As Variant, monthNr As Integer, ListPerWeek As ArrayList) As ArrayList

Dim m As Variant
Dim weekPerson As clsPerson
Dim MonthPerson As clsPerson
Dim found As Boolean
Dim idx As Integer

'DebugPrint "Creating list of unique persons per month"

For Each weekPerson In ListPerWeek
    found = False
    For Each MonthPerson In MonthsList(monthNr)
        If (weekPerson.ID = MonthPerson.ID) Then
            found = True
            Exit For
        End If
    Next
    
    If found = False Then
        MonthsList(monthNr).Add weekPerson
    End If
Next

Set CreateMonthlyList = MonthsList(monthNr)

End Function


' add all unique persons in the year to a list
' Note: for each person in the list, the Days property contains the highest number
'       found for that person. Resets etc. are corrected.
'       To accomplish this, as the Since-date the date of first appearance is taken. The days
'       are then calculated from that date and the column 'available per' is ignored.
' Adds bench exit data to the Dictionaries in the clsPersonInstances
' Note: the WeekAssigned gets updated in dict(person.ID) for each sheet the person is present
'   When a person stops appearing, the WeekAssigned does not get updated and the value automatically
'   represents the last full week a person was on the bench (per interval, a person can be assigned and later come back to the bench)
Function CreatePersonsYearList(WeeksList As ArrayList) As ArrayList
FuncName = "CreatePersonsYearList"
On Error GoTo Catch

Dim person As clsPerson
Dim idx As Integer
Dim ListPerWeek As New ArrayList
Dim YearList As New ArrayList
Dim dict As New Scripting.Dictionary
Dim val As Variant
Dim Days As Integer
Dim dayOfSheetDate As Integer
Dim monthOfSheetDate As Integer
Dim yearOfSheetDate As Integer
Dim sheetName As String
Dim sheetDate As Date                       ' the date of current sheet
Dim newSince As Date
Dim test As String
Dim testWeekOrderNr As Integer
Dim benchDays As Integer
'DebugPrint "Creating Repaired YearList"
Dim lastDayPrevYear As Date
Dim joiners As Integer

' set the dictionary to case-independent
dict.CompareMode = TextCompare
dict.RemoveAll

lastDayPrevYear = DateSerial(ReviewYear - 1, 12, 31)

For Each ListPerWeek In WeeksList
    ' check that the weeks appear in ascending order
    If ListPerWeek(0).weekNr <= testWeekOrderNr Then
        GoTo Catch
    End If
    testWeekOrderNr = ListPerWeek(0).weekNr
    
    For Each person In ListPerWeek
        sheetDate = DateSerial(person.yearNr, person.monthNr, person.DayNr)
        sheetName = CStr(person.yearNr) + " wk" + CStr(person.weekNr) + " " + CStr(person.DayNr) + "-" + CStr(person.monthNr)
        person.WeekAssigned = person.weekNr
        person.MonthAssigned = person.monthNr
                
        If dict.Exists(person.ID) <> True Then
            ' The first time a person appears, take the Days and work back from the
            ' date of the sheet to establish the first date of availability and enter that in Since
            person.Since = DateAdd("d", -person.Days, sheetDate)
            dict.Add key:=person.ID, Item:=person
            ' and store the start of the new bench interval
            If (dict(person.ID).StoreNewBenchJoin(person.Since, sheetDate, person.WeekAssigned, person.MonthAssigned)) <> True Then
                GoTo Catch
            Else
                Set person.IntervalList = dict(person.ID).IntervalList.Clone
            End If
        Else
          If person.WeekAssigned > 0 And ((person.weekNr - dict(person.ID).WeekAssigned) > WeeksRemoved) Then
'                DebugPrint "Break. ID: " + CStr(person.ID) + ", assigned week: " + CStr(dict(person.ID).WeekAssigned) + ", reappeared in week " + CStr(person.weekNr)
                
                ' this person was assigned, but came back on the bench
                ' store the previous interval in the dict-person
                ' reset the Since date in the dict-person to the current sheetDate, corrected for the days already available in that sheet
                newSince = DateAdd("d", -person.Days, sheetDate)
                If (dict(person.ID).StoreNewBenchExit(newSince, sheetDate)) <> True Then
                    GoTo Catch
                End If
                ' and store the start of the new bench interval
                If (dict(person.ID).StoreNewBenchJoin(newSince, sheetDate, person.WeekAssigned, person.MonthAssigned)) <> True Then
                    GoTo Catch
                End If
            End If
            
            ' keep the highest number of days across the weeks for this person
            ' calc the proper number of days based on the Since date from
            ' the first time this person appeared in the list (for the current interval)
            person.Days = DateDiff("d", dict(person.ID).Since, sheetDate)
            If person.Days < 0 Then
                DebugPrint "Days: " + CStr(person.Days)
                GoTo Catch              ' extra check
            End If
            ' copy interval data from the person in the dictionary if there is any, otherwise
            ' it will get lost during the delete/add operation below
            If dict(person.ID).CopyIntervalsTo(person) <> True Then
                DebugPrint "Error in CopyIntervalsTo"
                GoTo Catch
            End If
            ' copy the available date that was entered when the person first appeared (for the current interval)
            person.Since = dict(person.ID).Since
            ' replace the person in the dictionary with the latest version of the person
            ' Note: updating dictionary items can only be done by deleting and then adding the updated item
           If person.Days >= dict(person.ID).Days Then
                dict.Remove person.ID
                dict.Add person.ID, person
            Else
Catch:
                errorCount = errorCount + 1
                ErrorTxt = Err.Description
            End If
        End If
            
    Next person
Next ListPerWeek

' load the persons from the dict into the yearList
' if there are still persons with sub grade = 0, then set it to 1
For Each val In dict.Items
    If val.SubGrade = 0 Then
        val.SubGrade = 1
    End If
    If val.CloseLastInterval() <> True Then
        errorCount = errorCount + 1
    End If
    If val.Unit = division.unknown Then
        errorCount = errorCount + 1
        DebugPrint CStr(person.ID) + ": Unknown Division "
    End If
    YearList.Add val
Next
        
Set CreatePersonsYearList = YearList

End Function

' return a list with persons that exited the bench (1 or more times)
' persons that exit in the week of the last sheet are considered to be still on the bench, because at the
' end of the analysis period any open intervals will be closed (with CloseLastInterval()) anyway, thus always
' having an ExitWeekNr equal to LatestWeek.
Function CreateBenchExitYearList(persYearList As ArrayList) As ArrayList
FuncName = "CreateBenchExitYearList"
On Error GoTo Catch

Dim val As Variant
Dim BenchExitList As New ArrayList
Dim person As clsPerson
Dim interv As clsInterval
Dim test As String

For Each person In persYearList
    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear And interv.ExitWeekNr < latestweek Then
            BenchExitList.Add person
            Exit For
        End If
        
    Next
Next

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
    
Bye:
Set CreateBenchExitYearList = BenchExitList
End Function

' Read the workforce data from the cells in the sheet
Sub ReadWorkforce()
    DMF_Nr = Sheets("Statistics").Cells.Item(4, "E").value
    EMB_Nr = Sheets("Statistics").Cells.Item(5, "E").value
    HWE_Nr = Sheets("Statistics").Cells.Item(6, "E").value
    TAE_Nr = Sheets("Statistics").Cells.Item(7, "E").value
    TE_Nr = Sheets("Statistics").Cells.Item(8, "E").value
End Sub
