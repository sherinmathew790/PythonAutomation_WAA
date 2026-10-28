Option Explicit

Const ChartGapWidth As Integer = 100
Const ChartExtraWidth As Integer = 100

' show the number of people that are available each month, per Division
Public Function ShowMonthlyTotalByDivision(chartSheet As Worksheet, MonthlyList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyTotalByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim person As clsPerson
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
Next


For idx = 1 To 12

    For Each person In MonthlyList(idx)
        Select Case person.Unit
            Case division.TAE
                YTAE(idx) = YTAE(idx) + 1
                
            Case division.DMF
                YDMF(idx) = YDMF(idx) + 1
                
            Case division.EMB
                YEMB(idx) = YEMB(idx) + 1
                
            Case division.TE
                YTE(idx) = YTE(idx) + 1
                
            Case division.HWE
                YHWE(idx) = YHWE(idx) + 1
                
            Case Else
            ' unknown unit
            unknown(idx) = unknown(idx) + 1
        End Select
    Next
    
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Employees"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Total Availability by Division"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
   End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyTotalByDivision = errorCount
End Function


' show the number of people that are available each month, per Grade
Public Function ShowMonthlyTotalByGrade(chartSheet As Worksheet, MonthlyList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyTotalByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YArr(1 To 12) As Variant
Dim person As clsPerson
Dim YGradeA(1 To 12) As Integer
Dim YGradeB(1 To 12) As Integer
Dim YGradeC(1 To 12) As Integer
Dim YGradeD(1 To 12) As Integer
Dim YGradeE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
Next


For idx = 1 To 12

    For Each person In MonthlyList(idx)
        Select Case person.Grade
            Case "A"
                YGradeA(idx) = YGradeA(idx) + 1
                
            Case "B"
                YGradeB(idx) = YGradeB(idx) + 1
            
            Case "C"
                YGradeC(idx) = YGradeC(idx) + 1
            
            Case "D"
                YGradeD(idx) = YGradeD(idx) + 1
            
            Case "E"
                YGradeE(idx) = YGradeE(idx) + 1
            
            Case "F"
            Case Else
            ' unknown grade
            unknown(idx) = unknown(idx) + 1
        End Select
    Next
    
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Grade"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Employees"

With cht
    
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Total Availability by Grade"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With


    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyTotalByGrade = errorCount
End Function

' Show the number of people that are available per year, by Division
Public Function ShowYearByDivision(chartSheet As Worksheet, persYearList As ArrayList, left As Integer, top As Integer) As Integer
    Dim FuncName As String
    FuncName = "ShowYearByDivision"
    On Error GoTo Catch

    Dim idx As Integer
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim Srs As Series
    Dim XArr(1 To 5) As Variant
    Dim person As clsPerson
    Dim YTAE As Integer
    Dim YDMF As Integer
    Dim YEMB As Integer
    Dim YTE As Integer
    Dim YHWE As Integer
    Dim unknown As Integer
    
    For idx = 1 To 5
        XArr(idx) = ""
    Next

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
    
    ' After calculating the division totals, create a result string
    Dim result As String
    result = "Division Totals:" & vbNewLine & _
             "TAE: " & YTAE & vbNewLine & _
             "DMF: " & YDMF & vbNewLine & _
             "EMB: " & YEMB & vbNewLine & _
             "TE: " & YTE & vbNewLine & _
             "HWE: " & YHWE

    ' Print to Immediate Window for VBA debugging
    Debug.Print result
     
    Set cht = chartSheet.Shapes.AddChart.Chart
    Set chartObj = cht.Parent
    chartObj.RoundedCorners = True
    chartObj.left = left
    chartObj.top = top
    chartObj.Width = chartObj.Width + ChartExtraWidth
    cht.ChartGroups(1).GapWidth = ChartGapWidth
    cht.Axes(xlPrimary).HasTitle = True
    cht.Axes(xlPrimary).AxisTitle.Text = "Division"
    cht.Axes(xlSecondary).HasTitle = True
    cht.Axes(xlSecondary).AxisTitle.Text = "Employees"

    With cht
        .HasTitle = True
        .ChartTitle.Text = ReviewYear + " - Annual Availability by Division"
        .ChartType = xlColumnClustered
        
        For Each Srs In .SeriesCollection: Srs.Delete: Next

        Set Srs = .SeriesCollection.NewSeries
        With Srs
            .Name = "TAE(" + CStr(TAE_Nr) + ")"
            .Values = YTAE
            .XValues = XArr
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionCenter
        End With

        Set Srs = .SeriesCollection.NewSeries
        With Srs
            .Name = "DMF (" + CStr(DMF_Nr) + ")"
            .Values = YDMF
            .XValues = XArr
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionCenter
        End With

        Set Srs = .SeriesCollection.NewSeries
        With Srs
            .Name = "EMB (" + CStr(EMB_Nr) + ")"
            .Values = YEMB
            .XValues = XArr
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionCenter
        End With
        
        Set Srs = .SeriesCollection.NewSeries
        With Srs
            .Name = "TE (" + CStr(TE_Nr) + ")"
            .Values = YTE
            .XValues = XArr
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionCenter
        End With

        Set Srs = .SeriesCollection.NewSeries
        With Srs
            .Name = "HWE (" + CStr(HWE_Nr) + ")"
            .Values = YHWE
            .XValues = XArr
            .Interior.Color = RGB(0, 255, 150)
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionCenter
        End With
    End With

    GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowYearByDivision = errorCount
    
End Function


' Show the number of people that are available per year, by Grade
Public Function ShowYearByGrade(chartSheet As Worksheet, persYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowYearByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 6) As Variant
Dim person As clsPerson
Dim YGradeA As Integer
Dim YGradeB As Integer
Dim YGradeC As Integer
Dim YGradeD As Integer
Dim YGradeE As Integer
Dim unknown As Integer

For idx = 1 To 6
    XArr(idx) = ""
Next

For Each person In persYearList

    Select Case person.Grade
        Case "A"
            YGradeA = YGradeA + 1
            
        Case "B"
            YGradeB = YGradeB + 1
        
        Case "C"
            YGradeC = YGradeC + 1
        
        Case "D"
            YGradeD = YGradeD + 1
        
        Case "E"
            YGradeE = YGradeE + 1
           
        Case Else
        ' unknown grade
        unknown = unknown + 1
    End Select
    
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Grade"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Employees"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Annual Availability by Grade"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowYearByGrade = errorCount
End Function

' Show the average days available per Division (based on bench exits)
Public Function ShowYearAvgByDivision(chartSheet As Worksheet, benchYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowYearAvgByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 5) As Variant
Dim YTAE As Integer
Dim YDMF As Integer
Dim YEMB As Integer
Dim YTE As Integer
Dim YHWE As Integer
Dim nTAE As Integer
Dim nDMF As Integer
Dim nEMB As Integer
Dim nTE As Integer
Dim nHWE As Integer
Dim unknown As Integer
Dim person As clsPerson
Dim meanTAE As Integer
Dim meanDMF As Integer
Dim meanEMB As Integer
Dim meanTE As Integer
Dim meanHWE As Integer

For idx = 1 To 5
    XArr(idx) = ""
Next

For Each person In benchYearList

    Select Case person.Unit
        Case division.TAE
            YTAE = YTAE + person.DaysTotal
            nTAE = nTAE + person.IntervalList.count
            
        Case division.DMF
            YDMF = YDMF + person.DaysTotal
            nDMF = nDMF + person.IntervalList.count
            
        Case division.EMB
            YEMB = YEMB + person.DaysTotal
            nEMB = nEMB + person.IntervalList.count
            
        Case division.TE
            YTE = YTE + person.DaysTotal
            nTE = nTE + person.IntervalList.count
            
        Case division.HWE
            YHWE = YHWE + person.DaysTotal
            nHWE = nHWE + person.IntervalList.count
            
        Case Else
        ' unknown unit
        unknown = unknown + 1
    End Select
    
Next


If nTAE > 0 Then meanTAE = YTAE \ nTAE
If nDMF > 0 Then meanDMF = YDMF \ nDMF
If nEMB > 0 Then meanEMB = YEMB \ nEMB
If nTE > 0 Then meanTE = YTE \ nTE
If nHWE > 0 Then meanHWE = YHWE \ nHWE

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Division"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Days"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Average Availability Duration by Division"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE (" + CStr(nTAE) + ")"
        .Values = meanTAE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF (" + CStr(nDMF) + ")"
        .Values = meanDMF
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB (" + CStr(nEMB) + ")"
        .Values = meanEMB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
   End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE (" + CStr(nTE) + ")"
        .Values = meanTE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE (" + CStr(nHWE) + ")"
        .Values = meanHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
         .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
   End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowYearAvgByDivision = errorCount
End Function

' Show the average days available per Grade (based on bench exits) for a year
Public Function ShowYearAvgByGrade(chartSheet As Worksheet, bExitList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowYearAvgByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 6) As Variant
Dim YGradeA As Integer
Dim YGradeB As Integer
Dim YGradeC As Integer
Dim YGradeD As Integer
Dim YGradeE As Integer
Dim nA As Integer
Dim nB As Integer
Dim nC As Integer
Dim nD As Integer
Dim nE As Integer
Dim unknown As Integer
Dim person As clsPerson
Dim min As Integer
Dim max As Integer
Dim meanA As Integer
Dim meanB As Integer
Dim meanC As Integer
Dim meanD As Integer
Dim meanE As Integer

For idx = 1 To 6
    XArr(idx) = ""
Next

For Each person In bExitList

    Select Case person.Grade
        Case "A"
            YGradeA = YGradeA + person.DaysTotal
            nA = nA + person.IntervalList.count
            
        Case "B"
            YGradeB = YGradeB + person.DaysTotal
            nB = nB + person.IntervalList.count
        
        Case "C"
            YGradeC = YGradeC + person.DaysTotal
            nC = nC + person.IntervalList.count
       
        Case "D"
            YGradeD = YGradeD + person.DaysTotal
            nD = nD + person.IntervalList.count

        Case "E"
             YGradeE = YGradeE + person.DaysTotal
            nE = nE + person.IntervalList.count
            
        Case Else
        ' unknown grade
        unknown = unknown + 1
    End Select
    
Next

If nA > 0 Then meanA = YGradeA \ nA
If nB > 0 Then meanB = YGradeB \ nB
If nC > 0 Then meanC = YGradeC \ nC
If nD > 0 Then meanD = YGradeD \ nD
If nE > 0 Then meanE = YGradeE \ nE

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Grade"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Days"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Average Availability Duration by Grade"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A (" + CStr(nA) + ")"
        .Values = meanA
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B (" + CStr(nB) + ")"
        .Values = meanB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C (" + CStr(nC) + ")"
        .Values = meanC
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D (" + CStr(nD) + ")"
        .Values = meanD
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E (" + CStr(nE) + ")"
        .Values = meanE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowYearAvgByGrade = errorCount
End Function

' Show the average days available per subGrade (based on bench exits) for a year
Public Function ShowYearAvgBySubGrade(chartSheet As Worksheet, bExitYearList As ArrayList, Grade As String, left As Integer, top As Integer) As Integer
FuncName = "ShowYearAvgBySubGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(0 To 5) As Variant
Dim YGrade0 As Integer
Dim YGrade1 As Integer
Dim YGrade2 As Integer
Dim YGrade3 As Integer
Dim YGrade4 As Integer
Dim YGrade5 As Integer
Dim mean1 As Integer
Dim mean2 As Integer
Dim mean3 As Integer
Dim mean4 As Integer
Dim mean5 As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim n3 As Integer
Dim n4 As Integer
Dim n5 As Integer
Dim person As clsPerson
Dim min As Integer
Dim max As Integer
Dim unknown As Integer

For idx = 0 To 5
    XArr(idx) = ""
Next

For Each person In bExitYearList

    If person.Grade = Grade Then
        Select Case person.SubGrade
            Case "1"
                YGrade1 = YGrade1 + person.DaysTotal
                n1 = n1 + person.IntervalList.count
            Case "2"
                YGrade2 = YGrade2 + person.DaysTotal
                n2 = n2 + person.IntervalList.count
            Case "3"
                YGrade3 = YGrade3 + person.DaysTotal
                n3 = n3 + person.IntervalList.count
            Case "4"
                YGrade4 = YGrade4 + person.DaysTotal
                n4 = n4 + person.IntervalList.count
            Case "5"
                YGrade5 = YGrade5 + person.DaysTotal
                n5 = n5 + person.IntervalList.count
            Case Else
            ' unknown sub grade
            unknown = unknown + 1
        End Select
    End If
       
Next

If n1 > 0 Then mean1 = YGrade1 \ n1
If n2 > 0 Then mean2 = YGrade2 \ n2
If n3 > 0 Then mean3 = YGrade3 \ n3
If n4 > 0 Then mean4 = YGrade4 \ n4
If n5 > 0 Then mean5 = YGrade5 \ n5


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Sub-Grade"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Days"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Average Duration by sub grade " + Grade
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade " + Grade + ".1 (" + CStr(n1) + ")"
        .Values = mean1
        .XValues = XArr
         .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
   End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade " + Grade + ".2 (" + CStr(n2) + ")"
        .Values = mean2
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade " + Grade + ".3 (" + CStr(n3) + ")"
        .Values = mean3
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade " + Grade + ".4 (" + CStr(n4) + ")"
        .Values = mean4
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade " + Grade + ".5 (" + CStr(n5) + ")"
        .Values = mean5
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
End With


GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowYearAvgBySubGrade = errorCount
End Function



Public Function ShowMonthlyPercentageByDivision(chartSheet As Worksheet, MonthlyList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyPercentageByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim person As clsPerson
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
Next


For idx = 1 To 12

    For Each person In MonthlyList(idx)
        Select Case person.Unit
            Case division.TAE
                YTAE(idx) = YTAE(idx) + 1
                
            Case division.DMF
                YDMF(idx) = YDMF(idx) + 1
                
            Case division.EMB
                YEMB(idx) = YEMB(idx) + 1
                
            Case division.TE
                YTE(idx) = YTE(idx) + 1
                
            Case division.HWE
                YHWE(idx) = YHWE(idx) + 1
                
            Case Else
            ' unknown unit
            unknown(idx) = unknown(idx) + 1
        End Select
    Next
    
    YTAE(idx) = YTAE(idx) / TAE_Nr * 100
    YDMF(idx) = YDMF(idx) / DMF_Nr * 100
    YEMB(idx) = YEMB(idx) / EMB_Nr * 100
    YTE(idx) = YTE(idx) / TE_Nr * 100
    YHWE(idx) = YHWE(idx) / HWE_Nr * 100
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "% of Division size"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Relative Availability by Division"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyPercentageByDivision = errorCount
End Function


Public Function ShowTotalPercentageByDivision(chartSheet As Worksheet, persYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowTotalPercentageByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 5) As Variant
Dim person As clsPerson
Dim YTAE As Integer
Dim YDMF As Integer
Dim YEMB As Integer
Dim YTE As Integer
Dim YHWE As Integer
Dim unknown As Integer

For idx = 1 To 5
    XArr(idx) = ""
Next


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
YTAE = YTAE / TAE_Nr * 100
YDMF = YDMF / DMF_Nr * 100
YEMB = YEMB / EMB_Nr * 100
YTE = YTE / TE_Nr * 100
YHWE = YHWE / HWE_Nr * 100


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Year to date"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "% of Division size"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Annual Relative Availability by Division"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowTotalPercentageByDivision = errorCount
End Function


' show the number of bench exits and the number of weeks they were available
Public Function ShowDurationByWeeks(chartSheet As Worksheet, bExitYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowDurationByWeeks"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim YArr(1 To 53) As Integer
Dim XArr(1 To 53) As Variant
Dim weeks As Integer
Dim person As clsPerson
Dim interv As clsInterval
Dim test As String

For idx = 1 To 52
    XArr(idx) = idx
Next
XArr(53) = "52+"

For Each person In bExitYearList
    weeks = 1
    For Each interv In person.IntervalList
        If interv.IntervalDays > 7 Then
            weeks = interv.IntervalDays \ 7
        End If
        If weeks <= 52 Then
            YArr(weeks) = YArr(weeks) + 1
        Else
            YArr(53) = YArr(53) + 1
        End If
    Next
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth * 4
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "number of weeks available"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench exits"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Availability by duration"
    .ChartType = xlLineMarkers
    .HasLegend = False
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Values = YArr
        .XValues = XArr
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowDurationByWeeks = errorCount
End Function



' show the number of bench exits and the number of weeks they were available, shown by grade
Public Function ShowDurationByWeeksByGrade(chartSheet As Worksheet, bExitYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowDurationByWeeksByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim YGradeA(1 To 53) As Integer
Dim YGradeB(1 To 53) As Integer
Dim YGradeC(1 To 53) As Integer
Dim YGradeD(1 To 53) As Integer
Dim YGradeE(1 To 53) As Integer
Dim XArr(1 To 53) As Variant
Dim unknown(1 To 53) As Integer
Dim weeks As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 52
    XArr(idx) = idx
Next
XArr(53) = "52+"

For Each person In bExitYearList
    weeks = 1
    For Each interv In person.IntervalList
        If interv.IntervalDays > 7 Then
            weeks = interv.IntervalDays \ 7
        End If
        If weeks > 52 Then
            weeks = 53
        End If
        Select Case person.Grade
            Case "A"
                YGradeA(weeks) = YGradeA(weeks) + 1

            Case "B"
                YGradeB(weeks) = YGradeB(weeks) + 1

            Case "C"
                YGradeC(weeks) = YGradeC(weeks) + 1
                
            Case "D"
                YGradeD(weeks) = YGradeD(weeks) + 1

            Case "E"
                YGradeE(weeks) = YGradeE(weeks) + 1

            Case Else
            ' unknown grade
            unknown(weeks) = unknown(weeks) + 1
        End Select
    Next
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth * 6
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "number of weeks available"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench exits"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Duration by Grade"
'    .ChartType = xlColumnStacked
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With


    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
'        .Interior.Color = RGB(0, 255, 150)
   End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowDurationByWeeksByGrade = errorCount
End Function


Public Function ShowMonthlyBenchExitsByDivision(chartSheet As Worksheet, bExitYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyBenchExitsByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim lastWeekOfYear As Integer
Dim intervalMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next

For Each person In bExitYearList

    For Each interv In person.IntervalList
        intervalMonth = interv.ExitMonthNr
        Select Case person.Unit
            Case division.TAE
                YTAE(intervalMonth) = YTAE(intervalMonth) + 1
    
            Case division.DMF
                YDMF(intervalMonth) = YDMF(intervalMonth) + 1
    
            Case division.EMB
                YEMB(intervalMonth) = YEMB(intervalMonth) + 1
    
            Case division.TE
                YTE(intervalMonth) = YTE(intervalMonth) + 1
    
            Case division.HWE
                YHWE(intervalMonth) = YHWE(intervalMonth) + 1
    
            Case Else
            ' unknown unit
            unknown(intervalMonth) = unknown(intervalMonth) + 1
        End Select
    Next
    
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench exits"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Bench Exits by Division"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyBenchExitsByDivision = errorCount
End Function


Public Function ShowMonthlyBenchExitsByGrade(chartSheet As Worksheet, bExitYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyBenchExitsByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YArr(1 To 12) As Variant
Dim YGradeA(1 To 12) As Integer
Dim YGradeB(1 To 12) As Integer
Dim YGradeC(1 To 12) As Integer
Dim YGradeD(1 To 12) As Integer
Dim YGradeE(1 To 12) As Integer
Dim YGradeF(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim intervalMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next

For Each person In bExitYearList

    For Each interv In person.IntervalList
        intervalMonth = interv.ExitMonthNr
        Select Case person.Grade
            Case "A"
                YGradeA(intervalMonth) = YGradeA(intervalMonth) + 1

            Case "B"
                YGradeB(intervalMonth) = YGradeB(intervalMonth) + 1

            Case "C"
                YGradeC(intervalMonth) = YGradeC(intervalMonth) + 1

            Case "D"
                YGradeD(intervalMonth) = YGradeD(intervalMonth) + 1

            Case "E"
                YGradeE(intervalMonth) = YGradeE(intervalMonth) + 1

            Case Else
            ' unknown grade
            unknown(intervalMonth) = unknown(intervalMonth) + 1
        End Select
    Next
   
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench exits"

With cht
    
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Bench Exits by Grade"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With


    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
   End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyBenchExitsByGrade = errorCount
End Function


Public Function ShowMonthlyBenchEntriesByDivision(chartSheet As Worksheet, persYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyBenchEntriesByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim lastWeekOfYear As Integer
Dim entryMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next

For Each person In persYearList

    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear Then
            entryMonth = interv.JoinMonthNr
            Select Case person.Unit
                Case division.TAE
                    YTAE(entryMonth) = YTAE(entryMonth) + 1
        
                Case division.DMF
                    YDMF(entryMonth) = YDMF(entryMonth) + 1
        
                Case division.EMB
                    YEMB(entryMonth) = YEMB(entryMonth) + 1
        
                Case division.TE
                    YTE(entryMonth) = YTE(entryMonth) + 1
        
                Case division.HWE
                    YHWE(entryMonth) = YHWE(entryMonth) + 1
        
                Case Else
                ' unknown unit
                unknown(entryMonth) = unknown(entryMonth) + 1
            End Select
        End If
    Next
    
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench joins"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Bench joins by Division"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyBenchEntriesByDivision = errorCount
End Function


' show the number of people that are available each month, per Grade
Public Function ShowMonthlyBenchEntriesByGrade(chartSheet As Worksheet, persYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyBenchEntriesByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YArr(1 To 12) As Variant
Dim YGradeA(1 To 12) As Integer
Dim YGradeB(1 To 12) As Integer
Dim YGradeC(1 To 12) As Integer
Dim YGradeD(1 To 12) As Integer
Dim YGradeE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim entryMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next


For Each person In persYearList

    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear Then
            entryMonth = interv.JoinMonthNr
            Select Case person.Grade
                Case "A"
                    YGradeA(entryMonth) = YGradeA(entryMonth) + 1
                    
                Case "B"
                    YGradeB(entryMonth) = YGradeB(entryMonth) + 1
                
                Case "C"
                    YGradeC(entryMonth) = YGradeC(entryMonth) + 1
                
                Case "D"
                    YGradeD(entryMonth) = YGradeD(entryMonth) + 1
                
                Case "E"
                    YGradeE(entryMonth) = YGradeE(entryMonth) + 1
                
                Case "F"
                Case Else
                ' unknown grade
                unknown(entryMonth) = unknown(entryMonth) + 1
            End Select
        End If
    Next
    
Next


Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Bench Joins"

With cht
    
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Bench Joins by Grade"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With


    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyBenchEntriesByGrade = errorCount
End Function


Public Function ShowMonthlyRelativeExitsByDivision(chartSheet As Worksheet, bExitYearList As ArrayList, MonthlyList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyRelativeExitsByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim XYTAE(1 To 12) As Integer
Dim XYDMF(1 To 12) As Integer
Dim XYEMB(1 To 12) As Integer
Dim XYTE(1 To 12) As Integer
Dim XYHWE(1 To 12) As Integer
Dim exitMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next

' get the number of persons per month for each division
For idx = 1 To 12

    For Each person In MonthlyList(idx)
        Select Case person.Unit
            Case division.TAE
                YTAE(idx) = YTAE(idx) + 1
                
            Case division.DMF
                YDMF(idx) = YDMF(idx) + 1
                
            Case division.EMB
                YEMB(idx) = YEMB(idx) + 1
                
            Case division.TE
                YTE(idx) = YTE(idx) + 1
                
            Case division.HWE
                YHWE(idx) = YHWE(idx) + 1
                
            Case Else
            ' unknown unit
            unknown(idx) = unknown(idx) + 1
        End Select
    Next
    
Next

' get the exits per month for each division
For Each person In bExitYearList

    For Each interv In person.IntervalList
        exitMonth = interv.ExitMonthNr
        Select Case person.Unit
            Case division.TAE
                XYTAE(exitMonth) = XYTAE(exitMonth) + 1
    
            Case division.DMF
                XYDMF(exitMonth) = XYDMF(exitMonth) + 1
    
            Case division.EMB
                XYEMB(exitMonth) = XYEMB(exitMonth) + 1
    
            Case division.TE
                XYTE(exitMonth) = XYTE(exitMonth) + 1
    
            Case division.HWE
                XYHWE(exitMonth) = XYHWE(exitMonth) + 1
    
            Case Else
            ' unknown unit
            unknown(exitMonth) = unknown(exitMonth) + 1
        End Select
    Next
    
Next

For idx = 1 To 12
    If YTAE(idx) > 0 Then YTAE(idx) = XYTAE(idx) / YTAE(idx) * 100
    If YDMF(idx) > 0 Then YDMF(idx) = XYDMF(idx) / YDMF(idx) * 100
    If YEMB(idx) > 0 Then YEMB(idx) = XYEMB(idx) / YEMB(idx) * 100
    If YTE(idx) > 0 Then YTE(idx) = XYTE(idx) / YTE(idx) * 100
    If YHWE(idx) > 0 Then YHWE(idx) = XYHWE(idx) / YHWE(idx) * 100
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "% of availability"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Relative Exits by Division"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyRelativeExitsByDivision = errorCount
End Function


Public Function ShowMonthlyRelativeExitsByGrade(chartSheet As Worksheet, bExitYearList As ArrayList, MonthlyList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyRelativeExitsByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YGradeA(1 To 12) As Integer
Dim YGradeB(1 To 12) As Integer
Dim YGradeC(1 To 12) As Integer
Dim YGradeD(1 To 12) As Integer
Dim YGradeE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim XYGradeA(1 To 12) As Integer
Dim XYGradeB(1 To 12) As Integer
Dim XYGradeC(1 To 12) As Integer
Dim XYGradeD(1 To 12) As Integer
Dim XYGradeE(1 To 12) As Integer
Dim exitMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

For idx = 1 To 12
    XArr(idx) = idx
Next

' get the number of persons per month for each grade
For idx = 1 To 12

    For Each person In MonthlyList(idx)
        Select Case person.Grade
            Case "A"
                YGradeA(idx) = YGradeA(idx) + 1
                
            Case "B"
                YGradeB(idx) = YGradeB(idx) + 1
            
            Case "C"
                YGradeC(idx) = YGradeC(idx) + 1
            
            Case "D"
                YGradeD(idx) = YGradeD(idx) + 1
            
            Case "E"
                YGradeE(idx) = YGradeE(idx) + 1
            
            Case Else
            ' unknown grade
            unknown(idx) = unknown(idx) + 1
        End Select
    Next
    
Next

' get the exits per month for each grade
For Each person In bExitYearList

    For Each interv In person.IntervalList
        exitMonth = interv.ExitMonthNr
        Select Case person.Grade
            Case "A"
                XYGradeA(exitMonth) = XYGradeA(exitMonth) + 1
                
            Case "B"
                XYGradeB(exitMonth) = XYGradeB(exitMonth) + 1
            
            Case "C"
                XYGradeC(exitMonth) = XYGradeC(exitMonth) + 1
            
            Case "D"
                XYGradeD(exitMonth) = XYGradeD(exitMonth) + 1
            
            Case "E"
                XYGradeE(exitMonth) = XYGradeE(exitMonth) + 1
            
            Case Else
            ' unknown grade
            unknown(exitMonth) = unknown(exitMonth) + 1
        End Select
    Next
    
Next

For idx = 1 To 12
    If YGradeA(idx) > 0 Then YGradeA(idx) = XYGradeA(idx) / YGradeA(idx) * 100
    If YGradeB(idx) > 0 Then YGradeB(idx) = XYGradeB(idx) / YGradeB(idx) * 100
    If YGradeC(idx) > 0 Then YGradeC(idx) = XYGradeC(idx) / YGradeC(idx) * 100
    If YGradeD(idx) > 0 Then YGradeD(idx) = XYGradeD(idx) / YGradeD(idx) * 100
    If YGradeE(idx) > 0 Then YGradeE(idx) = XYGradeE(idx) / YGradeE(idx) * 100
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Month"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "% of availability"


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Relative Exits by Grade"
    .ChartType = xlColumnStacked
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyRelativeExitsByGrade = errorCount
End Function


' Show the total number of days available per Division for the analysis period.
' Note: this uses all availability, not just the intervals with joins and exits inside the period.
Public Function ShowTotalDaysByDivision(chartSheet As Worksheet, persYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowTotalDaysByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 5) As Variant
Dim YTAE As Long
Dim YDMF As Long
Dim YEMB As Long
Dim YTE As Long
Dim YHWE As Long
Dim total As Long
Dim unknown As Integer
Dim person As clsPerson


For idx = 1 To 5
    XArr(idx) = ""
Next

For Each person In persYearList

    Select Case person.Unit
        Case division.TAE
            YTAE = YTAE + person.DaysTotal
            
        Case division.DMF
            YDMF = YDMF + person.DaysTotal
            
        Case division.EMB
            YEMB = YEMB + person.DaysTotal
            
        Case division.TE
            YTE = YTE + person.DaysTotal
            
        Case division.HWE
            YHWE = YHWE + person.DaysTotal
            
        Case Else
            ' unknown unit
            unknown = unknown + 1
    End Select
    
Next

total = YTAE + YDMF + YEMB + YTE + YHWE

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Division"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Days"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Total Availability by Division"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
   End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
         .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
   End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1 + unknown
    ErrorTxt = Err.Description
Bye:
    ShowTotalDaysByDivision = errorCount
End Function


' Show the total number of days available per Grade for the analysis period.
' Note: this uses all availability, not just the intervals not just the intervals with joins and exits inside the period.
Public Function ShowTotalDaysByGrade(chartSheet As Worksheet, persYearList As ArrayList, left As Integer, top As Integer) As Integer
FuncName = "ShowTotalDaysByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 6) As Variant
Dim YGradeA As Long
Dim YGradeB As Long
Dim YGradeC As Long
Dim YGradeD As Long
Dim YGradeE As Long
Dim total As Long
Dim unknown As Integer
Dim person As clsPerson

For idx = 1 To 6
    XArr(idx) = ""
Next

For Each person In persYearList

    Select Case person.Grade
        Case "A"
            YGradeA = YGradeA + person.DaysTotal
            
        Case "B"
            YGradeB = YGradeB + person.DaysTotal
        
        Case "C"
            YGradeC = YGradeC + person.DaysTotal
       
        Case "D"
            YGradeD = YGradeD + person.DaysTotal

        Case "E"
             YGradeE = YGradeE + person.DaysTotal
            
        Case Else
            ' unknown grade
            unknown = unknown + 1
    End Select
    
Next

total = YGradeA + YGradeB + YGradeC + YGradeD + YGradeE + unknown

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlPrimary).HasTitle = True
cht.Axes(xlPrimary).AxisTitle.Text = "Grade"
cht.Axes(xlSecondary).HasTitle = True
cht.Axes(xlSecondary).AxisTitle.Text = "Days"

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Total Availability by Grade"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
        .HasDataLabels = True
        .DataLabels.Position = xlLabelPositionCenter
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1 + unknown
    ErrorTxt = Err.Description
Bye:
    ShowTotalDaysByGrade = errorCount
End Function

' show ratio of joins/ exits per month by Division
Public Function ShowMonthlyRatioByDivision(chartSheet As Worksheet, persYearList As Variant, bExitYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyRatioByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YTAE(1 To 12) As Single
Dim YDMF(1 To 12) As Single
Dim YEMB(1 To 12) As Single
Dim YTE(1 To 12) As Single
Dim YHWE(1 To 12) As Single
Dim unknown(1 To 12) As Integer
Dim lastWeekOfYear As Integer
Dim entryMonth As Integer
Dim exitMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

Dim TAE_Join(1 To 12) As Integer
Dim DMF_Join(1 To 12) As Integer
Dim EMB_Join(1 To 12) As Integer
Dim TE_Join(1 To 12) As Integer
Dim HWE_Join(1 To 12) As Integer
Dim TAE_Exit(1 To 12) As Integer
Dim DMF_Exit(1 To 12) As Integer
Dim EMB_Exit(1 To 12) As Integer
Dim TE_Exit(1 To 12) As Integer
Dim HWE_Exit(1 To 12) As Integer
Dim GreenArea(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
    GreenArea(idx) = 1
Next

For Each person In persYearList

    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear Then
            entryMonth = interv.JoinMonthNr
            Select Case person.Unit
                Case division.TAE
                    TAE_Join(entryMonth) = TAE_Join(entryMonth) + 1
        
                Case division.DMF
                    DMF_Join(entryMonth) = DMF_Join(entryMonth) + 1
        
                Case division.EMB
                    EMB_Join(entryMonth) = EMB_Join(entryMonth) + 1
        
                Case division.TE
                    TE_Join(entryMonth) = TE_Join(entryMonth) + 1
        
                Case division.HWE
                    HWE_Join(entryMonth) = HWE_Join(entryMonth) + 1
        
                Case Else
                ' unknown unit
                unknown(entryMonth) = unknown(entryMonth) + 1
            End Select
        End If
    Next
    
Next


For Each person In bExitYearList

    For Each interv In person.IntervalList
        exitMonth = interv.ExitMonthNr
        Select Case person.Unit
            Case division.TAE
                TAE_Exit(exitMonth) = TAE_Exit(exitMonth) + 1
    
            Case division.DMF
                DMF_Exit(exitMonth) = DMF_Exit(exitMonth) + 1
    
            Case division.EMB
                EMB_Exit(exitMonth) = EMB_Exit(exitMonth) + 1
    
            Case division.TE
                TE_Exit(exitMonth) = TE_Exit(exitMonth) + 1
    
            Case division.HWE
                HWE_Exit(exitMonth) = HWE_Exit(exitMonth) + 1
    
            Case Else
            ' unknown unit
            unknown(exitMonth) = unknown(exitMonth) + 1
        End Select
    Next
    
Next

For idx = 1 To 12
    If TAE_Exit(idx) > 0# Then YTAE(idx) = TAE_Join(idx) / TAE_Exit(idx)
    If DMF_Exit(idx) > 0# Then YDMF(idx) = DMF_Join(idx) / DMF_Exit(idx)
    If EMB_Exit(idx) > 0# Then YEMB(idx) = EMB_Join(idx) / EMB_Exit(idx)
    If TE_Exit(idx) > 0# Then YTE(idx) = TE_Join(idx) / TE_Exit(idx)
    If HWE_Exit(idx) > 0# Then YHWE(idx) = HWE_Join(idx) / HWE_Exit(idx)
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlCategory).HasTitle = True
cht.Axes(xlCategory).AxisTitle.Text = "Month"
cht.Axes(xlValue).HasTitle = True
cht.Axes(xlValue).AxisTitle.Text = "Ratio"
cht.ChartArea.Height = 300
cht.Axes(xlValue).MajorUnit = 0.5
cht.Axes(xlValue).MinorUnit = 0.5
cht.Axes(xlValue).MinorTickMark = xlTickMarkOutside
cht.Axes(xlValue).HasMinorGridlines = True

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Ratio joins / Exits by Division"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

    Set Srs = .SeriesCollection.NewSeries
        .Legend.LegendEntries(.SeriesCollection.count).Delete 'remove from legend
    With Srs
        .Type = xlArea
        .Values = GreenArea
        .Interior.Color = RGB(220, 255, 220)
    End With

End With


GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyRatioByDivision = errorCount
End Function


' show ratio of joins/ exits per month by Grade
Public Function ShowMonthlyRatioByGrade(chartSheet As Worksheet, persYearList As Variant, bExitYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyRatioByGrade"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim YGradeA(1 To 12) As Single
Dim YGradeB(1 To 12) As Single
Dim YGradeC(1 To 12) As Single
Dim YGradeD(1 To 12) As Single
Dim YGradeE(1 To 12) As Single
Dim XArr(1 To 12) As Variant
Dim unknown(1 To 12) As Integer
Dim lastWeekOfYear As Integer
Dim entryMonth As Integer
Dim exitMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

Dim A_Join(1 To 12) As Integer
Dim B_Join(1 To 12) As Integer
Dim C_Join(1 To 12) As Integer
Dim D_Join(1 To 12) As Integer
Dim E_Join(1 To 12) As Integer
Dim A_Exit(1 To 12) As Integer
Dim B_Exit(1 To 12) As Integer
Dim C_Exit(1 To 12) As Integer
Dim D_Exit(1 To 12) As Integer
Dim E_Exit(1 To 12) As Integer
Dim GreenArea(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
    GreenArea(idx) = 1
Next

For Each person In persYearList

    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear Then
            entryMonth = interv.JoinMonthNr
            Select Case person.Grade
                Case "A"
                    A_Join(entryMonth) = A_Join(entryMonth) + 1
        
                Case "B"
                    B_Join(entryMonth) = B_Join(entryMonth) + 1
        
                Case "C"
                    C_Join(entryMonth) = C_Join(entryMonth) + 1
        
                Case "D"
                    D_Join(entryMonth) = D_Join(entryMonth) + 1
        
                Case "E"
                    E_Join(entryMonth) = E_Join(entryMonth) + 1
        
                Case Else
                ' unknown unit
                unknown(entryMonth) = unknown(entryMonth) + 1
            End Select
        End If
    Next
    
Next


For Each person In bExitYearList

    For Each interv In person.IntervalList
        exitMonth = interv.ExitMonthNr
        Select Case person.Grade
            Case "A"
                A_Exit(exitMonth) = A_Exit(exitMonth) + 1
    
            Case "B"
                B_Exit(exitMonth) = B_Exit(exitMonth) + 1
    
            Case "C"
                C_Exit(exitMonth) = C_Exit(exitMonth) + 1
    
            Case "D"
                D_Exit(exitMonth) = D_Exit(exitMonth) + 1
    
            Case "E"
                E_Exit(exitMonth) = E_Exit(exitMonth) + 1
    
            Case Else
            ' unknown unit
            unknown(exitMonth) = unknown(exitMonth) + 1
        End Select
    Next
    
Next

For idx = 1 To 12
    If A_Exit(idx) > 0 Then YGradeA(idx) = A_Join(idx) / A_Exit(idx)
    If B_Exit(idx) > 0 Then YGradeB(idx) = B_Join(idx) / B_Exit(idx)
    If C_Exit(idx) > 0 Then YGradeC(idx) = C_Join(idx) / C_Exit(idx)
    If D_Exit(idx) > 0 Then YGradeD(idx) = D_Join(idx) / D_Exit(idx)
    If E_Exit(idx) > 0 Then YGradeE(idx) = E_Join(idx) / E_Exit(idx)
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlCategory).HasTitle = True
cht.Axes(xlCategory).AxisTitle.Text = "Month"
cht.Axes(xlValue).HasTitle = True
cht.Axes(xlValue).AxisTitle.Text = "Ratio"
cht.ChartArea.Height = 300
cht.Axes(xlValue).MajorUnit = 0.5
cht.Axes(xlValue).MinorUnit = 0.5


With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Ratio joins / Exits by Grade"
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade A"
        .Values = YGradeA
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade B"
        .Values = YGradeB
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade C"
        .Values = YGradeC
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade D"
        .Values = YGradeD
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "grade E"
        .Values = YGradeE
        .XValues = XArr
        .Interior.Color = RGB(255, 0, 0)
    End With

    Set Srs = .SeriesCollection.NewSeries
        .Legend.LegendEntries(.SeriesCollection.count).Delete 'remove from legend
    With Srs
        .Type = xlArea
        .Values = GreenArea
        .Interior.Color = RGB(220, 255, 220)
    End With
End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyRatioByGrade = errorCount
End Function


' show ratio of joins - exits per month by Division
Public Function ShowMonthlyDeltaByDivision(chartSheet As Worksheet, persYearList As Variant, bExitYearList As Variant, left As Integer, top As Integer) As Integer
FuncName = "ShowMonthlyDeltaByDivision"
On Error GoTo Catch

Dim idx As Integer
Dim chartObj As ChartObject
Dim cht As Chart
Dim Srs As Series
Dim XArr(1 To 12) As Variant
Dim YTAE(1 To 12) As Integer
Dim YDMF(1 To 12) As Integer
Dim YEMB(1 To 12) As Integer
Dim YTE(1 To 12) As Integer
Dim YHWE(1 To 12) As Integer
Dim unknown(1 To 12) As Integer
Dim lastWeekOfYear As Integer
Dim entryMonth As Integer
Dim exitMonth As Integer
Dim person As clsPerson
Dim interv As clsInterval

Dim TAE_Join(1 To 12) As Integer
Dim DMF_Join(1 To 12) As Integer
Dim EMB_Join(1 To 12) As Integer
Dim TE_Join(1 To 12) As Integer
Dim HWE_Join(1 To 12) As Integer
Dim TAE_Exit(1 To 12) As Integer
Dim DMF_Exit(1 To 12) As Integer
Dim EMB_Exit(1 To 12) As Integer
Dim TE_Exit(1 To 12) As Integer
Dim HWE_Exit(1 To 12) As Integer

For idx = 1 To 12
    XArr(idx) = idx
Next

For Each person In persYearList

    For Each interv In person.IntervalList
        If interv.JoinYear = ReviewYear Then
            entryMonth = interv.JoinMonthNr
            Select Case person.Unit
                Case division.TAE
                    TAE_Join(entryMonth) = TAE_Join(entryMonth) + 1
        
                Case division.DMF
                    DMF_Join(entryMonth) = DMF_Join(entryMonth) + 1
        
                Case division.EMB
                    EMB_Join(entryMonth) = EMB_Join(entryMonth) + 1
        
                Case division.TE
                    TE_Join(entryMonth) = TE_Join(entryMonth) + 1
        
                Case division.HWE
                    HWE_Join(entryMonth) = HWE_Join(entryMonth) + 1
        
                Case Else
                ' unknown unit
                unknown(entryMonth) = unknown(entryMonth) + 1
            End Select
        End If
    Next
    
Next


For Each person In bExitYearList

    For Each interv In person.IntervalList
        exitMonth = interv.ExitMonthNr
        Select Case person.Unit
            Case division.TAE
                TAE_Exit(exitMonth) = TAE_Exit(exitMonth) + 1
    
            Case division.DMF
                DMF_Exit(exitMonth) = DMF_Exit(exitMonth) + 1
    
            Case division.EMB
                EMB_Exit(exitMonth) = EMB_Exit(exitMonth) + 1
    
            Case division.TE
                TE_Exit(exitMonth) = TE_Exit(exitMonth) + 1
    
            Case division.HWE
                HWE_Exit(exitMonth) = HWE_Exit(exitMonth) + 1
    
            Case Else
            ' unknown unit
            unknown(exitMonth) = unknown(exitMonth) + 1
        End Select
    Next
    
Next

For idx = 1 To 12
    YTAE(idx) = TAE_Join(idx) - TAE_Exit(idx)
    YDMF(idx) = DMF_Join(idx) - DMF_Exit(idx)
    YEMB(idx) = EMB_Join(idx) - EMB_Exit(idx)
    YTE(idx) = TE_Join(idx) - TE_Exit(idx)
    YHWE(idx) = HWE_Join(idx) - HWE_Exit(idx)
Next

Set cht = chartSheet.Shapes.AddChart.Chart
Set chartObj = cht.Parent
chartObj.RoundedCorners = True
chartObj.left = left
chartObj.top = top
chartObj.Width = chartObj.Width + ChartExtraWidth * 6
cht.ChartGroups(1).GapWidth = ChartGapWidth
cht.Axes(xlCategory).HasTitle = True
cht.Axes(xlCategory).AxisTitle.Text = "Month"
cht.Axes(xlValue).HasTitle = True
cht.Axes(xlValue).AxisTitle.Text = "<-- Exits  Joins -->"
cht.ChartArea.Height = 300

With cht
    .HasTitle = True
    .ChartTitle.Text = ReviewYear + " - Monthly Delta: Joins - Exits by Division"
'    .ChartType = xlColumnStacked
    .ChartType = xlColumnClustered
    
    For Each Srs In .SeriesCollection: Srs.Delete: Next

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TAE"
        .Values = YTAE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "DMF"
        .Values = YDMF
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "EMB"
        .Values = YEMB
        .XValues = XArr
    End With
    
    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "TE"
        .Values = YTE
        .XValues = XArr
    End With

    Set Srs = .SeriesCollection.NewSeries
    With Srs
        .Name = "HWE"
        .Values = YHWE
        .XValues = XArr
        .Interior.Color = RGB(0, 255, 150)
    End With

End With

GoTo Bye
Catch:
    errorCount = errorCount + 1
    ErrorTxt = Err.Description
Bye:
    ShowMonthlyDeltaByDivision = errorCount
End Function


