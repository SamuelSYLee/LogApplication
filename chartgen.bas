Sub Create_ACDC_Chart()
    'Only for generating voltage/current and efficiency/current chart in ACDC application
    Dim Hz_cnt As Integer, Hz_CAT_cnt As Integer, cnt As Integer
    Dim Vac_cnt As Integer, I_cnt As Integer
    Dim cat As String

    Hz_cnt = 1
    cnt = 1
    Hz_CAT_cnt = 0
    Vac_cnt = 0
    I_cnt = 0
    
    'Select data in pivot chart
    Do While Cells(5, 2 + Vac_cnt).Value <> ""
        Vac_cnt = Vac_cnt + 1
    Loop

    Do While IsNumeric(Cells(6 + I_cnt, 1)) = True
        I_cnt = I_cnt + 1
    Loop

    ActiveSheet.Range(Cells(4, 1), Cells(5 + I_cnt, 1 + Vac_cnt)).Select
    Selection.Copy

    'Create new sheet, and paste on selected ACDC data
    Sheets.Add
    ActiveSheet.Paste

    'Pre-processing for data label
    Do While Cells(1, 2 + Hz_cnt).Value = ""
        Cells(1, 2 + Hz_cnt).Value = Cells(1, 2)
        Hz_cnt = Hz_cnt + 1
    Loop
    Hz_CAT_cnt = Hz_CAT_cnt + 1

    Do While Cells(1, 2 + Hz_cnt * Hz_CAT_cnt).Value <> ""
        Do While cnt <= Hz_cnt - 1
            Cells(1, 2 + Hz_cnt * Hz_CAT_cnt + cnt).Value = Cells(1, 2 + Hz_cnt * Hz_CAT_cnt)
            cnt = cnt + 1
        Loop
        Hz_CAT_cnt = Hz_CAT_cnt + 1
        cnt = 1
    Loop

    'Combine Vac and Hz, and delete no use row
    Rows("3").Insert
    For i = 1 To Vac_cnt Step 1
		Cells(3, i + 1).Value = Cells(2, i + 1) & "Vac/" & Cells(1, i + 1) & "Hz"
	Next
    Cells(3, 1).Value = Cells(2, 1) & "(A)"
    For i = 1 To 2 Step 1
        Rows(1).Delete
    Next

    cat = Inputbox("Please type in the category (1 or 2) of y-axis, refer to the following description:" & vbCrlf & "1. Voltage" & vbCrlf & "2. Efficiency")

    'Delete abnormal data
    For i = 2 To 1 + I_cnt Step 1
        For j = 2 To 1 + Vac_cnt Step 1
            Select Case cat
            Case 1
                If Cells(i, j).Value <= 5 Then
                    Cells(i, j) = ""
                End If
            Case 2
                If Cells(i, j).Value <= 0.3 Then
                    Cells(i, j) = ""
                End If
            End Select
        Next
    Next

    'Create XYScatter chart
    ActiveSheet.Range(Cells(1, 1), Cells(1 + I_cnt, 1 + Vac_cnt)).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterLines
    ActiveChart.SetSourceData Source := ActiveSheet.Range(Cells(1, 1), Cells(1 + I_cnt, 1 + Vac_cnt)), PlotBy := xlColumns
    ActiveChart.Legend.Position = xlLegendPositionTop

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Curret Load (A)"
        .Axes(xlCategory).MinimumScale = Cells(2, 1).Value
        .Axes(xlCategory).MaximumScale = Cells(1 + I_cnt, 1).Value
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue).HasMajorGridlines = True

        Select Case cat
        Case 1
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage (V)"
            .Axes(xlValue).MinimumScale = 15.97
            .Axes(xlValue).MaximumScale = 16.04
        Case 2
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Efficiency (%)"
            .Axes(xlValue).MinimumScale = 0.75
            .Axes(xlValue).MaximumScale = 0.95
        End Select
    End With

    'Set efficiency data to 0.00%
    If cat = 2 Then
        Range(Cells(2, 2), Cells(1 + I_cnt, 1 + Vac_cnt)).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    End If

    'Modify chart size and color of gridline
    With ActiveSheet.ChartObjects(1)
        .Activate
        .Height = 950
        .Width = 900
    End With
    ActiveChart.Axes(xlCategory).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(217, 217, 217)
        .Transparency = 0
    End With
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(217, 217, 217)
        .Transparency = 0
    End With
    ActiveChart.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(217, 217, 217)
        .Transparency = 0
    End With
    ActiveChart.Axes(xlValue).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(217, 217, 217)
        .Transparency = 0
    End With

    'Modify all font to 18
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.font.Size = 18
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.font.Size = 18
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.font.Size = 18
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Format.TextFrame2.TextRange.font.Size = 18
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Format.TextFrame2.TextRange.font.Size = 18
    'Delete chart border
    ActiveSheet.Shapes(ActiveChart.Parent.Name).Line.Visible = msoFalse  
End Sub
