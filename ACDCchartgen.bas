Sub Create_AC_Chart(row as Integer, I_cnt As Integer, cat As Integer, pos As String)
    'Only for generating voltage/current and efficiency/current chart in ACDC application
    Dim Hz_cnt As Integer, Hz_CAT_cnt As Integer, cnt_H As Integer
    Dim Vac_cnt As Integer

    Hz_cnt = 1
    cnt_H = 1
    Hz_CAT_cnt = 0
    Vac_cnt = 0
    
    
    'Select data in pivot chart
    Do While Cells(5, 2 + Vac_cnt).Value <> ""
        Vac_cnt = Vac_cnt + 1
    Loop

    ActiveSheet.Range(Cells(4, 1), Cells(5 + I_cnt, 1 + Vac_cnt)).Select
    Selection.Copy

    'Paste on selected ACDC data
    Cells(row, 1).Select
    ActiveSheet.Paste

    'Pre-processing for data label
    If cat <> 4 Then
        Do While Cells(row, 2 + Hz_cnt).Value = ""
            Cells(row, 2 + Hz_cnt).Value = Cells(row, 2)
            Hz_cnt = Hz_cnt + 1
        Loop
        Hz_CAT_cnt = Hz_CAT_cnt + 1

        Do While Cells(row, 2 + Hz_cnt * Hz_CAT_cnt).Value <> ""
            Do While cnt_H <= Hz_cnt - 1
                Cells(row, 2 + Hz_cnt * Hz_CAT_cnt + cnt_H).Value = Cells(row, 2 + Hz_cnt * Hz_CAT_cnt)
                cnt_H = cnt_H + 1
            Loop
            Hz_CAT_cnt = Hz_CAT_cnt + 1
            cnt_H = 1
        Loop
    End If

    'Combine Vac and Hz, and delete no use row
    If cat <> 4 Then
        Rows(row + 2).Insert
        For i = 1 To Vac_cnt Step 1
		    Cells(row + 2, i + 1).Value = Cells(row + 1, i + 1) & "Vac/" & Cells(row, i + 1) & "Hz"
	    Next
        Cells(row + 2, 1).Value = Cells(row + 1, 1) & "(A)"
        For i = 1 To 2 Step 1
            Rows(row).Delete
        Next
    Else
        Rows(row).Delete
        Cells(row, 1).Value = "Vac" & "/" & Cells(4, 2) & "Hz"
        Cells(row + 1, 1).Value = "Voltage Difference (V)"
    End If

    'Delete abnormal data
    For i = 1 To 1 + I_cnt Step 1
        For j = 2 To 1 + Vac_cnt Step 1
            Select Case cat
            Case 1
                If Cells(row + i, j).Value <= 5 Then
                    Cells(row + i, j) = ""
                End If
            Case 2
                If Cells(row + i, j).Value <= 0.3 Then
                    Cells(row + i, j) = ""
                End If
            Case 3
                If Cells(row + i, j).Value >= 1 Then
                    Cells(row + i, j) = ""
                End If
            End Select
        Next
    Next

    'Create chart
    ActiveSheet.Range(Cells(row, 1), Cells(row + I_cnt, 1 + Vac_cnt)).Select
    ActiveSheet.Shapes.AddChart.Select

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        If cat = 4 Then
            .ChartType = xlXYScatter
            .SetSourceData Source := ActiveSheet.Range(Cells(row, 2), Cells(row + I_cnt, 1 + Vac_cnt))
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "AC Instrument Voltage (Vrms)"
            .Axes(xlCategory).MinimumScale = Cells(row, 2).Value
            .Axes(xlCategory).MaximumScale = Cells(row, 1 + Vac_cnt).Value
            .Axes(xlCategory).Select
            .Axes(xlCategory).MajorUnit = 10
        Else
            .ChartType = xlXYScatterLines
            .SetSourceData Source := ActiveSheet.Range(Cells(row, 1), Cells(row + I_cnt, 1 + Vac_cnt)), PlotBy := xlColumns
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Curret Load (A)"
            .Axes(xlCategory).MinimumScale = Cells(row + 1, 1).Value
            .Axes(xlCategory).MaximumScale = Cells(row + I_cnt, 1).Value
        End If
        .Legend.Position = xlLegendPositionTop
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
        Case 3
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage Difference (V)"
            .Axes(xlValue).MinimumScale = -0.01
            .Axes(xlValue).MaximumScale = 0.08
        Case 4
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage Difference (V)"
            .Axes(xlValue).MinimumScale = 0.5
            .Axes(xlValue).MaximumScale = 3
        End Select
    End With

    Select Case cat
    Case 2
        'Set efficiency data to 0.00%
        Range(Cells(row + 1, 2), Cells(row + I_cnt, 1 + Vac_cnt)).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    Case 3
        'Set Voltage Difference chart to xlBubble
        ActiveChart.ChartType = xlBubble
        ActiveChart.ChartGroups(1).BubbleScale = 10
        ActiveChart.Axes(xlCategory).Select
        Selection.TickLabelPosition = xlLow
    Case 4
        'Set Voltage Difference chart to xlBubble
        ActiveChart.ChartType = xlBubble
        ActiveChart.ChartGroups(1).BubbleScale = 10
    End Select

    'Modify chart size and color of gridline
    With ActiveSheet.ChartObjects(cat)
        .Activate
        If cat = 4 Then
            .Height = 450
            .Width = 900
        Else
            .Height = 950
            .Width = 900
        End If
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
    If cat = 4 Then 
        Selection.Delete
    Else
        Selection.Format.TextFrame2.TextRange.font.Size = 18
    End If
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
    'Move position of the chart
    ActiveSheet.ChartObjects(cat).Activate
    ActiveChart.Parent.Cut
    Range(pos).Select
    ActiveSheet.Paste
End Sub

Sub acdc_pivot_chart()
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim cnt As Integer, log_name As String
    Dim data_name As String
    Dim comment_name As String
    Dim col_cnt As Integer
    Dim cat As Integer '1. Voltage; 2. Efficiency; 3. AC/DC Voltage Difference; 4. AC Voltage Difference

    cnt = 1
    col_cnt = 0

    'Type in the sheet name under analysis
    log_name = Inputbox("In AC/DC analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
    Do While Sheets(cnt).Name <> log_name
        If log_name = "end" Then
            Exit Do
        End If

        cnt = cnt + 1
        If cnt > Sheets.Count Then
            cnt = 1
            log_name = Inputbox("In AC/DC analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
        End If
    Loop

    If  log_name <> "end" Then
        Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType := xlDatabase, SourceData := Sheets(log_name).Range("A1").CurrentRegion.Address)
        Set PT = PTCache.CreatePivotTable (TableDestination := "", TableName:="acdc_pivot")

        'Set column in the new pivot chart
        With PT
            'Set [comment] as the filter
            .PivotFields("comment").Orientation = xlPageField
            With .PivotFields("comment")
                For cnt = 1 To .PivotItems.Count 
                    comment_name = .PivotItems(cnt).Name
                    If comment_name <> "AC/DC" Then
                        .PivotItems(comment_name).Visible = False
                    End If
                Next cnt
            End With

            'Set [ACFrequency] and [ACVoltage] as the row label
            .PivotFields("ACFrequency").Orientation = xlColumnField
            .PivotFields("ACVoltage").Orientation = xlColumnField
            'Set [Load1Current] as the column label
            .PivotFields("Load1Current").Orientation = xlRowField
            'Set the Load1Voltage in analysis at first
            With .PivotFields("Load1Voltage")
                .Orientation = xlDataField
                .Function = xlSum
            End With
        End With
    
        'Turn on the PivotTable Field List
        Application.CommandBars("PivotTable Field List").Enabled = True

        'Data calculation setting
        With ActiveSheet.PivotTables("acdc_pivot")
            .PivotFields("ACFrequency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("ACVoltage").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("Load1Current").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("Load1Voltage").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With

        Do While IsNumeric(Cells(6 + col_cnt, 1)) = True
            col_cnt = col_cnt + 1
        Loop
        
        'Create ACDC V/I curve
        Create_AC_Chart row := col_cnt + 8, I_cnt := col_cnt, cat := 1, pos := "A61"

        'Change parameter of pivot chart to Efficiency
        ActiveSheet.PivotTables("acdc_pivot").PivotFields("加總 - Load1Voltage").Orientation = xlHidden
        ActiveSheet.PivotTables("acdc_pivot").AddDataField ActiveSheet.PivotTables("acdc_pivot").PivotFields("L1/AC_Eff"), "加總 - L1/AC_Eff", xlSum

        'Create ACDC Efficiency curve
        Create_AC_Chart row := col_cnt * 2 + 8 + 2, I_cnt := col_cnt, cat := 2, pos := "M60"

        'Change parameter of pivot chart to ACDC Voltage Difference
        ActiveSheet.PivotTables("acdc_pivot").PivotFields("加總 - L1/AC_Eff").Orientation = xlHidden
        ActiveSheet.PivotTables("acdc_pivot").AddDataField ActiveSheet.PivotTables("acdc_pivot").PivotFields("acdc_Diff"), "加總 - acdc_Diff", xlSum

        'Create ACDC Voltage Difference curve
        Create_AC_Chart row := col_cnt * 3 + 8 + 2 * 2, I_cnt := col_cnt, cat := 3, pos := "Y59"

        'Change parameter of pivot chart to AC Voltage Difference
        With ActiveSheet.PivotTables("acdc_pivot")
            .PivotFields("comment").PivotItems("AC measurement").Visible = True
            .PivotFields("comment").PivotItems("AC/DC").Visible = False
            .PivotFields("加總 - acdc_Diff").Orientation = xlHidden
            .AddDataField ActiveSheet.PivotTables("acdc_pivot").PivotFields("ac_Diff"), "加總 - ac_Diff", xlSum
        End With

        'Create AC Voltage Difference curve
        Create_AC_Chart row := col_cnt * 4 + 8 + 2 * 3, I_cnt := 1, cat := 4, pos := "AK58"

        ActiveSheet.Name = "ACDC"
        ActiveWindow.Zoom = 40
        Msgbox "Test result has been generated successfully."
    Else
        Msgbox "Abort the pivot chart generating!"
    End If
End Sub