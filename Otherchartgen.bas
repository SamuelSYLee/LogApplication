Sub Create_Other_Chart_wData(row as Integer, I_cnt As Integer, cat As Integer, pos As String)
    Dim V_cnt As Integer

    V_cnt = 0
    
    'Select data in pivot chart
    Do While IsNumeric(Cells(4, 2 + V_cnt)) = True
        V_cnt = V_cnt + 1
    Loop

    ActiveSheet.Range(Cells(4, 1), Cells(4 + I_cnt, 1 + V_cnt)).Select
    Selection.Copy

    'Paste on selected MAIN data
    Cells(row, 1).Select
    ActiveSheet.Paste

    'Combine row label
    Rows(row + 1).Insert
    For i = 1 To V_cnt Step 1
		Cells(row + 1, i + 1).Value = Cells(row, i + 1) & "V"
	Next
    Cells(row + 1, 1).Value = Cells(row, 1) & "(A)"
    Rows(row).Delete

    'Delete abnormal data
    For i = 1 To 1 + I_cnt Step 1
        For j = 2 To 1 + V_cnt Step 1
            Select Case cat
            Case 2
                If Cells(row + i, j).Value <= 0.3 Then
                    Cells(row + i, j) = ""
                End If
            Case 3
                If Cells(row + i, j).Value >= 0.4 Then
                    Cells(row + i, j) = ""
                End If
            End Select
        Next
    Next

    'Create chart
    ActiveSheet.Range(Cells(row, 1), Cells(row + I_cnt, 1 + V_cnt)).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLineMarkers
    ActiveChart.SetSourceData Source := ActiveSheet.Range(Cells(row, 2), Cells(row + I_cnt, 1 + V_cnt)), PlotBy := xlColumns
    ActiveChart.Legend.Position = xlLegendPositionTop

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Curret Load (A)"
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue).HasMajorGridlines = True
        .FullSeriesCollection(1).XValues = Range(Cells(row + 1, 1), Cells(row + I_cnt, 1))
        .SetElement (msoElementDataTableWithLegendKeys)

        Select Case cat
        Case 1
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage (V)"
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 18
        Case 2
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Efficiency (%)"
            .Axes(xlValue).MinimumScale = 0.81
            .Axes(xlValue).MaximumScale = 0.97
        Case 3
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage Difference (V)"
            .Axes(xlValue).MinimumScale = 0.00
            .Axes(xlValue).MaximumScale = 0.12
        End Select
    End With

    Select Case cat
    Case 1
        Range(Cells(row + 1, 2), Cells(row + I_cnt, 1 + V_cnt)).Select
        Selection.NumberFormat = "0.0"
    Case 2
        'Set efficiency data to 0.00%
        Range(Cells(row + 1, 2), Cells(row + I_cnt, 1 + V_cnt)).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    Case 3
        'Set Voltage Difference chart to xlBubble
        With ActiveChart
            .ChartType = xlBubble
            .ChartGroups(1).BubbleScale = 10
            .Axes(xlCategory).MinimumScale = Cells(row + 1, 1).Value
            .Axes(xlCategory).MaximumScale = Cells(row + I_cnt, 1).Value
            .Axes(xlCategory).Select
        End With
        Selection.TickLabelPosition = xlLow
    End Select

    'Modify chart size and color of gridline
    With ActiveSheet.ChartObjects(cat)
        .Activate
        Select Case cat
        Case 1
            .Height = 5000
            .Width = 2500
        Case 2
            .Height = 1500
            .Width = 1500
        Case 3
            .Height = 900
            .Width = 900
        End Select
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
    
    If cat <> 3 Then
        ActiveChart.DataTable.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(217, 217, 217)
            .Transparency = 0
        End With
    End If

    'Modify all font
    Select Case cat
        Case 1
            ActiveChart.Legend.Select
            Selection.Format.TextFrame2.TextRange.font.Size = 25
            ActiveChart.Axes(xlValue).Select
            Selection.TickLabels.font.Size = 25
            ActiveChart.Axes(xlCategory).Select
            Selection.TickLabels.font.Size = 25
            ActiveChart.Axes(xlValue).AxisTitle.Select
            Selection.Format.TextFrame2.TextRange.font.Size = 25
            ActiveChart.Axes(xlCategory).AxisTitle.Select
            Selection.Format.TextFrame2.TextRange.font.Size = 25
            ActiveChart.DataTable.Select
            Selection.Format.TextFrame2.TextRange.font.Size = 25
        Case 2
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
            ActiveChart.DataTable.Select
            Selection.Format.TextFrame2.TextRange.font.Size = 18
        Case 3
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
    End Select
    'Delete chart border
    ActiveSheet.Shapes(ActiveChart.Parent.Name).Line.Visible = msoFalse  
    'Move position of the chart
    ActiveSheet.ChartObjects(cat).Activate
    ActiveChart.Parent.Cut
    Range(pos).Select
    ActiveSheet.Paste
End Sub

Sub Create_Other_Chart(row as Integer, I_cnt As Integer, cat As Integer, pos As String)
    Dim V_cnt As Integer

    V_cnt = 0
    
    'Select data in pivot chart
    Do While IsNumeric(Cells(4, 2 + V_cnt)) = True
        V_cnt = V_cnt + 1
    Loop

    ActiveSheet.Range(Cells(4, 1), Cells(4 + I_cnt, 1 + V_cnt)).Select
    Selection.Copy

    'Paste on selected MAIN data
    Cells(row, 1).Select
    ActiveSheet.Paste

    'Combine row label
    Rows(row + 1).Insert
    For i = 1 To V_cnt Step 1
		Cells(row + 1, i + 1).Value = Cells(row, i + 1) & "V"
	Next
    Cells(row + 1, 1).Value = Cells(row, 1) & "(A)"
    Rows(row).Delete

    'Delete abnormal data
    For i = 1 To 1 + I_cnt Step 1
        For j = 2 To 1 + V_cnt Step 1
            Select Case cat
            Case 2
                If Cells(row + i, j).Value <= 0.3 Then
                    Cells(row + i, j) = ""
                End If
            Case 3
                If Cells(row + i, j).Value >= 0.4 Then
                    Cells(row + i, j) = ""
                End If
            End Select
        Next
    Next

    'Create chart
    ActiveSheet.Range(Cells(row, 1), Cells(row + I_cnt, 1 + V_cnt)).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterLines
    ActiveChart.SetSourceData Source := ActiveSheet.Range(Cells(row, 1), Cells(row + I_cnt, 1 + V_cnt)), PlotBy := xlColumns
    ActiveChart.Legend.Position = xlLegendPositionTop

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Curret Load (A)"
        .Axes(xlCategory).MinimumScale = Cells(row + 1, 1).Value
        .Axes(xlCategory).MaximumScale = Cells(row + I_cnt, 1).Value
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue).HasMajorGridlines = True

        Select Case cat
        Case 1
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage (V)"
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 18
        Case 2
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Efficiency (%)"
            .Axes(xlValue).MinimumScale = 0.81
            .Axes(xlValue).MaximumScale = 0.97
        Case 3
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage Difference (V)"
            .Axes(xlValue).MinimumScale = 0.00
            .Axes(xlValue).MaximumScale = 0.12
        End Select
    End With

    Select Case cat
    Case 1
        Range(Cells(row + 1, 2), Cells(row + I_cnt, 1 + V_cnt)).Select
        Selection.NumberFormat = "0.0"
    Case 2
        'Set efficiency data to 0.00%
        Range(Cells(row + 1, 2), Cells(row + I_cnt, 1 + V_cnt)).Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    Case 3
        'Set Voltage Difference chart to xlBubble
        With ActiveChart
            .ChartType = xlBubble
            .ChartGroups(1).BubbleScale = 10
            .Axes(xlCategory).MinimumScale = Cells(row + 1, 1).Value
            .Axes(xlCategory).MaximumScale = Cells(row + I_cnt, 1).Value
            .Axes(xlCategory).Select
        End With
        Selection.TickLabelPosition = xlLow
    End Select

    'Modify chart size and color of gridline
    With ActiveSheet.ChartObjects(cat)
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
    
    'Move position of the chart
    ActiveSheet.ChartObjects(cat).Activate
    ActiveChart.Parent.Cut
    Range(pos).Select
    ActiveSheet.Paste
End Sub

Sub Other_pivot_chart()
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim cnt As Integer
    Dim log_name As String, row_name As String, column_name As String, data_name As String
    Dim comment_name As String, comment_reg As String, Eff_reg As String, Diff_reg As String
    Dim col_cnt As Integer
    Dim cat As Integer '1. Voltage; 2. Efficiency; 3. MAIN Voltage Difference

    cnt = 1
    col_cnt = 0

    'Type in the sheet name for analysis
    log_name = Inputbox("Please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
    Do While Sheets(cnt).Name <> log_name
        If log_name = "end" Then
            Exit Do
        End If

        cnt = cnt + 1
        If cnt > Sheets.Count Then
            cnt = 1
            log_name = Inputbox("Please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
        End If
    Loop

    'Type in the row label for the chart
    If  log_name <> "end" Then
        row = Inputbox("In analysis, please type in the Battery Input type for the chart." & vbCrlf & "1 for DC1Voltage (EXT Battery Input)" & vbCrlf & _
        "2 for DC2Voltage (INT Battery Input)" & vbCrlf & "0 to leave")
        Do While row < 0 Or row > 2
            If row = 0 Then
                Exit Do
            End If            
            row = Inputbox("In analysis, please type in the Battery Input type for the chart." & vbCrlf & "1 for DC1Voltage (EXT Battery Input)" & vbCrlf & _
            "2 for DC2Voltage (INT Battery Input)" & vbCrlf & "0 to leave")
        Loop
             
        Select Case row
        Case 1
            row_name = "DC1Voltage"
        Case 2
            row_name = "DC2Voltage"
        End Select
    End If

    'Type in the column label for the chart
    If  log_name <> "end" Then
        col = Inputbox(row_name & " is selected!" & vbCrlf & "In analysis, please type in the load number for the chart."  & vbCrlf & _
        "2 for MAIN O/P (Load2Current)" & vbCrlf & "3 for PRINTER O/P (Load3Current)" & vbCrlf & "4 for 12V O/P (Load4Current)" & vbCrlf & _
        "5 for 24V O/P (Load5Current)" & vbCrlf & "0 to leave")
        Do While col < 0 Or col > 5
            If col = 0 Then
                Exit Do
            End If  
            col = Inputbox(row_name & " is selected!" & vbCrlf & "In analysis, please type in the load number for the chart." & vbCrlf & _
            "2 for MAIN O/P (Load2Current)" & vbCrlf & "3 for PRINTER O/P (Load3Current)" & vbCrlf & "4 for 12V O/P (Load4Current)" & vbCrlf & _
            "5 for 24V O/P (Load5Current)" & vbCrlf & "0 to leave")
        Loop
             
        Select Case col
        Case 2
            column_name = "Load2Current"
        Case 3
            column_name = "Load3Current"
        Case 4
            column_name = "Load4Current"
        Case 5
            column_name = "Load5Current"
        End Select
    End If

    If log_name <> "end" And row <> 0 And col <> 0 Then
        Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType := xlDatabase, SourceData := Sheets(log_name).Range("A1").CurrentRegion.Address)
        Set PT = PTCache.CreatePivotTable (TableDestination := "", TableName:="Other_pivot")

        'Set column in the new pivot chart
        With PT
            'Set [comment] as the filter
            .PivotFields("comment").Orientation = xlPageField
            With .PivotFields("comment")
                For cnt = 1 To .PivotItems.Count 
                    comment_name = .PivotItems(cnt).Name
                    Select Case col
                    Case 2  'MAIN O/P Current
                        data_name = "Load2Voltage"
                        Select Case row_name
                        Case "DC1Voltage"  'MAIN from EXT
                            comment_reg = "MAIN from EXT"
                            Eff_reg = "L2/DC1_Eff"
                        Case "DC2Voltage"  'MAIN from INT
                            comment_reg = "MAIN from INT"
                            Eff_reg = "L2/DC2_Eff"
                        End Select
                    Case 3  'PRINTER O/P Current
                        data_name = "Load3Voltage"
                        Select Case row_name
                        Case "DC1Voltage"  'PRINTER from EXT
                            comment_reg = "PRINTER from EXT"
                            Eff_reg = "L3/DC1_Eff"
                        Case "DC2Voltage"  'PRINTER from INT
                            comment_reg = "PRINTER from INT"
                            Eff_reg = "L3/DC2_Eff"
                        End Select                        
                    Case 4  '12V+ O/P Current
                        data_name = "Load4Voltage"
                        Select Case row_name
                        Case "DC1Voltage"  '12V+ from EXT
                            comment_reg = "12V+ from EXT"
                            Eff_reg = "L4/DC1_Eff"
                        Case "DC2Voltage"  '12V+ from INT
                            comment_reg = "12V+ from INT"
                            Eff_reg = "L4/DC2_Eff"
                        End Select                          
                    Case 5  '24V+ O/P Current
                        data_name = "Load5Voltage"
                        Select Case row_name
                        Case "DC1Voltage"  '24V+ from EXT
                            comment_reg = "24V+ from EXT"
                            Eff_reg = "L5/DC1_Eff"
                        Case "DC2Voltage"  '24V+ from INT
                            comment_reg = "24V+ from INT"
                            Eff_reg = "L5/DC2_Eff"
                        End Select                         
                    End Select

                    If comment_name <> comment_reg Then
                        .PivotItems(comment_name).Visible = False
                    End If
                Next cnt
            End With
            'Set the row label
            .PivotFields(row_name).Orientation = xlColumnField
            'Set the column label
            .PivotFields(column_name).Orientation = xlRowField
            'Set the data in analysis
            With .PivotFields(data_name)
                .Orientation = xlDataField
                .Function = xlSum
            End With
        End With
    
        'Turn on the PivotTable Field List
        Application.CommandBars("PivotTable Field List").Enabled = True

        'Data calculation setting
        With ActiveSheet.PivotTables("Other_pivot")
            .PivotFields(row_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields(column_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields(data_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With

        Do While IsNumeric(Cells(5 + col_cnt, 1)) = True
            col_cnt = col_cnt + 1
        Loop
        
        'Create V/I curve
        Create_Other_Chart row:= col_cnt + 7, I_cnt := col_cnt, cat := 1, pos := "A71"

        'Change parameter of pivot chart to Efficiency
        ActiveSheet.PivotTables("Other_pivot").PivotFields("加總 - " & data_name).Orientation = xlHidden
        ActiveSheet.PivotTables("Other_pivot").AddDataField ActiveSheet.PivotTables("Other_pivot").PivotFields(Eff_reg), "加總 - " & Eff_reg, xlSum

        'Create Efficiency curve
        Create_Other_Chart row := col_cnt * 2 + 7 + 2, I_cnt := col_cnt, cat := 2, pos := "AK71"

        If col = 2 Then
            'Change parameter of pivot chart to Voltage Difference
            ActiveSheet.PivotTables("Other_pivot").PivotFields("加總 - " & Eff_reg).Orientation = xlHidden
            ActiveSheet.PivotTables("Other_pivot").AddDataField ActiveSheet.PivotTables("Other_pivot").PivotFields("MAIN_volt_Diff"), "加總 - MAIN_volt_Diff", xlSum

            'Create Voltage Difference curve
            Create_Other_Chart row := col_cnt * 3 + 7 + 2 * 2, I_cnt := col_cnt, cat := 3, pos := "BG71"
        End If

        ActiveSheet.Name = comment_reg
        ActiveWindow.Zoom = 40
        Msgbox "Test result has been generated successfully."
    Else
        Msgbox "Abort the pivot chart generating!"
    End If
End Sub