Sub ReleaseAllData()
    MsgBox "Release All Data!"
    ActiveSheet.ShowAllData
End Sub

Sub SelectThermalPivot_ACDC_Data()
    ActiveSheet.PivotTables(1).PivotSelect "Load1Current" ,xlDataOnly
    Selection.Copy
End Sub

Sub SelectThermalPivot_Column()
    ActiveSheet.PivotTables(1).PivotSelect "Minutes" ,xlLabelOnly
    Selection.Copy
End Sub

Sub acdc_pivot()
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim cnt As Integer, log_name As String
    Dim data_name As String
    Dim comment_name As String

    cnt = 1

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

    'Type in the data for the chart
    If  log_name <> "end" Then
        data_name = Inputbox("In AC/DC analysis, please type in the analysis data for the chart. e.g. Efficiency: L1/AC_Eff" & vbCrlf & "Or, type in 'end' to leave.")
        cnt = 1
        Do While Sheets(log_name).Cells(1, cnt).Value <> data_name
            If data_name = "end" Then
                Exit Do
            End If

            cnt = cnt + 1
            If Sheets(log_name).Cells(1, cnt).Value = "" Then
                data_name = Inputbox("In AC/DC analysis, please type in the analysis data for the chart. e.g. Efficiency: L1/AC_Eff" & vbCrlf & "Or, type in 'end' to leave.")
                cnt = 1
            End If
        Loop
    End If

    If  log_name <> "end" And data_name <> "end" Then
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
            'Set the data in analysis
            With .PivotFields(data_name)
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
            .PivotFields(data_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
    Else
        Msgbox "Abort the pivot chart generating!"
    End If
End Sub

Sub Other_pivot()
    On Error GoTo ErrorHandler 
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim cnt As Integer, row As Integer, col As Integer
    Dim log_name As String, row_name As String, column_name As String, data_name As String
    Dim comment_name As String

    cnt = 1

    'Type in the sheet name under analysis
    log_name = Inputbox("In analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
    Do While Sheets(cnt).Name <> log_name
        If log_name = "end" Then
            Exit Do
        End If

        cnt = cnt + 1
        If cnt > Sheets.Count Then
            cnt = 1
            log_name = Inputbox("In analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
        End If
    Loop

    'Type in the row label for the chart
    If  log_name <> "end" Then
        row = Inputbox("In analysis, please type in the Battery Input type for the chart." & vbCrlf & "1 for DC1Voltage (EXT Battery Input)" & vbCrlf & _
        "2 for DC2Voltage (INT Battery Input)")
        Do While row < 1 Or row > 2
            row = Inputbox("In analysis, please type in the Battery Input type for the chart." & vbCrlf & "1 for DC1Voltage (EXT Battery Input)" & vbCrlf & _
            "2 for DC2Voltage (INT Battery Input)")
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
        col = Inputbox(row_name & " is selected!" & vbCrlf & "In analysis, please type in the load number for the chart." & vbCrlf & "Number range is from 2 to 5.")
        Do While col < 2 Or col > 5
            col = Inputbox(row_name & " is selected!" & vbCrlf & "In analysis, please type in the load number for the chart." & vbCrlf & "Number range is from 2 to 5.")
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

    'Type in the data for the chart
    If  log_name <> "end" Then
        data_name = Inputbox(row_name & " & " & column_name & " are selected!" & vbCrlf & "In analysis, please type in the analysis data for the chart. " & _
        "e.g. Efficiency: L2/DC2_Eff" & vbCrlf & "Or, type in 'end' to leave.")

        cnt = 1
        Do While Sheets(log_name).Cells(1, cnt).Value <> data_name
            If data_name = "end" Then
                Exit Do
            End If

            cnt = cnt + 1
            If Sheets(log_name).Cells(1, cnt).Value = "" Then
                data_name = Inputbox(row_name & " & " & column_name & " are selected!" & vbCrlf & "In analysis, please type in the analysis data for the chart. " & _
                "e.g. Efficiency: L2/DC2_Eff" & vbCrlf & "Or, type in 'end' to leave.")

                cnt = 1
            End If
        Loop
    End If

    If  log_name <> "end" And row_name <> "end" And data_name <> "end" Then
        Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType := xlDatabase, SourceData := Sheets(log_name).Range("A1").CurrentRegion.Address)
        Set PT = PTCache.CreatePivotTable (TableDestination := "", TableName := "pivot")
        
        'Set column in the new pivot chart
        With PT
            'Set [comment] as the filter
            .PivotFields("comment").Orientation = xlPageField
            With .PivotFields("comment")
                For cnt = 1 To .PivotItems.Count 
                    comment_name = .PivotItems(cnt).Name
                    Select Case col
                    Case 2  'MAIN O/P Current
                        Select Case row_name
                        Case "DC1Voltage"  'MAIN from EXT
                            If comment_name <> "MAIN from EXT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        Case "DC2Voltage"  'MAIN from INT
                            If comment_name <> "MAIN from INT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        End Select
                    Case 3  'PRINTER O/P Current
                        Select Case row_name
                        Case "DC1Voltage"  'PRINTER from EXT
                            If comment_name <> "PRINTER from EXT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        Case "DC2Voltage"  'PRINTER from INT
                            If comment_name <> "PRINTER from INT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        End Select                        
                    Case 4  '12V+ O/P Current
                        Select Case row_name
                        Case "DC1Voltage"  '12V+ from EXT
                            If comment_name <> "12V+ from EXT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        Case "DC2Voltage"  '12V+ from INT
                            If comment_name <> "12V+ from INT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        End Select                          
                    Case 5  '24V+ O/P Current
                        Select Case row_name
                        Case "DC1Voltage"  '24V+ from EXT
                            If comment_name <> "24V+ from EXT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        Case "DC2Voltage"  '24V+ from INT
                            If comment_name <> "24V+ from INT" Then
                                .PivotItems(comment_name).Visible = False
                            End If
                        End Select                         
                    End Select
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
        With ActiveSheet.PivotTables("pivot")
            .PivotFields(row_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields(column_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields(data_name).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
    Else
        Msgbox "Abort the pivot chart generating!"
    End If

    Exit Sub

'Error resolve for the comment not found in the path
ErrorHandler:
        MsgBox "Error " & Err.Number & ":" & Err.Description & vbCrlf & _
        "There is no comment (" & column_name & " from " & row_name & ") in the log!"
        Exit Sub
    Resume  
End Sub

Sub acdc_thermal_pivot()
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim cnt As Integer, log_name As String
    Dim comment_name As String, data_name As String, load_current As String

    cnt = 1

    'Type in the sheet name under analysis
    log_name = Inputbox("In AC/DC THERMAL analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
    Do While Sheets(cnt).Name <> log_name
        If log_name = "end" Then
            Exit Do
        End If

        cnt = cnt + 1
        If cnt > Sheets.Count Then
            cnt = 1
            log_name = Inputbox("In AC/DC THERMAL analysis, please type in the sheet name for creating pivot chart." & vbCrlf & "Or, type in 'end' to leave.")
        End If
    Loop

    If  log_name <> "end" And data_name <> "end" Then
        Set PTCache = ThisWorkbook.PivotCaches.Add _
        (SourceType := xlDatabase, SourceData := Sheets(log_name).Range("A1").CurrentRegion.Address)
        Set PT = PTCache.CreatePivotTable (TableDestination := "", TableName:="acdc_thermal_pivot")

        MsgBox "Begin to generate the pivot chart. Please wait!"

        'Set column in the new pivot chart
        With PT
            'Set [comment] as the filter
            .PivotFields("comment").Orientation = xlPageField
            With .PivotFields("comment")
                For cnt = 1 To .PivotItems.Count 
                    comment_name = .PivotItems(cnt).Name
                    If comment_name <> "AC/DC heatsink temperature measurement" Then
                        .PivotItems(comment_name).Visible = False
                    End If
                Next cnt
            End With

            'Set [Load1Current] as the row label
            .PivotFields("Load1Current").Orientation = xlColumnField
            With .PivotFields("Load1Current")
                'Filter Load1Current = 1 to increase the processing efficiency
                For cnt = 1 To .PivotItems.Count 
                    load_current = .PivotItems(cnt).Name
                    If load_current <> "1" Then
                        .PivotItems(load_current).Visible = False
                    End If
                Next cnt
            End With

            'Set [Minutes] as the column label
            .PivotFields("Minutes").Orientation = xlRowField
            'Set the thermal data in analysis
            .PivotFields("temperature0").Orientation = xlDataField
            .PivotFields("temperature1").Orientation = xlDataField
            .PivotFields("temperature2").Orientation = xlDataField
            .PivotFields("temperature3").Orientation = xlDataField
            .PivotFields("temperature4").Orientation = xlDataField
            .PivotFields("temperature5").Orientation = xlDataField
            .PivotFields("temperature6").Orientation = xlDataField
            .PivotFields("temperature7").Orientation = xlDataField
        End With

        'Turn on the PivotTable Field List
        Application.CommandBars("PivotTable Field List").Enabled = True

        'Data calculation setting
        With ActiveSheet.PivotTables("acdc_thermal_pivot")
            .PivotFields("comment").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("Load1Current").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("Minutes").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature0").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature4").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature5").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature6").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .PivotFields("temperature7").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End With
        
        MsgBox "Complete the pivot chart."

    Else
        Msgbox "Abort the pivot chart generating!"
    End If
End Sub
