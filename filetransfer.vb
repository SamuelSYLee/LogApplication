Sub txtToxls()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim cnt As Integer, i As Integer

    'Delete all existing data
    Cells.Select
    Selection.Delete

    'Set Application.FileDialog GUI
    With fd
        .AllowMultiSelect = False
        .Title = "Please choose the log file for analysis."
        .InitialFileName = Application.ActiveWorkbook.Path
        .Filters.Clear
        .Filters.Add ".txt file", "*.txt"
    End With
    OutputFileNum = FreeFile
    cnt = 1

    If fd.Show = -1 Then
        MsgBox "Processing... " & vbCrlf & fd.SelectedItems(1)
        
        'Transferring .txt file to .xls file
        Open fd.SelectedItems(1) For Input As #OutputFileNum
        Do Until EOF(OutputFileNum)
            Line Input #OutputFileNum, LineFromFile
            Cells(cnt, 1).Value = LineFromFile
            cnt = cnt + 1
        Loop
        Close #OutputFileNum

        'Delete the .txt file description
        For i = 1 To 3 Step 1
            Rows(1).Delete
        Next

        'Select all data
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        'Seperate each data by semicolon
        Selection.TextToColumns Semicolon := True

        Msgbox fd.SelectedItems(1) & vbCrlf & "has already been transfered!"
    Else
        MsgBox "There is no .txt file under selection!"
    End If
End Sub
