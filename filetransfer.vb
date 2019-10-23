Sub txtToxls()
    On Error GoTo ErrorHandler 
    Dim FileName As String, PathName  As String
    Dim cnt As Integer, i As Integer

    PathName = Application.ActiveWorkbook.Path
    OutputFileNum = FreeFile
    cnt = 1

    FileName = Inputbox("Please type in the .txt file name to transfer to .xls file.")

    Open PathName & "\" & FileName & ".txt" For Input As #OutputFileNum
    Msgbox "Processing..." & vbCrlf & PathName & "\" & FileName & ".txt"

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

    Msgbox FileName & ".txt has already been transfered!"

    Exit Sub

    'Error resolve for the file not found in the path
ErrorHandler:
        MsgBox "Error " & Err.Number & "：" & Err.Description & vbCrlf & _
        "There is no " & FileName & ".txt in the path."
        FileName = Inputbox("Please type in the .txt file name to transfer to .xls file.")
    Resume  
End Sub