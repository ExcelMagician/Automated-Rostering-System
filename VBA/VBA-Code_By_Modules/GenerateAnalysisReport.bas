Attribute VB_Name = "GenerateAnalysisReport"
' Reusable function to generate side-by-side shift analysis blocks
Sub GenerateShiftAnalysisBlock(wsAnalysis As Worksheet, rosterSheet As Worksheet, _
                                personnelSheetName As String, tableName As String, _
                                slotTitle As String, rosterCol As Long, startCol As Long)

    Dim wsPersonnel As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range, dutyCounterList As Range
    Dim dict As Object
    Dim empName As String
    Dim cell As Range, cellValue As String, currStaff As String
    Dim rowOffset As Long: rowOffset = 4
    Dim lastRow As Long, nextRow As Long
    Dim tableRange As Range, analysisTable As ListObject
    Dim i As Long, tableWidth As Long: tableWidth = 5

    Set wsPersonnel = Sheets(personnelSheetName)
    Set tbl = wsPersonnel.ListObjects(tableName)
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Add small section title
    With wsAnalysis.Cells(3, startCol).Resize(1, 3)
        .Merge
        .Value = slotTitle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With

    ' Add header
    With wsAnalysis
        .Cells(rowOffset, startCol).Value = "Name"
        .Cells(rowOffset, startCol + 1).Value = "System Counter"
        .Cells(rowOffset, startCol + 2).Value = "Actual Counter"
        .Cells(rowOffset, startCol + 3).Value = "Difference"
        .Cells(rowOffset, startCol + 4).Value = "% Difference"
    End With

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.Count
        empName = UCase(Trim(nameList.Cells(i, 1).Value))
        wsAnalysis.Cells(rowOffset + i, startCol).Value = empName
        wsAnalysis.Cells(rowOffset + i, startCol + 1).Value = dutyCounterList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    For i = 6 To 186
        Set cell = rosterSheet.Cells(i, rosterCol)
        cellValue = cell.Value

        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
        Else
            currStaff = UCase(Trim(cellValue))
        End If

        If Len(currStaff) > 0 And currStaff <> "CLOSED" Then
            If dict.exists(currStaff) Then
                dict(currStaff) = dict(currStaff) + 1
            Else
                nextRow = wsAnalysis.Cells(wsAnalysis.Rows.Count, startCol).End(xlUp).row + 1
                wsAnalysis.Cells(nextRow, startCol).Value = currStaff
                wsAnalysis.Cells(nextRow, startCol + 1).Value = 0
                wsAnalysis.Cells(nextRow, startCol + 2).Value = 1
                wsAnalysis.Cells(nextRow, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
                wsAnalysis.Cells(nextRow, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,"""",RC[-1]/RC[-3]*100)"
                dict(currStaff) = 1
                wsAnalysis.Range(wsAnalysis.Cells(nextRow, startCol), wsAnalysis.Cells(nextRow, startCol + 4)).Interior.Color = RGB(255, 255, 153)
            End If
        End If
    Next i

    For i = rowOffset + 1 To wsAnalysis.Cells(wsAnalysis.Rows.Count, startCol).End(xlUp).row
        empName = UCase(Trim(wsAnalysis.Cells(i, startCol).Value))
        If dict.exists(empName) Then
            wsAnalysis.Cells(i, startCol + 2).Value = dict(empName)
            wsAnalysis.Cells(i, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
            wsAnalysis.Cells(i, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3]*100)"
        End If
    Next i

    ' Format as Table
    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.Count, startCol).End(xlUp).row
    Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), wsAnalysis.Cells(lastRow, startCol + tableWidth - 1))
    Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    analysisTable.Name = Replace(slotTitle, " ", "") & "Table"
    analysisTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"
End Sub

' Master sub to create the whole Analysis Report layout
Sub MasterGenerateAllAnalyses()
    Dim wsAnalysis As Worksheet
    Dim rosterSheet As Worksheet
    Dim latestRosterName As String
    Dim newestDate As Date
    Dim sht As Worksheet

    newestDate = 0
    For Each sht In ThisWorkbook.Sheets
        If sht.Name Like "ActualRoster_*" Then
            Dim dtPart As String
            dtPart = Replace(Mid(sht.Name, 14), "_", " ")
            On Error Resume Next
            Dim parsedDate As Date
            parsedDate = CDate(Left(dtPart, 4) & "/" & Mid(dtPart, 5, 2) & "/" & Mid(dtPart, 7, 2) & " " & Mid(dtPart, 10, 2) & ":" & Mid(dtPart, 12, 2))
            If Err.Number = 0 Then
                If parsedDate > newestDate Then
                    newestDate = parsedDate
                    latestRosterName = sht.Name
                End If
            End If
            On Error GoTo 0
        End If
    Next sht

    If latestRosterName = "" Then
        MsgBox "No ActualRoster_* sheet found.", vbExclamation
        Exit Sub
    End If

    Set rosterSheet = Sheets(latestRosterName)
    ' Prompt user to click on any cell in the target ActualRoster_* sheet
    On Error Resume Next
    Set userRange = Application.InputBox( _
        Prompt:="Please choose one 'ActualRoster' sheet to analyse." & vbCrLf & _
                "After that, click on any cell on the selected 'ActualRoster' sheet." & vbCrLf & _
                "The sheet name must start with 'ActualRoster_'", _
        title:="Select Actual Roster Sheet", _
        Type:=8)
    On Error GoTo 0

    If userRange Is Nothing Then Exit Sub ' User cancelled

    Set selectedSheet = userRange.Worksheet
    If selectedSheet.Name Like "ActualRoster_*" = False Then
        MsgBox "Invalid selection. Please choose a sheet that starts with 'ActualRoster_'.", vbExclamation
        Exit Sub
    End If
    Set rosterSheet = selectedSheet

    ' Create clean MorningAnalysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.Count))
    wsAnalysis.Name = "AnalysisReport"

    ' Big title
    With wsAnalysis.Range("A1:Z1")
        .Merge
        .Value = "Analysis Report"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(255, 199, 206)
        .HorizontalAlignment = xlCenter
    End With

    ' Generate all 5 analyses side by side
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Loan Mail Box PersonnelList", "LoanMailBoxMainList", "Loan Mail Box Slot Analysis", LMB_COL, 1
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Morning PersonnelList", "MorningMainList", "Morning Slot Analysis", MOR_COL, 7
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Afternoon PersonnelList", "AfternoonMainList", "Afternoon Slot Analysis", AFT_COL, 13
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "AOH PersonnelList", "AOHMainList", "AOH Slot Analysis", AOH_COL, 19
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Sat AOH PersonnelList", "SatAOHMainList", "Sat AOH Slot Analysis", SAT_AOH_COL1, 25
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Sat AOH PersonnelList", "SatAOHMainList", _
    "Sat AOH Slot Analysis", SAT_AOH_COL1, 25, SAT_AOH_COL2
    GenerateTotalSummaryTable wsAnalysis
    
    With wsAnalysis.Cells
        .Locked = True
    End With
    
    wsAnalysis.Protect password:="nuslib2017@52", _
                        AllowSorting:=True, _
                        AllowFiltering:=True, _
                        AllowFormattingCells:=True
                        
    

    MsgBox "All shift analyses completed for '" & rosterSheet.Name & "'!", vbInformation
End Sub


    MsgBox "All shift analyses completed successfully!"
End Sub


