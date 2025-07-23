Attribute VB_Name = "rEPORT1"
Sub GenerateMorningShiftAnalysis()
    Dim wsPersonnel As Worksheet
    Dim wsAnalysis As Worksheet
    Dim wsRosterCopy As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range
    Dim dutyCounterList As Range
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim empName As String
    Dim latestRosterName As String
    Dim sht As Worksheet
    Dim newestDate As Date
    Dim match As Object
    Set match = CreateObject("Scripting.Dictionary")

    ' Find latest duplicated actual roster sheet
    newestDate = 0
    For Each sht In ThisWorkbook.Sheets
        If sht.Name Like "ActualRoster_*" Then
            Dim dtPart As String
            dtPart = Replace(Mid(sht.Name, 14), "_", " ") ' Extract "yyyymmdd hhnn"
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
    
    Debug.Print "latest actual roster:" & latestRosterName
    
    If latestRosterName = "" Then
        MsgBox "No ActualRoster_* sheet found.", vbExclamation
        Exit Sub
    End If

    Set wsRosterCopy = Sheets(latestRosterName)
    Set wsPersonnel = Sheets("Morning PersonnelList")

    ' Create or clear analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.Count))
    wsAnalysis.Name = "MorningAnalysis"

    ' Get the personnel table
    Set tbl = wsPersonnel.ListObjects("MorningMainList")
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Header row
    With wsAnalysis
        .Range("A1").Value = "Name"
        .Range("B1").Value = "System Counter"
        .Range("C1").Value = "Actual Counter"
        .Range("D1").Value = "Difference"
    End With

    ' Copy names and system counters
    For i = 1 To nameList.Rows.Count
        wsAnalysis.Cells(i + 1, 1).Value = nameList.Cells(i, 1).Value
        wsAnalysis.Cells(i + 1, 2).Value = dutyCounterList.Cells(i, 1).Value
    Next i

    ' Create dictionary to count actual appearances
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.Count
        empName = nameList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    ' Count appearances (ignoring strikethrough and multiple lines)
    'lastRow = wsRoster.Cells(wsRoster.Rows.Count, MOR_COL).End(xlUp).row
    Dim cell As Range, firstLine As String
    For i = START_ROW To 186 'LAST_ROW_ROSTER
        Set cell = wsRosterCopy.Cells(i, MOR_COL)
        cellValue = cell.Value
        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
        Else
            currStaff = UCase(Trim(cellValue))
        End If
        
        If dict.exists(currStaff) Then
            dict(currStaff) = dict(currStaff) + 1
            Debug.Print currStaff; ": " & dict(currStaff)
        End If
    Next i

    ' Write results
    For i = 2 To nameList.Rows.Count + 1
        empName = UCase(Trim(wsAnalysis.Cells(i, 1).Value))
        wsAnalysis.Cells(i, 3).Value = dict(empName) ' Actual Counter
        wsAnalysis.Cells(i, 4).FormulaR1C1 = "=RC[-2]-RC[-1]" ' System - Actual
    Next i

    MsgBox "Morning shift analysis generated using '" & latestRosterName & "'.", vbInformation
End Sub

