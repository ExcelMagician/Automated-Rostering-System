Attribute VB_Name = "InsertStaff"
Sub InsertStaff(dutyType As String)
    'On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim staffName As String, dept As String
    Dim availType As String, workDays As String, percentage As String
    Dim checkRow As Long
    Dim specificDaysTbl As ListObject
    Dim specificRow As ListRow
    
    ' Set worksheet and tables based on dutyType
    Select Case UCase(dutyType)
        Case "LOANMAILBOX"
            Set ws = ThisWorkbook.Sheets("Loan Mail Box PersonnelList")
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("Afternoon PersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
            ' No specificDaysTbl for Sat AOH
        Case Else
            MsgBox "Invalid duty type. Use 'LoanMailBox', 'Morning', 'Afternoon', 'AOH', or 'Sat_AOH'.", vbExclamation
            Exit Sub
    End Select
    
    If ws Is Nothing Then
        MsgBox "Worksheet for " & dutyType & " not found.", vbExclamation
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "Table 'MainList' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    If specificDaysTbl Is Nothing And UCase(Trim(ws.Range("D7").Value)) = "SPECIFIC DAYS" And UCase(dutyType) <> "SAT_AOH" Then
        MsgBox "Table 'SpecificDaysWorkingStaff' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    staffName = UCase(Trim(ws.Range("D5").Value)) ' Name
    dept = Trim(ws.Range("D6").Value)             ' Department
    availType = UCase(Trim(ws.Range("D7").Value)) ' Availability Type
    workDays = Trim(ws.Range("D8").Value)         ' Working Days
    percentage = Trim(ws.Range("D9").Value)       ' Duties Percentage


    ' Auto-fill logic based on Availability Type (skip for Sat_AOH)
    If availType = "ALL DAYS" Then
            percentage = "100"
            workDays = ""
        ElseIf availType = "SPECIFIC DAYS" Then
            If workDays = "" Then
                MsgBox "Please enter Working Days for Specific Days availability.", vbExclamation
                Exit Sub
            End If
    End If
    
    ' Validate percentage
    If percentage = "" Or Not IsNumeric(percentage) Or Val(percentage) <= 0 Or Val(percentage) > 100 Then
        MsgBox "Please enter a valid Duties Percentage (1-100).", vbExclamation
        Exit Sub
    End If

    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in Name and Department.", vbExclamation
        Exit Sub
    End If
    
    ' Check for duplicate names
    For checkRow = 1 To tbl.ListRows.Count
        If UCase(Trim(tbl.ListRows(checkRow).Range.Cells(1, GetColumnIndex(tbl, "Name")).Value)) = staffName Then
            MsgBox "This staff name already exists.", vbExclamation
            Exit Sub
        End If
    Next checkRow
    
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    With newRow.Range
        Dim nameIndex As Long, deptIndex As Long, availIndex As Long
        Dim percIndex As Long, maxIndex As Long, counterIndex As Long
        
        nameIndex = GetColumnIndex(tbl, "Name")
        deptIndex = GetColumnIndex(tbl, "Department")
        availIndex = GetColumnIndex(tbl, "Availability Type")
        percIndex = GetColumnIndex(tbl, "Duties Percentage (%)")
        maxIndex = GetColumnIndex(tbl, "Max Duties")
        counterIndex = GetColumnIndex(tbl, "Duties Counter")
        
        If nameIndex = -1 Or deptIndex = -1 Or counterIndex = -1 Or availIndex = -1 Or percIndex = -1 Or maxIndex = -1 Then
            MsgBox "Required columns 'Name', 'Department', 'Availability Type', 'Duties Percentage', 'Max Duties', or 'Duties Counter' not found in '" & tbl.Name & "'.", vbExclamation
            newRow.Delete
            Exit Sub
        End If
        
        .Cells(1, nameIndex).Value = staffName
        .Cells(1, deptIndex).Value = dept
        .Cells(1, availIndex).Value = availType
        .Cells(1, percIndex).Value = Val(percentage)
        .Cells(1, counterIndex).Value = 0
        ' Max Duties will be calculated later
    End With

    ' Handle specific days workers table
    If availType = "SPECIFIC DAYS" Then
        Set specificRow = specificDaysTbl.ListRows.Add(AlwaysInsert:=True)
        With specificRow.Range
            Dim specNameIndex As Long, specDaysIndex As Long
            specNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            specDaysIndex = GetColumnIndex(specificDaysTbl, "Working Days")
            If specNameIndex = -1 Or specDaysIndex = -1 Then
                MsgBox "Columns 'Name' or 'Working Days' not found in '" & specificDaysTbl.Name & "'.", vbExclamation
                specificRow.Delete
                newRow.Delete
                Exit Sub
            End If
            .Cells(1, specNameIndex).Value = staffName
            .Cells(1, specDaysIndex).Value = workDays
        End With
    End If

    CalculateMaxDuties.CalculateMaxDuties dutyType

    ' Clear input of data entry
    ws.Range("D5:D9").ClearContents

    MsgBox "Staff added and Max Duties calculated successfully for " & dutyType & "!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Duty Type: " & dutyType & vbCrLf & _
           "Worksheet: " & IIf(ws Is Nothing, "Not Set", ws.Name), vbCritical
    If Not newRow Is Nothing Then newRow.Delete
    If Not specificRow Is Nothing Then specificRow.Delete
    Exit Sub
End Sub

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

' Wrapper subroutines for different shifts
Sub RunInsertStaffLMB()
    InsertStaff "LoanMailBox"
End Sub

Sub RunInsertStaffMorning()
    InsertStaff "Morning"
End Sub

Sub RunInsertStaffAfternoon()
    InsertStaff "Afternoon"
End Sub

Sub RunInsertStaffAOH()
    InsertStaff "AOH"
End Sub

Sub RunInsertStaffSatAOH()
    InsertStaff "Sat_AOH"
End Sub
