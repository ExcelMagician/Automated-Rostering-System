Attribute VB_Name = "DeleteStaff"
Sub DeleteStaff()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selectedCell As Range
    Dim rowToDelete As ListRow
    Dim dutyType As String
    Dim password As String
    Dim specificDaysTbl As ListObject
    Dim availIndex As Long
    
    ' Protection password
    password = "nuslib2025"
    
    ' Authorization check
    Dim userPassword As String
    userPassword = InputBox("Enter the password to remove the staff:", "Authorization Required")
    If userPassword <> password Then
        MsgBox "Unauthorized access.", vbExclamation
        Exit Sub
    End If
    
    ' Duty type based on the active sheet
    Select Case UCase(ActiveSheet.Name)
        Case UCase("Loan Mail Box PersonnelList")
            dutyType = "LOANMAILBOX"
            Set ws = ThisWorkbook.Sheets("Loan Mail Box PersonnelList")
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case UCase("Morning PersonnelList")
            dutyType = "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case UCase("Afternoon PersonnelList")
            dutyType = "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("Afternoon PersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case UCase("AOH PersonnelList")
            dutyType = "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case UCase("Sat AOH PersonnelList")
            dutyType = "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
            ' No specificDaysTbl for Sat AOH
        Case Else
            MsgBox "This sheet is not a personnel list. Please select a valid personnel list sheet.", vbExclamation
            Exit Sub
    End Select
    
    ' Validate worksheet and table
    If ws Is Nothing Then
        MsgBox "Worksheet for " & dutyType & " not found.", vbExclamation
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "Table 'MainList' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Get the selected cell
    Set selectedCell = ActiveCell
    If Not Intersect(selectedCell, tbl.Range) Is Nothing Then
        ' Check if the selected cell is in the "Name" column
        Dim nameIndex As Long
        nameIndex = GetColumnIndex(tbl, "Name")
        If nameIndex = -1 Then
            MsgBox "Column 'Name' not found in '" & tbl.Name & "'.", vbExclamation
            GoTo ReprotectAndExit
        End If
        If selectedCell.Column = tbl.Range.Cells(1, nameIndex).Column Then
            ' Find the row in the table
            Dim rowIndex As Long
            rowIndex = selectedCell.row - tbl.Range.row
            If rowIndex >= 1 And rowIndex <= tbl.ListRows.Count Then
                Set rowToDelete = tbl.ListRows(rowIndex)
                ' Get the Availability Type index
                availIndex = GetColumnIndex(tbl, "Availability Type")
                If availIndex = -1 Then
                    MsgBox "Column 'Availability Type' not found in '" & tbl.Name & "'.", vbExclamation
                    GoTo ReprotectAndExit
                End If
                ' Check if the staff is a Specific Days worker
                If UCase(Trim(rowToDelete.Range.Cells(1, availIndex).Value)) = "SPECIFIC DAYS" And Not specificDaysTbl Is Nothing Then
                    Dim sdRow As ListRow
                    Dim sdNameIndex As Long
                    sdNameIndex = GetColumnIndex(specificDaysTbl, "Name")
                    If sdNameIndex = -1 Then
                        MsgBox "Column 'Name' not found in '" & specificDaysTbl.Name & "'.", vbExclamation
                        GoTo ReprotectAndExit
                    End If
                    ' Find and delete the corresponding row in SpecificDaysWorkingStaff
                    For Each sdRow In specificDaysTbl.ListRows
                        If UCase(Trim(sdRow.Range.Cells(1, sdNameIndex).Value)) = UCase(Trim(selectedCell.Value)) Then
                            sdRow.Delete
                            Exit For
                        End If
                    Next sdRow
                End If
                rowToDelete.Delete
                CalculateMaxDuties.CalculateMaxDuties dutyType
                MsgBox "Staff deleted and Max Duties recalculated successfully for " & dutyType & ".", vbInformation
            Else
                MsgBox "Selected cell is not within a valid table row.", vbExclamation
                GoTo ReprotectAndExit
            End If
        Else
            MsgBox "Please select a cell in the 'Name' column to delete the staff.", vbExclamation
            GoTo ReprotectAndExit
        End If
    Else
        MsgBox "Please select a cell within the table to delete a staff.", vbExclamation
        GoTo ReprotectAndExit
    End If
    Exit Sub

ReprotectAndExit:
    ' Reprotect the worksheet
    ws.Protect password, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Duty Type: " & dutyType & vbCrLf & _
           "Worksheet: " & IIf(ws Is Nothing, "Not Set", ws.Name), vbCritical
    ws.Protect password, DrawingObjects:=True, Contents:=True, Scenarios:=True ' Reprotect on error
    Exit Sub
End Sub

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

