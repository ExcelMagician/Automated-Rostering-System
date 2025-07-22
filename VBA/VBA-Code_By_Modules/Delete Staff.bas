Attribute VB_Name = "DeleteStaff"
Sub DeleteStaff()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selectedCell As Range
    Dim rowToDelete As ListRow
    Dim dutyType As String
    Dim specificDaysTbl As ListObject
    Dim availIndex As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Determine the duty type based on the active sheet
    Select Case UCase(ws.Name)
        Case UCase("Loan Mail Box PersonnelList")
            dutyType = "LOANMAILBOX"
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case UCase("Morning PersonnelList")
            dutyType = "MORNING"
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case UCase("Afternoon PersonnelList")
            dutyType = "AFTERNOON"
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case UCase("AOH PersonnelList")
            dutyType = "AOH"
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case UCase("Sat AOH PersonnelList")
            dutyType = "SAT_AOH"
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

    ' Unprotect the worksheet (unlock)
    ws.Unprotect
    
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
        If selectedCell.Column <> tbl.Range.Cells(1, nameIndex).Column Then
            MsgBox "Please select a cell in the 'Name' column to delete the staff.", vbExclamation
            GoTo ReprotectAndExit
        End If
        
        ' Find the row in the table
        Dim rowIndex As Long
        rowIndex = selectedCell.row - tbl.Range.row
        If rowIndex > 0 And rowIndex <= tbl.ListRows.Count Then
            Set rowToDelete = tbl.ListRows(rowIndex)
            ' Get the Availability Type index
            availIndex = GetColumnIndex(tbl, "Availability Type")
            If availIndex = -1 Then
                MsgBox "Column 'Availability Type' not found in '" & tbl.Name & "'.", vbExclamation
                GoTo ReprotectAndExit
            End If
            
            ' Clear filters if any
            If tbl.ShowAutoFilter Then
                tbl.AutoFilter.ShowAllData
            End If
            If Not specificDaysTbl Is Nothing And specificDaysTbl.ShowAutoFilter Then
                specificDaysTbl.AutoFilter.ShowAllData
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
                ' Clear filters on specific days table
                If specificDaysTbl.ShowAutoFilter Then
                    specificDaysTbl.AutoFilter.ShowAllData
                End If
                ' Find and delete the corresponding row in SpecificDaysWorkingStaff
                For Each sdRow In specificDaysTbl.ListRows
                    If UCase(Trim(sdRow.Range.Cells(1, sdNameIndex).Value)) = UCase(Trim(selectedCell.Value)) Then
                        sdRow.Delete
                        Exit For
                    End If
                Next sdRow
            End If
            
            ' Delete the row from the main table
            rowToDelete.Delete
            CalculateMaxDuties.CalculateMaxDuties dutyType
            MsgBox "Staff deleted and Max Duties recalculated successfully for " & dutyType & ".", vbInformation
            GoTo ReprotectAndExit
        Else
            MsgBox "Selected cell is not within a valid table row.", vbExclamation
            GoTo ReprotectAndExit
        End If
    Else
        MsgBox "Please select a cell within the main table to delete a staff.", vbExclamation
        GoTo ReprotectAndExit
    End If
    Exit Sub

ReprotectAndExit:
    ' Reprotect the worksheet and lock table ranges
    With ws
        If Not tbl Is Nothing Then
            .ListObjects(tbl.Name).Range.Locked = True
        End If
        If Not specificDaysTbl Is Nothing Then
            .ListObjects(specificDaysTbl.Name).Range.Locked = True
        End If
        .Range("D5:D9").Locked = False ' keep data entry remains unlocked
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                 AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
    End With
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Duty Type: " & dutyType & vbCrLf & _
           "Worksheet: " & IIf(ws Is Nothing, "Not Set", ws.Name), vbCritical
    GoTo ReprotectAndExit
End Sub

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

