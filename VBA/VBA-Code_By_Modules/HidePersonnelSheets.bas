Attribute VB_Name = "HidePersonnelSheets"
Sub HidePersonnelSheetsWithPassword()
    Dim ws As Worksheet
    Dim password As String
    
    ' Set the password
    password = "rostering2025"
    
    ' Loop through all worksheets and hide personnel lists
    For Each ws In ThisWorkbook.Sheets
        Select Case UCase(ws.Name)
            Case UCase("AOH PersonnelList"), UCase("Sat AOH PersonnelList"), _
                 UCase("Loan Mail Box PersonnelList"), UCase("Morning PersonnelList"), _
                 UCase("Afternoon PersonnelList")
                ' Protect the entire sheet
                ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True
                ' Set to "Very Hidden" (not visible in UI, only via VBA)
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws
    
    MsgBox "Personnel list sheets have been hidden.", vbInformation
End Sub
