Attribute VB_Name = "DuplicateActualRoster"
Sub DuplicateActualRoster()
    Dim srcSheet As Worksheet
    Dim copySheet As Worksheet
    Dim sheetName As String
    
    Set srcSheet = Sheets("Roster")
    sheetName = "ActualRoster_" & Format(Now, "yyyymmdd_hhnn")
    
    srcSheet.Copy After:=Sheets(Sheets.Count)
    Set copySheet = ActiveSheet
    copySheet.Name = sheetName
    
    With srcSheet.UsedRange
        .Copy Destination:=copySheet.Range("A1")
    End With

    copySheet.Protect password:="nuslib2017@52", _
                        AllowSorting:=True, _
                        AllowFiltering:=True, _
                        AllowFormattingCells:=True
                        
    
    srcSheet.Activate
    
End Sub

