Attribute VB_Name = "ResetAllCounters"
Sub ResetAllCounters()
Attribute ResetAllCounters.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ResetAllCounters Macro
'

'
    Sheets("Loan Mail Box PersonnelList").Select
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("LoanMailBoxMainList[Duties Counter]")
    Range("LoanMailBoxMainList[Duties Counter]").Select
    Sheets("Morning PersonnelList").Select
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("MorningMainList[Duties Counter]")
    Range("MorningMainList[Duties Counter]").Select
    Sheets("Afternoon PersonnelList").Select
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AfternoonMainList[Duties Counter]")
    Range("AfternoonMainList[Duties Counter]").Select
    Sheets("AOH PersonnelList").Select
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "0"
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("AOHMainList[Duties Counter]")
    Range("AOHMainList[Duties Counter]").Select
    Sheets("Sat AOH PersonnelList").Select
    Range("H14").Select
    ActiveCell.FormulaR1C1 = "0"
    Selection.AutoFill Destination:=Range("SatAOHMainList[Duties Counter]")
    Range("SatAOHMainList[Duties Counter]").Select
    
    Set wsRoster = Sheets("Roster")
    wsRoster.Activate
End Sub
