Attribute VB_Name = "ReprotectSheet"
Sub ReprotectSheet()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim password As String
    
    ' the protection password
    password = "nuslib2025"
    
    Set ws = ActiveSheet
    
    ' Reprotect the worksheet with password
    ws.Protect password, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub

ErrHandler:
    MsgBox "An error occurred while reprotecting the sheet: " & Err.Description, vbCritical
    Exit Sub
End Sub

