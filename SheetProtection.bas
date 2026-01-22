Attribute VB_Name = "SheetProtection"

Option Explicit

' ===================== SHEET PROTECTION HELPERS =====================
' Centralized password - change this to your desired password
Private Const SHEET_PASSWORD As String = "3450"

' ===================== MACRO PASSWORD PROTECTION =====================
' Returns True if password is correct, False if cancelled or wrong
Public Function AuthorizeMacro(Optional macroName As String = "this action") As Boolean
    Dim enteredPwd As String
    
    enteredPwd = InputBox("Enter password to run " & macroName & ":", "Password Required")
    
    ' User clicked Cancel or left blank
    If Len(Trim(enteredPwd)) = 0 Then
        AuthorizeMacro = False
        Exit Function
    End If
    
    ' Check password (uses same password as sheet protection)
    If StrComp(enteredPwd, SHEET_PASSWORD, vbTextCompare) = 0 Then
        AuthorizeMacro = True
    Else
        MsgBox "Incorrect password.", vbExclamation, "Access Denied"
        AuthorizeMacro = False
    End If
End Function

' Unprotect a worksheet (call before making changes)
Public Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

' Re-protect a worksheet (call after making changes)
Public Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD, _
        UserInterfaceOnly:=False, _
        AllowFiltering:=True, _
        AllowSorting:=True
    On Error GoTo 0
End Sub

' Protect all sheets in the workbook (run once after setup)
Public Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws
    Next ws
    MsgBox "All sheets are now protected.", vbInformation
End Sub

' Unprotect all sheets (for admin use only)
Public Sub UnprotectAllSheets()
    ' === PASSWORD PROTECTION ===
    If Not AuthorizeMacro("Unprotect All Sheets") Then Exit Sub
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        UnprotectSheet ws
    Next ws
    MsgBox "All sheets are now unprotected.", vbInformation
End Sub

' Protect the workbook structure (prevent adding/deleting/renaming sheets)
Public Sub ProtectWorkbookStructure()
    On Error Resume Next
    ThisWorkbook.Protect Password:=SHEET_PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0
    MsgBox "Workbook structure is now protected.", vbInformation
End Sub

' Unprotect the workbook structure
Public Sub UnprotectWorkbookStructure()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
    MsgBox "Workbook structure is now unprotected.", vbInformation
End Sub



