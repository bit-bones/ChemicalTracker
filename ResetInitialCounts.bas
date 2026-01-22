Attribute VB_Name = "ResetInitialCounts"

Option Explicit

Const TRANSACTIONS_SHEET As String = "Transactions"
Const TRANSACTIONS_TABLE As String = "tblTransactions"

' Event names
Const EVT_INITIAL_FULL As String = "INITIAL_FULL"
Const EVT_INITIAL_EMPTY As String = "INITIAL_EMPTY"
Const EVT_INITIAL_PARTIAL As String = "INITIAL_PARTIAL"  ' For NaOH only
Const EVT_DELIVERY As String = "DELIVERY"
Const EVT_OPEN As String = "OPEN"
Const EVT_EMPTY As String = "EMPTY"
Const EVT_PICKUP As String = "PICKUP"

' Fix-count event names (must match the main module)
Const EVT_FIX_FULL As String = "FIX_FULL"
Const EVT_FIX_PARTIAL As String = "FIX_PARTIAL"  ' For NaOH only
Const EVT_FIX_EMPTY As String = "FIX_EMPTY"

Const CHEM_H2O2 As String = "H2O2"
Const CHEM_NAOH As String = "NAOH"

' -------------------------
' Full wipe:  delete all transactions (keeps header)
' -------------------------
Public Sub ClearAllTransactions_Reset()
    ' === PASSWORD PROTECTION ===
    If Not AuthorizeMacro("Clear All Transactions") Then Exit Sub
    
    Dim ws As Worksheet, lo As ListObject
    Dim wsMain As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Transactions sheet '" & TRANSACTIONS_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Transactions table '" & TRANSACTIONS_TABLE & "' not found on sheet '" & TRANSACTIONS_SHEET & "'.", vbExclamation
        Exit Sub
    End If

    If lo.DataBodyRange Is Nothing Then
        MsgBox "Transactions table already empty.", vbInformation
        Exit Sub
    End If

    If MsgBox("This will DELETE ALL transaction rows in '" & TRANSACTIONS_TABLE & "'.    This cannot be undone except from a backup.   Continue?", vbYesNo + vbExclamation, "Confirm delete all") <> vbYes Then
        Exit Sub
    End If

    ' === UNPROTECT before deleting ===
    UnprotectSheet ws
    
    lo.DataBodyRange.Delete
    
    ' === RE-PROTECT after deleting ===
    ProtectSheet ws
    
    ' Clear the "Last Updated" cell on Inventory
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets("Inventory")
    On Error GoTo 0
    
    If Not wsMain Is Nothing Then
        ' === UNPROTECT Inventory before changes ===
        UnprotectSheet wsMain
        
        wsMain.Range("A2").ClearContents
        
        ' === RE-PROTECT Inventory after changes ===
        ProtectSheet wsMain
    End If
    
    MsgBox "All transactions cleared.   You can now add new initial stock using AddInitialStock.", vbInformation
End Sub

' -------------------------
' Quick display of current counts (includes NaOH partial)
' -------------------------
Public Sub ShowCurrentCounts_Reset()
    Dim msg As String
    Dim hf As Long, ef As Long
    Dim hn As Long, np As Long, en As Long

    On Error Resume Next
    ' H2O2 FULL:  INITIAL_FULL + DELIVERY - EMPTY + FIX_FULL
    hf = SumEventsSimple(CHEM_H2O2, EVT_INITIAL_FULL) _
       + SumEventsSimple(CHEM_H2O2, EVT_DELIVERY) _
       - SumEventsSimple(CHEM_H2O2, EVT_EMPTY) _
       + SumEventsSimple(CHEM_H2O2, EVT_FIX_FULL)

    ' H2O2 EMPTY: INITIAL_EMPTY + EMPTY - PICKUP + FIX_EMPTY
    ef = SumEventsSimple(CHEM_H2O2, EVT_INITIAL_EMPTY) _
       + SumEventsSimple(CHEM_H2O2, EVT_EMPTY) _
       - SumEventsSimple(CHEM_H2O2, EVT_PICKUP) _
       + SumEventsSimple(CHEM_H2O2, EVT_FIX_EMPTY)

    ' NaOH FULL: INITIAL_FULL + DELIVERY - OPEN + FIX_FULL
    hn = SumEventsSimple(CHEM_NAOH, EVT_INITIAL_FULL) _
       + SumEventsSimple(CHEM_NAOH, EVT_DELIVERY) _
       - SumEventsSimple(CHEM_NAOH, EVT_OPEN) _
       + SumEventsSimple(CHEM_NAOH, EVT_FIX_FULL)

    ' NaOH PARTIAL: INITIAL_PARTIAL + OPEN - EMPTY + FIX_PARTIAL
    np = SumEventsSimple(CHEM_NAOH, EVT_INITIAL_PARTIAL) _
       + SumEventsSimple(CHEM_NAOH, EVT_OPEN) _
       - SumEventsSimple(CHEM_NAOH, EVT_EMPTY) _
       + SumEventsSimple(CHEM_NAOH, EVT_FIX_PARTIAL)

    ' NaOH EMPTY:  INITIAL_EMPTY + EMPTY - PICKUP + FIX_EMPTY
    en = SumEventsSimple(CHEM_NAOH, EVT_INITIAL_EMPTY) _
       + SumEventsSimple(CHEM_NAOH, EVT_EMPTY) _
       - SumEventsSimple(CHEM_NAOH, EVT_PICKUP) _
       + SumEventsSimple(CHEM_NAOH, EVT_FIX_EMPTY)
    On Error GoTo 0

    msg = "Current counts (computed from tblTransactions):" & vbCrLf & vbCrLf
    msg = msg & "H2O2 - Full: " & hf & " | Empty: " & ef & vbCrLf
    msg = msg & "NaOH - Full: " & hn & " | Partial: " & np & " | Empty: " & en & vbCrLf

    MsgBox msg, vbInformation, "Current Inventory"
End Sub

' -------------------------
' Helper:  sum events safely
' -------------------------
Private Function SumEventsSimple(chemID As String, evt As String) As Long
    On Error GoTo SafeZero
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    SumEventsSimple = Application.WorksheetFunction.SumIfs( _
                        ws.Range("tblTransactions[Qty]"), _
                        ws.Range("tblTransactions[ChemicalID]"), chemID, _
                        ws.Range("tblTransactions[Event]"), evt)
    Exit Function
SafeZero:
    SumEventsSimple = 0
End Function

Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If StrComp(c.Name, colName, vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next c
    ColumnExists = False
End Function



