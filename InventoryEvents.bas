Attribute VB_Name = "InventoryEvents"

Option Explicit

' ===================== CONFIGURATION =====================
Const TRANSACTIONS_SHEET As String = "Transactions"
Const TRANSACTIONS_TABLE As String = "tblTransactions"
Const SNAPSHOT_SHEET As String = "DailySnapshot"

' Event names (H2O2 uses simplified flow, NaOH uses Full -> Partial -> Empty)
Const EVT_INITIAL_FULL As String = "INITIAL_FULL"
Const EVT_INITIAL_EMPTY As String = "INITIAL_EMPTY"
Const EVT_INITIAL_PARTIAL As String = "INITIAL_PARTIAL"  ' For NaOH only
Const EVT_DELIVERY As String = "DELIVERY"
Const EVT_OPEN As String = "OPEN"            ' Converts Full -> Partial (NaOH only)
Const EVT_EMPTY As String = "EMPTY"          ' H2O2: Full -> Empty; NaOH:  Partial -> Empty
Const EVT_PICKUP As String = "PICKUP"

' Fix-count event names (no operational meaning)
Const EVT_FIX_FULL As String = "FIX_FULL"
Const EVT_FIX_PARTIAL As String = "FIX_PARTIAL"  ' For NaOH only
Const EVT_FIX_EMPTY As String = "FIX_EMPTY"

' Chemical IDs used in the table
Const CHEM_H2O2 As String = "H2O2"
Const CHEM_NAOH As String = "NAOH"

' ===================== PUBLIC BUTTON MACROS ======================
' Hydrogen Peroxide
Public Sub Empty_H2O2()
    LogEmpty CHEM_H2O2, "Hydrogen Peroxide"
End Sub

Public Sub Delivery_H2O2()
    LogDeliveryWithOptionalPickup CHEM_H2O2, "Hydrogen Peroxide"
End Sub

Public Sub Pickup_H2O2()
    LogPickupEmpties CHEM_H2O2, "Hydrogen Peroxide"
End Sub

' Sodium Hydroxide
Public Sub Open_NaOH()
    LogOpen CHEM_NAOH, "Sodium Hydroxide"
End Sub

Public Sub Empty_NaOH()
    LogEmpty CHEM_NAOH, "Sodium Hydroxide"
End Sub

Public Sub Delivery_NaOH()
    LogDeliveryWithOptionalPickup CHEM_NAOH, "Sodium Hydroxide"
End Sub

Public Sub Pickup_NaOH()
    LogPickupEmpties CHEM_NAOH, "Sodium Hydroxide"
End Sub

' Undo last transaction by the current user (handles batch transactions and snapshots)
Public Sub UndoMyLastTransaction()
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup
    
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET).ListObjects(TRANSACTIONS_TABLE)
    
    Dim userName As String
    userName = Environ$("Username")
    If Len(Trim(userName)) = 0 Then userName = Application.userName
    
    Dim i As Long
    Dim lr As ListRow
    Dim found As Boolean:  found = False
    Dim batchID As String
    Dim transactionDateTime As Date
    Dim triggeredSnapshot As Boolean
    
    ' Search bottom-up for last transaction by this user
    For i = lo.ListRows.count To 1 Step -1
        Set lr = lo.ListRows(i)
        Dim tech As String
        On Error Resume Next
        tech = CStr(lr.Range(lo.ListColumns("Technician").Index).Value)
        On Error GoTo 0
        If Len(Trim(tech)) = 0 Then
            ' skip rows without Technician (not owned)
        ElseIf StrComp(Trim(tech), Trim(userName), vbTextCompare) = 0 Then
            ' Get the BatchID and TriggeredSnapshot flag
            batchID = ""
            triggeredSnapshot = False
            On Error Resume Next
            If ColumnExists(lo, "BatchID") Then
                batchID = CStr(lr.Range(lo.ListColumns("BatchID").Index).Value)
            End If
            If ColumnExists(lo, "TriggeredSnapshot") Then
                triggeredSnapshot = CBool(lr.Range(lo.ListColumns("TriggeredSnapshot").Index).Value)
            End If
            transactionDateTime = CDate(lr.Range(lo.ListColumns("DateTime").Index).Value)
            On Error GoTo 0
            
            ' Collect all transactions in this batch for display
            Dim batchRows As Collection
            Set batchRows = New Collection
            Dim msg As String
            Dim j As Long
            Dim tempBatchID As String
            
            If Len(Trim(batchID)) > 0 Then
                ' Find all rows with this BatchID
                For j = lo.ListRows.count To 1 Step -1
                    On Error Resume Next
                    tempBatchID = CStr(lo.ListRows(j).Range(lo.ListColumns("BatchID").Index).Value)
                    On Error GoTo 0
                    If StrComp(Trim(tempBatchID), Trim(batchID), vbTextCompare) = 0 Then
                        batchRows.Add j
                        ' Check if any row in batch triggered snapshot
                        On Error Resume Next
                        If ColumnExists(lo, "TriggeredSnapshot") Then
                            If CBool(lo.ListRows(j).Range(lo.ListColumns("TriggeredSnapshot").Index).Value) Then
                                triggeredSnapshot = True
                            End If
                        End If
                        On Error GoTo 0
                    End If
                Next j
            Else
                ' Single transaction (no batch)
                batchRows.Add i
            End If
            
            ' Build confirmation message
            msg = "This will DELETE the following transaction(s) you logged:" & vbCrLf & vbCrLf
            
            Dim rowIdx As Variant
            Dim dt As Variant, chem As Variant, evt As Variant, qty As Variant, note As Variant
            For Each rowIdx In batchRows
                Set lr = lo.ListRows(CLng(rowIdx))
                dt = lr.Range(lo.ListColumns("DateTime").Index).Value
                chem = lr.Range(lo.ListColumns("ChemicalID").Index).Value
                evt = lr.Range(lo.ListColumns("Event").Index).Value
                qty = lr.Range(lo.ListColumns("Qty").Index).Value
                On Error Resume Next
                note = lr.Range(lo.ListColumns("Note").Index).Value
                On Error GoTo 0
                
                msg = msg & "  - " & Format(dt, "yyyy-mm-dd hh:nn:ss") & " | " & chem & " | " & evt & " | Qty: " & qty
                If Len(Trim(CStr(note))) > 0 Then msg = msg & " | " & CStr(note)
                msg = msg & vbCrLf
            Next rowIdx
            
            If triggeredSnapshot Then
                msg = msg & vbCrLf & "NOTE: This will also remove the snapshot triggered by this transaction." & vbCrLf
            End If
            
            msg = msg & vbCrLf & "Delete " & batchRows.count & " transaction(s)?"
            
            If MsgBox(msg, vbYesNo + vbQuestion, "Undo last transaction") = vbYes Then
                Dim wsTrans As Worksheet
                Set wsTrans = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
                
                ' === UNPROTECT before delete ===
                UnprotectSheet wsTrans
                
                ' Delete rows in reverse order (highest index first) to avoid index shifting
                Dim sortedRows() As Long
                ReDim sortedRows(1 To batchRows.count)
                Dim k As Long:  k = 1
                For Each rowIdx In batchRows
                    sortedRows(k) = CLng(rowIdx)
                    k = k + 1
                Next rowIdx
                
                ' Sort descending
                Call SortArrayDescending(sortedRows)
                
                For k = 1 To UBound(sortedRows)
                    lo.ListRows(sortedRows(k)).Delete
                Next k
                
                ' === RE-PROTECT after delete ===
                ProtectSheet wsTrans
                
                ' Remove snapshot if this transaction triggered one
                If triggeredSnapshot Then
                    RemoveSnapshotForDate transactionDateTime
                End If
                
                ' Update the "Last Updated" cell to reflect the new latest transaction
                UpdateLastUpdatedCell
                
                Application.StatusBar = "Deleted " & batchRows.count & " transaction(s) by " & userName
                MsgBox batchRows.count & " transaction(s) removed.", vbInformation
            Else
                MsgBox "No change made.", vbInformation
            End If
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "No previous transaction found owned by user '" & userName & "'.", vbInformation
    End If
    
    ' Restore original sheet
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' Helper to sort array descending
Private Sub SortArrayDescending(arr() As Long)
    Dim i As Long, j As Long, temp As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' Update the "Last Updated" cell (A2 on Inventory) with the latest transaction datetime
Private Sub UpdateLastUpdatedCell()
    Dim wsMain As Worksheet
    Dim wsTrans As Worksheet
    Dim lo As ListObject
    Dim latestDt As Variant
    
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets("Inventory")
    On Error GoTo 0
    
    If wsMain Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set wsTrans = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = wsTrans.ListObjects(TRANSACTIONS_TABLE)
    On Error GoTo 0
    
    If lo Is Nothing Then Exit Sub
    
    ' === UNPROTECT Inventory before changes ===
    UnprotectSheet wsMain
    
    ' Check if there are any transactions left
    If lo.DataBodyRange Is Nothing Then
        ' No transactions left - clear the cell
        wsMain.Range("A2").ClearContents
    Else
        ' Get the max DateTime from remaining transactions
        On Error Resume Next
        latestDt = Application.WorksheetFunction.Max(wsTrans.Range("tblTransactions[DateTime]"))
        On Error GoTo 0
        
        If IsEmpty(latestDt) Or IsError(latestDt) Then
            wsMain.Range("A2").ClearContents
        Else
            wsMain.Range("A2").Value = latestDt
        End If
    End If
    
    ' === RE-PROTECT Inventory after changes ===
    ProtectSheet wsMain
End Sub

' Remove snapshot that matches the date of the transaction (columns A-F only)
Private Sub RemoveSnapshotForDate(transactionDate As Date)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim snapshotDate As Date
    Dim targetDate As Date
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SNAPSHOT_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    targetDate = Int(transactionDate) ' Get date portion only
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    ' Search for snapshot matching this date (search from bottom to find most recent)
    For i = lastRow To 2 Step -1
        If IsDate(ws.Cells(i, "A").Value) Then
            snapshotDate = Int(CDate(ws.Cells(i, "A").Value))
            If snapshotDate = targetDate Then
                ' Check if row above is a header row (same values as row 1)
                If i > 2 Then
                    If ws.Cells(i - 1, "A").Value = ws.Cells(1, "A").Value Then
                        ' Clear the header row columns A-F only (don't delete entire row)
                        ws.Range("A" & (i - 1) & ":F" & (i - 1)).ClearContents
                    End If
                End If
                ' Clear the snapshot row columns A-F only (don't delete entire row)
                ws.Range("A" & i & ":F" & i).ClearContents
                Exit For
            End If
        End If
    Next i
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws
End Sub

' ===================== CORE EVENT LOGGING =====================

' OPEN:  Converts a Full tote to Partial (NaOH only)
Private Sub LogOpen(chemID As String, chemName As String)
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup
    
    Dim defaultQty As Long:  defaultQty = 1
    Dim qtyV As Variant
    qtyV = Application.InputBox("How many FULL totes are now PARTIAL for " & chemName & "?", "Tote(s) made partial", defaultQty, Type:=1)
    If VarType(qtyV) = vbBoolean And qtyV = False Then
        GoTo RestoreSheet
    End If
    If Not IsNumeric(qtyV) Then
        MsgBox "Please enter a number.", vbExclamation
        GoTo RestoreSheet
    End If
    Dim qty As Long: qty = CLng(qtyV)
    If qty <= 0 Then
        MsgBox "Quantity must be >= 1", vbExclamation
        GoTo RestoreSheet
    End If
    
    ' Check against FULL count
    Dim fullAvail As Long: fullAvail = GetFullCount(chemID)
    If qty > fullAvail Then
        MsgBox "Cannot change " & qty & " totes. Only " & fullAvail & " full tote(s) available.", vbExclamation
        GoTo RestoreSheet
    End If
    
    AddEvent chemID, EVT_OPEN, qty, "Changed " & qty & " full tote(s) to partial"

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' EMPTY: For H2O2 - converts Full -> Empty directly
'        For NaOH - converts Partial -> Empty
Private Sub LogEmpty(chemID As String, chemName As String)
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup
    
    Dim defaultQty As Long:  defaultQty = 1
    Dim qtyV As Variant
    Dim promptText As String
    Dim availCount As Long
    Dim availType As String
    
    ' Determine what we're converting based on chemical
    If StrComp(chemID, CHEM_NAOH, vbTextCompare) = 0 Then
        ' NaOH:  Empty converts Partial -> Empty
        availCount = GetPartialCount(chemID)
        availType = "PARTIAL"
        promptText = "How many PARTIAL totes are now EMPTY for " & chemName & "?"
    Else
        ' H2O2: Empty converts Full -> Empty directly
        availCount = GetFullCount(chemID)
        availType = "FULL"
        promptText = "How many FULL totes are now EMPTY for " & chemName & "?"
    End If
    
    qtyV = Application.InputBox(promptText, "Tote(s) Empty", defaultQty, Type:=1)
    If VarType(qtyV) = vbBoolean And qtyV = False Then
        GoTo RestoreSheet
    End If
    If Not IsNumeric(qtyV) Then
        MsgBox "Please enter a number.", vbExclamation
        GoTo RestoreSheet
    End If
    Dim qty As Long: qty = CLng(qtyV)
    If qty <= 0 Then
        MsgBox "Quantity must be >= 1", vbExclamation
        GoTo RestoreSheet
    End If
    
    ' Check against available count
    If qty > availCount Then
        MsgBox "Cannot mark " & qty & " empties. Only " & availCount & " " & LCase(availType) & " tote(s) available.", vbExclamation
        GoTo RestoreSheet
    End If
    
    If StrComp(chemID, CHEM_NAOH, vbTextCompare) = 0 Then
        AddEvent chemID, EVT_EMPTY, qty, "Marked " & qty & " partial tote(s) as empty"
    Else
        AddEvent chemID, EVT_EMPTY, qty, "Marked " & qty & " full tote(s) as empty"
    End If

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' ===================== MODIFIED DELIVERY LOGGING =====================

' DELIVERY WITH OPTIONAL PICKUP:   Logs both events as a batch with shared BatchID
' Now also checks for and removes matching upcoming deliveries
Private Sub LogDeliveryWithOptionalPickup(chemID As String, chemName As String)
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup

    Dim delivered As Variant, pickedUp As Variant
    Dim deliveryQty As Long, pickupQty As Long
    
    ' Step 1: Collect BOTH inputs before logging anything
    delivered = Application.InputBox("Number of FULL totes delivered for " & chemName & ":", "Delivery", 0, Type:=1)
    If VarType(delivered) = vbBoolean And delivered = False Then GoTo RestoreSheet
    If Not IsNumeric(delivered) Then
        MsgBox "Please enter a number.", vbExclamation
        GoTo RestoreSheet
    End If
    deliveryQty = CLng(delivered)

    pickedUp = Application.InputBox("Number of EMPTY totes picked up:", "Pickup Empties", 0, Type:=1)
    If VarType(pickedUp) = vbBoolean And pickedUp = False Then GoTo RestoreSheet
    If Not IsNumeric(pickedUp) Then
        MsgBox "Please enter a number.", vbExclamation
        GoTo RestoreSheet
    End If
    pickupQty = CLng(pickedUp)

    ' Step 2: Validate pickup against current empties BEFORE logging anything
    If pickupQty > 0 Then
        Dim empties As Long:   empties = GetEmptyCount(chemID)
        If pickupQty > empties Then
            MsgBox "Cannot pick up more empties than present (" & empties & ").", vbExclamation
            pickupQty = 0  ' Reset to 0, but still allow delivery to proceed
        End If
    End If

    ' Generate a unique BatchID for this combined delivery+pickup operation
    Dim batchID As String
    batchID = Format(Now, "yyyymmddhhnnss") & "_" & chemID & "_DELIVERY"
    
    ' Track if snapshot was triggered
    Dim snapshotTriggered As Boolean
    snapshotTriggered = False

    ' Step 3: Log delivery WITHOUT triggering snapshot
    If deliveryQty > 0 Then AddEventWithBatch chemID, EVT_DELIVERY, deliveryQty, "Delivery", batchID, False

    ' Step 4: Log pickup WITHOUT triggering snapshot
    If pickupQty > 0 Then AddEventWithBatch chemID, EVT_PICKUP, pickupQty, "Empty(s) picked up with delivery", batchID, False

    ' Step 5: Trigger daily snapshot AFTER both events are logged
    If deliveryQty > 0 Or pickupQty > 0 Then
        On Error Resume Next
        snapshotTriggered = MaybeLogDailySnapshotWithReturn()
        On Error GoTo 0
        
        ' Mark the batch as having triggered a snapshot
        If snapshotTriggered Then
            MarkBatchAsTriggeredSnapshot batchID
        End If
    End If
    
    ' Step 6: Check for and remove matching upcoming delivery
    If deliveryQty > 0 Then
        Dim upcomingRemoved As Boolean
        On Error Resume Next
        upcomingRemoved = TryRemoveMatchingUpcomingDelivery(chemID, deliveryQty)
        On Error GoTo 0
        
        If upcomingRemoved Then
            Application.StatusBar = "Delivery logged and matching upcoming delivery removed."
        End If
    End If

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

Private Sub LogPickupEmpties(chemID As String, chemName As String)
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup

    Dim qtyV As Variant
    qtyV = Application.InputBox("Number of EMPTY totes picked up for " & chemName & ":", "Pickup Empties", 0, Type:=1)
    If VarType(qtyV) = vbBoolean And qtyV = False Then GoTo RestoreSheet
    If Not IsNumeric(qtyV) Then
        MsgBox "Please enter a number.", vbExclamation
        GoTo RestoreSheet
    End If
    If CLng(qtyV) <= 0 Then GoTo RestoreSheet

    Dim empties As Long: empties = GetEmptyCount(chemID)
    If CLng(qtyV) > empties Then
        MsgBox "Cannot pick up more empties than present (" & empties & ").", vbExclamation
        GoTo RestoreSheet
    End If

    AddEvent chemID, EVT_PICKUP, CLng(qtyV), "Standalone pickup"

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' ===================== EVENT STORAGE =====================
Private Sub EnsureTransactionsSetup()
    Dim ws As Worksheet, lo As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        ws.Name = TRANSACTIONS_SHEET
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    On Error GoTo 0

    If lo Is Nothing Then
        ' === UNPROTECT before creating table ===
        UnprotectSheet ws
        
        ' Added BatchID and TriggeredSnapshot columns
        ws.Range("A1: H1").Value = Array("DateTime", "ChemicalID", "Event", "Qty", "Technician", "Note", "BatchID", "TriggeredSnapshot")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H1"), , xlYes)
        lo.Name = TRANSACTIONS_TABLE
        
        ' === RE-PROTECT after creating table ===
        ProtectSheet ws
    Else
        ' Ensure new columns exist in existing table
        EnsureColumnExists lo, "BatchID"
        EnsureColumnExists lo, "TriggeredSnapshot"
    End If
End Sub

' Helper to ensure a column exists in the table
Private Sub EnsureColumnExists(lo As ListObject, colName As String)
    If Not ColumnExists(lo, colName) Then
        Dim ws As Worksheet
        Set ws = lo.Parent
        
        ' === UNPROTECT before adding column ===
        UnprotectSheet ws
        
        lo.ListColumns.Add
        lo.ListColumns(lo.ListColumns.count).Name = colName
        
        ' === RE-PROTECT after adding column ===
        ProtectSheet ws
    End If
End Sub

' Add event AND trigger daily snapshot (for normal transactions)
Private Sub AddEvent(chemID As String, evt As String, qty As Long, noteText As String)
    Dim lo As ListObject, lr As ListRow
    Dim ws As Worksheet
    Dim snapshotTriggered As Boolean
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    
    ' Generate unique BatchID for single transactions
    Dim batchID As String
    batchID = Format(Now, "yyyymmddhhnnss") & "_" & chemID & "_" & evt
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    Set lr = lo.ListRows.Add

    With lr
        .Range(lo.ListColumns("DateTime").Index).Value = Now
        .Range(lo.ListColumns("ChemicalID").Index).Value = chemID
        .Range(lo.ListColumns("Event").Index).Value = evt
        .Range(lo.ListColumns("Qty").Index).Value = qty
        If ColumnExists(lo, "Technician") Then .Range(lo.ListColumns("Technician").Index).Value = Environ$("Username")
        If ColumnExists(lo, "Note") Then .Range(lo.ListColumns("Note").Index).Value = noteText
        If ColumnExists(lo, "BatchID") Then .Range(lo.ListColumns("BatchID").Index).Value = batchID
        If ColumnExists(lo, "TriggeredSnapshot") Then .Range(lo.ListColumns("TriggeredSnapshot").Index).Value = False
    End With
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws

    Application.StatusBar = evt & " logged (" & qty & ") for " & chemID
    
    ' Take daily snapshot AFTER the transaction (with new counts)
    ' Skip during initial stock events (handled by LogInitialSnapshot)
    If Not (evt = EVT_INITIAL_FULL Or evt = EVT_INITIAL_EMPTY Or evt = EVT_INITIAL_PARTIAL) Then
        On Error Resume Next
        snapshotTriggered = MaybeLogDailySnapshotWithReturn()
        On Error GoTo 0
        
        ' Mark this transaction as having triggered a snapshot
        If snapshotTriggered Then
            MarkBatchAsTriggeredSnapshot batchID
        End If
    End If
End Sub

' Add event with specific BatchID (for batch operations like delivery+pickup)
Private Sub AddEventWithBatch(chemID As String, evt As String, qty As Long, noteText As String, batchID As String, triggeredSnapshot As Boolean)
    Dim lo As ListObject, lr As ListRow
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    Set lr = lo.ListRows.Add

    With lr
        .Range(lo.ListColumns("DateTime").Index).Value = Now
        .Range(lo.ListColumns("ChemicalID").Index).Value = chemID
        .Range(lo.ListColumns("Event").Index).Value = evt
        .Range(lo.ListColumns("Qty").Index).Value = qty
        If ColumnExists(lo, "Technician") Then .Range(lo.ListColumns("Technician").Index).Value = Environ$("Username")
        If ColumnExists(lo, "Note") Then .Range(lo.ListColumns("Note").Index).Value = noteText
        If ColumnExists(lo, "BatchID") Then .Range(lo.ListColumns("BatchID").Index).Value = batchID
        If ColumnExists(lo, "TriggeredSnapshot") Then .Range(lo.ListColumns("TriggeredSnapshot").Index).Value = triggeredSnapshot
    End With
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws

    Application.StatusBar = evt & " logged (" & qty & ") for " & chemID
End Sub

' Mark all transactions in a batch as having triggered a snapshot
Private Sub MarkBatchAsTriggeredSnapshot(batchID As String)
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim i As Long
    Dim tempBatchID As String
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    
    If Not ColumnExists(lo, "TriggeredSnapshot") Then Exit Sub
    If Not ColumnExists(lo, "BatchID") Then Exit Sub
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    For i = 1 To lo.ListRows.count
        On Error Resume Next
        tempBatchID = CStr(lo.ListRows(i).Range(lo.ListColumns("BatchID").Index).Value)
        On Error GoTo 0
        If StrComp(Trim(tempBatchID), Trim(batchID), vbTextCompare) = 0 Then
            lo.ListRows(i).Range(lo.ListColumns("TriggeredSnapshot").Index).Value = True
        End If
    Next i
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws
End Sub

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

' ===================== AGGREGATION HELPERS =====================

Private Function GetFullCount(chemID As String) As Long
    ' For H2O2: Fulls = INITIAL_FULL + DELIVERY - EMPTY + FIX_FULL
    ' For NaOH:  Fulls = INITIAL_FULL + DELIVERY - OPEN + FIX_FULL (OPEN converts Full -> Partial)
    If StrComp(chemID, CHEM_NAOH, vbTextCompare) = 0 Then
        GetFullCount = SumEvents(chemID, EVT_INITIAL_FULL) _
                     + SumEvents(chemID, EVT_DELIVERY) _
                     - SumEvents(chemID, EVT_OPEN) _
                     + SumEvents(chemID, EVT_FIX_FULL)
    Else
        GetFullCount = SumEvents(chemID, EVT_INITIAL_FULL) _
                     + SumEvents(chemID, EVT_DELIVERY) _
                     - SumEvents(chemID, EVT_EMPTY) _
                     + SumEvents(chemID, EVT_FIX_FULL)
    End If
End Function

Private Function GetPartialCount(chemID As String) As Long
    ' Partials = INITIAL_PARTIAL + OPEN - EMPTY + FIX_PARTIAL
    ' Note: Only applicable for NaOH
    GetPartialCount = SumEvents(chemID, EVT_INITIAL_PARTIAL) _
                    + SumEvents(chemID, EVT_OPEN) _
                    - SumEvents(chemID, EVT_EMPTY) _
                    + SumEvents(chemID, EVT_FIX_PARTIAL)
End Function

Private Function GetEmptyCount(chemID As String) As Long
    ' Empties = INITIAL_EMPTY + EMPTY - PICKUP + FIX_EMPTY
    ' (EMPTY adds to empties regardless of whether it came from Full or Partial)
    GetEmptyCount = SumEvents(chemID, EVT_INITIAL_EMPTY) _
                  + SumEvents(chemID, EVT_EMPTY) _
                  - SumEvents(chemID, EVT_PICKUP) _
                  + SumEvents(chemID, EVT_FIX_EMPTY)
End Function

Private Function SumEvents(chemID As String, evt As String) As Long
    On Error GoTo SafeZero
    Dim ws As Worksheet:  Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Dim rngQty As Range, rngChem As Range, rngEvt As Range
    Set rngQty = ws.Range("tblTransactions[Qty]")
    Set rngChem = ws.Range("tblTransactions[ChemicalID]")
    Set rngEvt = ws.Range("tblTransactions[Event]")
    SumEvents = Application.WorksheetFunction.SumIfs(rngQty, rngChem, chemID, rngEvt, evt)
    Exit Function
SafeZero:
    SumEvents = 0
End Function

' ===================== INITIAL STOCK =====================
Public Sub AddInitialStock_Sequential()
    ' === PASSWORD PROTECTION ===
    If Not AuthorizeMacro("Add Initial Stock") Then Exit Sub
    
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup
    
    Dim chemList(1 To 2) As String
    chemList(1) = CHEM_H2O2
    chemList(2) = CHEM_NAOH
    
    Dim chemNameList(1 To 2) As String
    chemNameList(1) = "Hydrogen Peroxide"
    chemNameList(2) = "Sodium Hydroxide"
    
    ' Generate a unique BatchID for all initial stock transactions
    Dim batchID As String
    batchID = Format(Now, "yyyymmddhhnnss") & "_INITIAL_STOCK"
    
    Dim i As Long
    For i = 1 To 2
        Dim chem As String:  chem = chemList(i)
        Dim chemName As String: chemName = chemNameList(i)
        
        Dim fullQty As Variant
        fullQty = PromptForNonNegativeNumber("Enter FULL tote count for " & chemName & " (" & chem & "):", "Initial Full for " & chemName, 0)
        If fullQty = -1 Then GoTo RestoreSheet
        
        ' For NaOH, also prompt for partial count
        Dim partialQty As Variant
        If StrComp(chem, CHEM_NAOH, vbTextCompare) = 0 Then
            partialQty = PromptForNonNegativeNumber("Enter PARTIAL (open) tote count for " & chemName & " (" & chem & "):", "Initial Partial for " & chemName, 0)
            If partialQty = -1 Then GoTo RestoreSheet
        Else
            partialQty = 0
        End If
        
        Dim emptyQty As Variant
        emptyQty = PromptForNonNegativeNumber("Enter EMPTY tote count on site for " & chemName & " (" & chem & "):", "Initial Empty for " & chemName, 0)
        If emptyQty = -1 Then GoTo RestoreSheet
        
        If CLng(fullQty) >= 0 Then AddEventWithBatch chem, EVT_INITIAL_FULL, CLng(fullQty), "Initial full stock", batchID, False
        If StrComp(chem, CHEM_NAOH, vbTextCompare) = 0 And CLng(partialQty) > 0 Then AddEventWithBatch chem, EVT_INITIAL_PARTIAL, CLng(partialQty), "Initial partial stock", batchID, False
        If CLng(emptyQty) > 0 Then AddEventWithBatch chem, EVT_INITIAL_EMPTY, CLng(emptyQty), "Initial empties on site", batchID, False
    Next i
    
    ' Log an explicit snapshot that does NOT count as weekly/monthly
    ' This also marks the batch as having triggered a snapshot
    LogInitialSnapshotWithBatch batchID
    
    MsgBox "Initial stock recorded for H2O2 and NAOH.", vbInformation

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' Log initial snapshot and mark the batch as having triggered it
Private Sub LogInitialSnapshotWithBatch(batchID As String)
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    ' Call the snapshot logging
    LogInitialSnapshot
    
    ' Mark the batch as having triggered the snapshot
    MarkBatchAsTriggeredSnapshot batchID
    
    ' Restore original sheet
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' ===================== FIX COUNTS (additive adjustments, history preserved) =====================
Public Sub FixCounts_AddAdjustments()
    Dim originalSheet As Object
    Set originalSheet = ActiveSheet
    
    EnsureTransactionsSetup
    
    Dim chemList(1 To 2) As String
    chemList(1) = CHEM_H2O2
    chemList(2) = CHEM_NAOH
    
    Dim chemNameList(1 To 2) As String
    chemNameList(1) = "Hydrogen Peroxide"
    chemNameList(2) = "Sodium Hydroxide"
    
    ' Generate a unique BatchID for this fix-count operation
    Dim batchID As String
    batchID = Format(Now, "yyyymmddhhnnss") & "_FIXCOUNTS"
    
    Dim i As Long
    Dim anyLogged As Boolean:  anyLogged = False
    
    For i = 1 To 2
        Dim chem As String: chem = chemList(i)
        Dim chemName As String: chemName = chemNameList(i)
        
        ' Current totals
        Dim curFull As Long, curPartial As Long, curEmpty As Long
        curFull = GetFullCount(chem)
        curEmpty = GetEmptyCount(chem)
        If StrComp(chem, CHEM_NAOH, vbTextCompare) = 0 Then
            curPartial = GetPartialCount(chem)
        Else
            curPartial = 0
        End If
        
        ' Prompt for true totals
        Dim trueFull As Variant
        trueFull = PromptForNonNegativeNumber( _
            "Current FULL for " & chemName & " (" & chem & ") is " & curFull & "." & vbCrLf & _
            "Enter true FULL tote count:", _
            "Fix Counts - FULL for " & chemName, curFull)
        If trueFull = -1 Then GoTo RestoreSheet
        
        ' For NaOH, also prompt for partial count
        Dim truePartial As Variant
        If StrComp(chem, CHEM_NAOH, vbTextCompare) = 0 Then
            truePartial = PromptForNonNegativeNumber( _
                "Current PARTIAL for " & chemName & " (" & chem & ") is " & curPartial & "." & vbCrLf & _
                "Enter true PARTIAL tote count:", _
                "Fix Counts - PARTIAL for " & chemName, curPartial)
            If truePartial = -1 Then GoTo RestoreSheet
        Else
            truePartial = 0
        End If
        
        Dim trueEmpty As Variant
        trueEmpty = PromptForNonNegativeNumber( _
            "Current EMPTY for " & chemName & " (" & chem & ") is " & curEmpty & "." & vbCrLf & _
            "Enter true EMPTY tote count:", _
            "Fix Counts - EMPTY for " & chemName, curEmpty)
        If trueEmpty = -1 Then GoTo RestoreSheet
        
        ' Compute deltas
        Dim dFull As Long, dPartial As Long, dEmpty As Long
        dFull = CLng(trueFull) - curFull
        dPartial = CLng(truePartial) - curPartial
        dEmpty = CLng(trueEmpty) - curEmpty
        
        ' Log adjustments using the batch helper (so they can be undone together)
        If dFull <> 0 Then
            AddEventWithBatch chem, EVT_FIX_FULL, dFull, "Fixed FULL count from " & curFull & " to " & CLng(trueFull), batchID, False
            anyLogged = True
        End If
        If StrComp(chem, CHEM_NAOH, vbTextCompare) = 0 And dPartial <> 0 Then
            AddEventWithBatch chem, EVT_FIX_PARTIAL, dPartial, "Fixed PARTIAL count from " & curPartial & " to " & CLng(truePartial), batchID, False
            anyLogged = True
        End If
        If dEmpty <> 0 Then
            AddEventWithBatch chem, EVT_FIX_EMPTY, dEmpty, "Fixed EMPTY count from " & curEmpty & " to " & CLng(trueEmpty), batchID, False
            anyLogged = True
        End If
    Next i
    
    ' After logging all fix events, trigger a single daily snapshot (if any events were logged)
    If anyLogged Then
        Dim snapshotTriggered As Boolean
        On Error Resume Next
        snapshotTriggered = MaybeLogDailySnapshotWithReturn()
        On Error GoTo 0
        
        If snapshotTriggered Then
            MarkBatchAsTriggeredSnapshot batchID
        End If
    End If
    
    MsgBox "Fix Counts complete.", vbInformation

RestoreSheet:
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' Helper:  prompt for a non-negative number; returns -1 when user confirmed overall cancel
Private Function PromptForNonNegativeNumber(promptText As String, titleText As String, defaultVal As Long) As Variant
    Dim resp As Variant
    Do
        resp = Application.InputBox(promptText, titleText, defaultVal, Type:=1)
        If VarType(resp) = vbBoolean And resp = False Then
            If MsgBox("Cancel this operation entirely? ", _
                      vbYesNo + vbQuestion, "Cancel?  ") = vbYes Then
                PromptForNonNegativeNumber = -1
                Exit Function
            Else
                ' continue
            End If
        Else
            If Not IsNumeric(resp) Then
                MsgBox "Please enter a number (0 or greater).", vbExclamation
            ElseIf CLng(resp) < 0 Then
                MsgBox "Please enter a number 0 or greater.", vbExclamation
            Else
                PromptForNonNegativeNumber = CLng(resp)
                Exit Function
            End If
        End If
    Loop
End Function

' ===================== END MODULE =====================





