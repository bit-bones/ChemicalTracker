Attribute VB_Name = "DailySnapshot"

Option Explicit

Const TRANSACTIONS_SHEET As String = "Transactions"
Const TRANSACTIONS_TABLE As String = "tblTransactions"

Const SNAPSHOT_SHEET As String = "DailySnapshot"

' Event names
Const EVT_INITIAL_FULL As String = "INITIAL_FULL"
Const EVT_INITIAL_EMPTY As String = "INITIAL_EMPTY"
Const EVT_INITIAL_PARTIAL As String = "INITIAL_PARTIAL"  ' For NaOH only
Const EVT_DELIVERY As String = "DELIVERY"
Const EVT_OPEN As String = "OPEN"            ' Converts Full -> Partial (NaOH only)
Const EVT_EMPTY As String = "EMPTY"
Const EVT_PICKUP As String = "PICKUP"

' Fix-count event names
Const EVT_FIX_FULL As String = "FIX_FULL"
Const EVT_FIX_PARTIAL As String = "FIX_PARTIAL"  ' For NaOH only
Const EVT_FIX_EMPTY As String = "FIX_EMPTY"

Const CHEM_H2O2 As String = "H2O2"
Const CHEM_NAOH As String = "NAOH"

' Logs one snapshot per calendar day AFTER the first transaction of the day.
' Returns True if a snapshot was actually created, False otherwise.
Public Function MaybeLogDailySnapshotWithReturn() As Boolean
    Dim ws As Worksheet
    Dim lastDate As Date
    Dim lastRow As Long
    Dim today As Date
    Dim targetRow As Long
    Dim originalSheet As Object

    MaybeLogDailySnapshotWithReturn = False
    today = Date

    ' Save the currently active sheet to restore later
    Set originalSheet = ActiveSheet

    ' Ensure snapshot sheet exists (and header)
    Set ws = EnsureSnapshotSheet()

    ' Find last used row based ONLY on column A (snapshot datetime)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    If lastRow < 2 Then
        ' No prior snapshots at all -> first ever snapshot
        targetRow = 2
        WriteSnapshotRow ws, targetRow, "Daily snapshot"
        MaybeLogDailySnapshotWithReturn = True
        ' Restore original sheet
        On Error Resume Next
        originalSheet.Activate
        On Error GoTo 0
        Exit Function
    End If

    ' Read last snapshot's date (ignore time-of-day)
    lastDate = CDate(Int(ws.Cells(lastRow, "A").Value))

    ' If there's already a snapshot for today, do nothing
    If lastDate = today Then
        ' Restore original sheet
        On Error Resume Next
        originalSheet.Activate
        On Error GoTo 0
        Exit Function
    End If

    ' Default row where we'll write today's snapshot
    targetRow = lastRow + 1

    ' --- Determine if this is first snapshot of week/month ---
    Dim isFirstOfWeek As Boolean
    Dim isFirstOfMonth As Boolean
    isFirstOfWeek = IsFirstSnapshotOfWeek(ws, today)
    isFirstOfMonth = IsFirstSnapshotOfMonth(ws, today)

    ' --- Require a FULL prior week/month of history before writing summary notes ---
    Dim allowWeekly As Boolean
    Dim allowMonthly As Boolean
    allowWeekly = isFirstOfWeek And HasFullPriorWeekHistory(ws, today)
    allowMonthly = isFirstOfMonth And HasFullPriorMonthHistory(ws, today)

    ' --- Build note text (weekly/monthly/daily) ---
    Dim noteText As String
    Dim wH2O2 As Long, wNaOH As Long
    Dim mH2O2 As Long, mNaOH As Long

    noteText = "Daily snapshot"

    ' Weekly snapshot:  summarizing last full week (Mon-Sun),
    ' only if a complete prior week exists in history
    If allowWeekly Then
        GetWeeklyEmpties today, wH2O2, wNaOH
        noteText = "Weekly snapshot.    " & wH2O2 & " H2O2 and " & wNaOH & " NAOH totes emptied last week."
    End If

    ' Monthly snapshot: summarizing last full month,
    ' only if a complete prior month exists in history
    If allowMonthly Then
        GetMonthlyEmpties today, mH2O2, mNaOH

        ' If it's also a weekly snapshot, append with a space
        If noteText <> "" And noteText <> "Daily snapshot" Then
            noteText = noteText & " "
        Else
            noteText = ""
        End If

        noteText = noteText & "Monthly snapshot.  " & mH2O2 & " H2O2 and " & mNaOH & " NAOH totes emptied last month."

        ' === UNPROTECT before inserting header row ===
        UnprotectSheet ws

        ' Insert a header row just above this first snapshot of the month (A: G only)
        ws.Rows(targetRow).Insert Shift:=xlDown
        With ws
            .Range("A" & targetRow & ":G" & targetRow).Value = .Range("A1:G1").Value
            .Range("A1:G1").Copy
            .Range("A" & targetRow & ":G" & targetRow).PasteSpecial xlPasteFormats
        End With
        Application.CutCopyMode = False

        ' === RE-PROTECT after inserting header row ===
        ProtectSheet ws

        ' Move targetRow down by one so the actual daily snapshot goes just below the new header
        targetRow = targetRow + 1
    End If

    ' Write today's snapshot row with the composed note
    WriteSnapshotRow ws, targetRow, noteText
    MaybeLogDailySnapshotWithReturn = True

    ' Restore original sheet
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Function

' === LEGACY: Keep old function for backward compatibility ===
Public Sub MaybeLogDailySnapshot()
    Dim result As Boolean
    result = MaybeLogDailySnapshotWithReturn()
End Sub

' === PUBLIC: manually trigger a daily snapshot ===
' Uses the same logic as the automatic snapshot: skips if one already exists today,
' applies weekly/monthly notes when applicable.
Public Sub ManualSnapshot()
    ' === PASSWORD PROTECTION ===
    If Not AuthorizeMacro("Manual Snapshot") Then Exit Sub

    Dim snapshotCreated As Boolean
    snapshotCreated = MaybeLogDailySnapshotWithReturn()

    If snapshotCreated Then
        MsgBox "Snapshot logged successfully.", vbInformation, "Manual Snapshot"
    Else
        MsgBox "A snapshot for today already exists. No new snapshot was created.", vbInformation, "Manual Snapshot"
    End If
End Sub

' === OPTIONAL PUBLIC: log an explicit "Initial snapshot" that never counts as weekly/monthly ===
Public Sub LogInitialSnapshot()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim originalSheet As Object

    ' Save the currently active sheet to restore later
    Set originalSheet = ActiveSheet

    Set ws = EnsureSnapshotSheet()

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        targetRow = 2
    Else
        targetRow = lastRow + 1
    End If

    WriteSnapshotRow ws, targetRow, "Initial snapshot"

    ' Restore original sheet
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0
End Sub

' === PUBLIC: clear all snapshot rows (reset DailySnapshot list) ===
' Preserves helper/average cells in columns H: Q and beyond.
Public Sub ClearAllSnapshots()
    ' === PASSWORD PROTECTION ===
    If Not AuthorizeMacro("Clear All Snapshots") Then Exit Sub

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim originalSheet As Object

    ' Save the currently active sheet to restore later
    Set originalSheet = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SNAPSHOT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Snapshot sheet '" & SNAPSHOT_SHEET & "' not found.", vbInformation
        Exit Sub
    End If

    ' Find last snapshot row based on column A only
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No snapshot rows to clear.", vbInformation
        Exit Sub
    End If

    If MsgBox("This will DELETE all snapshot data in columns A:G (dates and counts)," & vbCrLf & _
              "but will leave any helper/average cells in columns H and beyond intact." & vbCrLf & vbCrLf & _
              "This cannot be undone except from a backup." & vbCrLf & vbCrLf & _
              "Continue? ", vbYesNo + vbExclamation, "Clear all daily snapshots") <> vbYes Then
        Exit Sub
    End If

    ' === UNPROTECT before clearing ===
    UnprotectSheet ws

    ' Clear only the snapshot data area (A:G), leave H:Q etc. untouched
    ws.Range("A2:G" & lastRow).ClearContents

    ' === RE-PROTECT after clearing ===
    ProtectSheet ws

    ' Restore original sheet
    On Error Resume Next
    originalSheet.Activate
    On Error GoTo 0

    MsgBox "All daily snapshots cleared from columns A:G.    Helper/average cells preserved.", vbInformation
End Sub

' === INTERNAL: ensure snapshot sheet and header exist ===
Private Function EnsureSnapshotSheet() As Worksheet
    Dim ws As Worksheet
    Dim originalSheet As Object

    ' Save the currently active sheet
    Set originalSheet = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SNAPSHOT_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        ws.Name = SNAPSHOT_SHEET

        ' === UNPROTECT before creating header ===
        UnprotectSheet ws

        With ws.Range("A1:G1")
            .Value = Array("SnapshotDateTime", _
                           "H2O2 Full", "H2O2 Empty", _
                           "NaOH Full", "NaOH Partial", "NaOH Empty", _
                           "Note")
            .Font.Bold = True
        End With

        ' === RE-PROTECT after creating header ===
        ProtectSheet ws

        ' Restore original sheet after creating new sheet
        On Error Resume Next
        originalSheet.Activate
        On Error GoTo 0
    End If

    Set EnsureSnapshotSheet = ws
End Function

' === INTERNAL: write snapshot to a specific row ===
Private Sub WriteSnapshotRow(ws As Worksheet, targetRow As Long, noteText As String)
    Dim hFull As Long, hEmpty As Long
    Dim nFull As Long, nPartial As Long, nEmpty As Long

    ' Compute current totals
    hFull = CurrentFullCount(CHEM_H2O2)
    hEmpty = CurrentEmptyCount(CHEM_H2O2)

    nFull = CurrentFullCount(CHEM_NAOH)
    nPartial = CurrentPartialCount(CHEM_NAOH)
    nEmpty = CurrentEmptyCount(CHEM_NAOH)

    ' === UNPROTECT before changes ===
    UnprotectSheet ws

    With ws
        .Cells(targetRow, "A").Value = Now
        .Cells(targetRow, "B").Value = hFull
        .Cells(targetRow, "C").Value = hEmpty
        .Cells(targetRow, "D").Value = nFull
        .Cells(targetRow, "E").Value = nPartial
        .Cells(targetRow, "F").Value = nEmpty

        If Len(noteText) > 0 Then
            .Cells(targetRow, "G").Value = noteText
        Else
            .Cells(targetRow, "G").Value = "Daily snapshot"
        End If
    End With

    ' === RE-PROTECT after changes ===
    ProtectSheet ws
End Sub

' === HELPER FUNCTIONS:  count calculations ===

Private Function CurrentFullCount(chemID As String) As Long
    ' For H2O2: Fulls = INITIAL_FULL + DELIVERY - EMPTY + FIX_FULL
    ' For NaOH:  Fulls = INITIAL_FULL + DELIVERY - OPEN + FIX_FULL
    If chemID = CHEM_NAOH Then
        CurrentFullCount = SumEventsSimple(chemID, EVT_INITIAL_FULL) _
                         + SumEventsSimple(chemID, EVT_DELIVERY) _
                         - SumEventsSimple(chemID, EVT_OPEN) _
                         + SumEventsSimple(chemID, EVT_FIX_FULL)
    Else
        CurrentFullCount = SumEventsSimple(chemID, EVT_INITIAL_FULL) _
                         + SumEventsSimple(chemID, EVT_DELIVERY) _
                         - SumEventsSimple(chemID, EVT_EMPTY) _
                         + SumEventsSimple(chemID, EVT_FIX_FULL)
    End If
End Function

Private Function CurrentPartialCount(chemID As String) As Long
    ' Partials = INITIAL_PARTIAL + OPEN - EMPTY + FIX_PARTIAL
    ' Only applicable for NaOH
    CurrentPartialCount = SumEventsSimple(chemID, EVT_INITIAL_PARTIAL) _
                        + SumEventsSimple(chemID, EVT_OPEN) _
                        - SumEventsSimple(chemID, EVT_EMPTY) _
                        + SumEventsSimple(chemID, EVT_FIX_PARTIAL)
End Function

Private Function CurrentEmptyCount(chemID As String) As Long
    ' Empties = INITIAL_EMPTY + EMPTY - PICKUP + FIX_EMPTY
    CurrentEmptyCount = SumEventsSimple(chemID, EVT_INITIAL_EMPTY) _
                      + SumEventsSimple(chemID, EVT_EMPTY) _
                      - SumEventsSimple(chemID, EVT_PICKUP) _
                      + SumEventsSimple(chemID, EVT_FIX_EMPTY)
End Function

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

' === WEEKLY / MONTHLY DETECTION AND COUNTS ===========================

' True if today is the first snapshot day of its ISO week (Mon-Sun) in the sheet.
' Initial snapshot rows (note begins with "Initial snapshot") are ignored.
Private Function IsFirstSnapshotOfWeek(ws As Worksheet, today As Date) As Boolean
    Dim lastRow As Long, i As Long
    Dim dt As Variant
    Dim thisWeek As Long
    Dim noteText As String

    thisWeek = datePart("ww", today, vbMonday, vbFirstFourDays)

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        dt = ws.Cells(i, "A").Value
        If IsDate(dt) Then
            noteText = CStr(ws.Cells(i, "G").Value)   ' Column G is now Note
            If Left$(LCase$(Trim$(noteText)), 16) <> "initial snapshot" Then
                If datePart("ww", CDate(dt), vbMonday, vbFirstFourDays) = thisWeek And _
                   Year(CDate(dt)) = Year(today) Then
                    IsFirstSnapshotOfWeek = False
                    Exit Function
                End If
            End If
        End If
    Next i

    IsFirstSnapshotOfWeek = True
End Function

' True if today is the first snapshot day of its month in the sheet.
' Initial snapshot rows (note begins with "Initial snapshot") are ignored.
Private Function IsFirstSnapshotOfMonth(ws As Worksheet, today As Date) As Boolean
    Dim lastRow As Long, i As Long
    Dim dt As Variant
    Dim noteText As String

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        dt = ws.Cells(i, "A").Value
        If IsDate(dt) Then
            noteText = CStr(ws.Cells(i, "G").Value)   ' Column G is now Note
            If Left$(LCase$(Trim$(noteText)), 16) <> "initial snapshot" Then
                If Month(CDate(dt)) = Month(today) And Year(CDate(dt)) = Year(today) Then
                    IsFirstSnapshotOfMonth = False
                    Exit Function
                End If
            End If
        End If
    Next i

    IsFirstSnapshotOfMonth = True
End Function

' Return the earliest non-initial snapshot date, or Empty if none
Private Function EarliestNonInitialSnapshotDate(ws As Worksheet) As Variant
    Dim lastRow As Long, i As Long
    Dim dt As Variant, noteText As String
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        dt = ws.Cells(i, "A").Value
        If IsDate(dt) Then
            noteText = CStr(ws.Cells(i, "G").Value)   ' Column G is now Note
            If Left$(LCase$(Trim$(noteText)), 16) <> "initial snapshot" Then
                EarliestNonInitialSnapshotDate = CDate(dt)
                Exit Function
            End If
        End If
    Next i
    EarliestNonInitialSnapshotDate = Empty
End Function

' True if there is at least one non-initial snapshot on or before the start of LAST week
Private Function HasFullPriorWeekHistory(ws As Worksheet, today As Date) As Boolean
    Dim dow As Integer
    Dim startOfThisWeek As Date
    Dim startOfLastWeek As Date
    Dim earliest As Variant

    dow = Weekday(today, vbMonday)                 ' Monday=1
    startOfThisWeek = DateAdd("d", 1 - dow, today) ' Monday of this week
    startOfLastWeek = DateAdd("d", -7, startOfThisWeek)

    earliest = EarliestNonInitialSnapshotDate(ws)
    If IsEmpty(earliest) Then
        HasFullPriorWeekHistory = False
    Else
        HasFullPriorWeekHistory = (CDate(earliest) <= startOfLastWeek)
    End If
End Function

' True if there is at least one non-initial snapshot on or before the FIRST DAY of LAST month
Private Function HasFullPriorMonthHistory(ws As Worksheet, today As Date) As Boolean
    Dim firstDayPrevMonth As Date
    Dim earliest As Variant

    firstDayPrevMonth = DateSerial(Year(today), Month(today) - 1, 1)
    earliest = EarliestNonInitialSnapshotDate(ws)
    If IsEmpty(earliest) Then
        HasFullPriorMonthHistory = False
    Else
        HasFullPriorMonthHistory = (CDate(earliest) <= firstDayPrevMonth)
    End If
End Function

' Count empties (EMPTY only) for the previous week (Mon-Sun) ending before "today".
Private Sub GetWeeklyEmpties(today As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    Dim startDate As Date, endDate As Date
    Dim dow As Integer

    today = Int(today) ' Strip time so range lands on calendar-day boundaries

    ' Find start of this week (Monday)
    dow = Weekday(today, vbMonday)      ' Monday=1, Sunday=7
    startDate = DateAdd("d", 1 - dow, today)

    ' Previous week range:  Monday..  Sunday of the week before this one
    endDate = DateAdd("d", -1, startDate)
    startDate = DateAdd("d", -7, startDate)

    Call SumEmptiesInDateRange(startDate, endDate, h2o2Count, naohCount)
End Sub

' Count empties (EMPTY only) for the previous month ending before "today".
Private Sub GetMonthlyEmpties(today As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    Dim prevMonth As Date
    Dim startDate As Date, endDate As Date

    today = Int(today) ' Strip time so range lands on calendar-day boundaries

    ' Previous month:  use DateSerial logic
    prevMonth = DateSerial(Year(today), Month(today) - 1, 1)
    startDate = prevMonth
    endDate = DateSerial(Year(prevMonth), Month(prevMonth) + 1, 0) ' last day of prev month

    Call SumEmptiesInDateRange(startDate, endDate, h2o2Count, naohCount)
End Sub

' Sum EMPTY events in tblTransactions for each chemical over a date range (inclusive).
' Counts only positive Qty values (<=0 ignored) so correction rows are excluded from snapshot totals.
Private Sub SumEmptiesInDateRange(startDate As Date, endDate As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    On Error GoTo SafeZero
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim dt As Variant, chem As Variant, evt As Variant, qty As Variant
    Dim qtyLong As Long

    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)

    h2o2Count = 0
    naohCount = 0

    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.DataBodyRange.Rows.count
        dt = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("DateTime").Index).Value
        If IsDate(dt) Then
            If Int(CDate(dt)) >= Int(startDate) And Int(CDate(dt)) <= Int(endDate) Then
                evt = CStr(lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("Event").Index).Value)
                ' Consider only true EMPTY conversion events.
                If evt = EVT_EMPTY Then
                    qty = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("Qty").Index).Value
                    ' Only count positive numeric Qty values; ignore zero/negative and non-numeric
                    If IsNumeric(qty) Then
                        qtyLong = CLng(qty)
                        If qtyLong > 0 Then
                            chem = CStr(lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("ChemicalID").Index).Value)
                            If chem = CHEM_H2O2 Then
                                h2o2Count = h2o2Count + qtyLong
                            ElseIf chem = CHEM_NAOH Then
                                naohCount = naohCount + qtyLong
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    Exit Sub
SafeZero:
    h2o2Count = 0
    naohCount = 0
End Sub

