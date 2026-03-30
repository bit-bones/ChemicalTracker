Attribute VB_Name = "UpdateSnapshotNotes"

Option Explicit

' Refresh existing DailySnapshot note lines (Weekly / Monthly) to use the new
' "count positive EMPTY and FIX_EMPTY transactions" logic.
'
'  - Find rows on the DailySnapshot sheet whose Note contains "Weekly snapshot"
'    and/or "Monthly snapshot" (ignores "Initial snapshot").
'  - Recompute the counts by summing actual EMPTY and FIX_EMPTY transactions
'    in tblTransactions for the same date ranges the snapshot code uses.
'  - Replace the Note cell for those rows with updated phrasing:
'       "Weekly snapshot.    X H2O2 and Y NAOH totes emptied last week."
'       "Monthly snapshot.  X H2O2 and Y NAOH totes emptied last month."
'    If a row originally contained both weekly and monthly, both sentences will
'    be included (weekly first, then monthly).
'
'  - This only writes to column G (Note) on DailySnapshot and will not delete rows.

Const TRANSACTIONS_SHEET As String = "Transactions"
Const TRANSACTIONS_TABLE As String = "tblTransactions"
Const SNAPSHOT_SHEET As String = "DailySnapshot"

Const EVT_EMPTY As String = "EMPTY"
Const EVT_FIX_EMPTY As String = "FIX_EMPTY"

Const CHEM_H2O2 As String = "H2O2"
Const CHEM_NAOH As String = "NAOH"

Public Sub RefreshSnapshotNotes()
    Dim wsSnap As Worksheet
    Dim lastRow As Long, i As Long
    Dim noteText As String
    Dim snapshotDate As Variant
    Dim updatedCount As Long:  updatedCount = 0
    
    On Error GoTo ErrHandler
    Set wsSnap = ThisWorkbook.Worksheets(SNAPSHOT_SHEET)
    
    lastRow = wsSnap.Cells(wsSnap.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No snapshot rows found on " & SNAPSHOT_SHEET, vbInformation
        Exit Sub
    End If
    
    ' Loop through snapshot rows
    For i = 2 To lastRow
        noteText = CStr(wsSnap.Cells(i, "G").Value)  ' Column G is now Note
        If Len(Trim(noteText)) = 0 Then GoTo NextRow
        
        ' Skip initial snapshot rows
        If LCase$(Left$(Trim(noteText), 16)) = "initial snapshot" Then GoTo NextRow
        
        ' Decide whether this row is a weekly/monthly snapshot (old notes contain those phrases)
        Dim wantsWeekly As Boolean:  wantsWeekly = (InStr(1, noteText, "Weekly snapshot", vbTextCompare) > 0)
        Dim wantsMonthly As Boolean: wantsMonthly = (InStr(1, noteText, "Monthly snapshot", vbTextCompare) > 0)
        
        If Not wantsWeekly And Not wantsMonthly Then GoTo NextRow
        
        ' Ensure the snapshot row has a valid date in column A
        snapshotDate = wsSnap.Cells(i, "A").Value
        If Not IsDate(snapshotDate) Then GoTo NextRow
        
        Dim wH2O2 As Long, wNaOH As Long
        Dim mH2O2 As Long, mNaOH As Long
        wH2O2 = 0: wNaOH = 0: mH2O2 = 0: mNaOH = 0
        
        If wantsWeekly Then
            Call GetWeeklyEmptiesForSnapshot(CDate(snapshotDate), wH2O2, wNaOH)
        End If
        If wantsMonthly Then
            Call GetMonthlyEmptiesForSnapshot(CDate(snapshotDate), mH2O2, mNaOH)
        End If
        
        ' Build new note text:  follow the same wording used by the snapshot module
        Dim newNote As String
        newNote = ""
        If wantsWeekly Then
            newNote = "Weekly snapshot.   " & wH2O2 & " H2O2 and " & wNaOH & " NAOH totes emptied last week."
        End If
        If wantsMonthly Then
            If Len(newNote) > 0 Then newNote = newNote & " "
            newNote = newNote & "Monthly snapshot.  " & mH2O2 & " H2O2 and " & mNaOH & " NAOH totes emptied last month."
        End If
        
        ' Write the updated note (only column G)
        UnprotectSheet wsSnap
        wsSnap.Cells(i, "G").Value = newNote
        ProtectSheet wsSnap
        
        updatedCount = updatedCount + 1
NextRow:
    Next i
    
    MsgBox "Snapshot notes refreshed for " & updatedCount & " row(s).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in RefreshSnapshotNotes: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' --- Helpers --------------------------------------------------------

' Compute previous full week (Mon-Sun) empties for the snapshot date (same logic as DailySnapshot).
' 'today' should be the snapshot row date (date portion).
Private Sub GetWeeklyEmptiesForSnapshot(today As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    Dim dow As Integer
    Dim startDate As Date, endDate As Date
    
    dow = Weekday(today, vbMonday) ' Monday = 1
    startDate = DateAdd("d", 1 - dow, today) ' Monday of this week
    endDate = DateAdd("d", -1, startDate)    ' Sunday of prior week
    startDate = DateAdd("d", -7, startDate)  ' Monday of prior week
    
    Call SumPositiveEmptyEventsInRange(startDate, endDate, h2o2Count, naohCount)
End Sub

' Compute previous full month empties for the snapshot date (same logic as DailySnapshot).
Private Sub GetMonthlyEmptiesForSnapshot(today As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    Dim prevMonth As Date
    Dim startDate As Date, endDate As Date
    
    prevMonth = DateSerial(Year(today), Month(today) - 1, 1)
    startDate = prevMonth
    endDate = DateSerial(Year(prevMonth), Month(prevMonth) + 1, 0) ' last day of previous month
    
    Call SumPositiveEmptyEventsInRange(startDate, endDate, h2o2Count, naohCount)
End Sub

' Sum EMPTY and FIX_EMPTY events in tblTransactions between startDate and endDate inclusive.
' Only positive numeric Qty values are counted (<=0 ignored).
Private Sub SumPositiveEmptyEventsInRange(startDate As Date, endDate As Date, ByRef h2o2Count As Long, ByRef naohCount As Long)
    On Error GoTo SafeZero
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim dt As Variant, evt As String, chem As String, qty As Variant
    Dim qtyLong As Long
    
    h2o2Count = 0
    naohCount = 0
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = ws.ListObjects(TRANSACTIONS_TABLE)
    
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    For i = 1 To lo.DataBodyRange.Rows.count
        dt = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("DateTime").Index).Value
        If Not IsDate(dt) Then GoTo NextRowLoop
        
        If CDate(dt) < startDate Or CDate(dt) > endDate Then GoTo NextRowLoop
        
        evt = CStr(lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("Event").Index).Value)
        If Not (evt = EVT_EMPTY Or evt = EVT_FIX_EMPTY) Then GoTo NextRowLoop
        
        qty = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("Qty").Index).Value
        If Not IsNumeric(qty) Then GoTo NextRowLoop
        qtyLong = CLng(qty)
        If qtyLong <= 0 Then GoTo NextRowLoop
        
        chem = CStr(lo.DataBodyRange.Rows(i).Columns(lo.ListColumns("ChemicalID").Index).Value)
        If chem = CHEM_H2O2 Then
            h2o2Count = h2o2Count + qtyLong
        ElseIf chem = CHEM_NAOH Then
            naohCount = naohCount + qtyLong
        End If
NextRowLoop:
    Next i
    
    Exit Sub
SafeZero:
    h2o2Count = 0
    naohCount = 0
End Sub

