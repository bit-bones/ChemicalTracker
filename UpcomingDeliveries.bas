Attribute VB_Name = "UpcomingDeliveries"

Option Explicit

' ===================== CONFIGURATION =====================
Const MAIN_SHEET As String = "Inventory"
Const UPCOMING_START_ROW As Long = 2        ' First data row for upcoming deliveries
Const COL_UPCOMING_DATE As String = "H"     ' Upcoming Delivery Date column
Const COL_H2O2_EXPECTED As String = "I"     ' H2O2 Expected column
Const COL_NAOH_EXPECTED As String = "J"     ' NAOH Expected column

' ===================== PUBLIC MACROS =====================

' Add a new upcoming delivery to the list
Public Sub AddUpcomingDelivery()
    Dim ws As Worksheet
    Dim deliveryDate As Variant
    Dim h2o2Qty As Variant
    Dim naohQty As Variant
    Dim targetRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & MAIN_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt for delivery date
    deliveryDate = Application.InputBox( _
        "Enter the expected delivery date (e.g., 2026-01-15):", _
        "Upcoming Delivery - Date", _
        Format(Date + 1, "yyyy-mm-dd"), _
        Type:=2)  ' Type 2 = Text
    
    If VarType(deliveryDate) = vbBoolean And deliveryDate = False Then Exit Sub
    If Len(Trim(CStr(deliveryDate))) = 0 Then Exit Sub
    
    ' Validate and convert to date
    If Not IsDate(deliveryDate) Then
        MsgBox "Invalid date format. Please enter a valid date.", vbExclamation
        Exit Sub
    End If
    deliveryDate = CDate(deliveryDate)
    
    ' Reject dates in the past (allow today)
    If Int(deliveryDate) < Date Then
        MsgBox "Cannot add an upcoming delivery with a date in the past.", vbExclamation, "Invalid Date"
        Exit Sub
    End If
    
    ' Prompt for H2O2 quantity
    h2o2Qty = Application.InputBox( _
        "Enter expected H2O2 totes for " & Format(deliveryDate, "yyyy-mm-dd") & ":", _
        "Upcoming Delivery - H2O2 Quantity", _
        0, Type:=1)  ' Type 1 = Number
    
    If VarType(h2o2Qty) = vbBoolean And h2o2Qty = False Then Exit Sub
    If Not IsNumeric(h2o2Qty) Or CLng(h2o2Qty) < 0 Then
        MsgBox "Please enter a valid non-negative number.", vbExclamation
        Exit Sub
    End If
    h2o2Qty = CLng(h2o2Qty)
    
    ' Prompt for NaOH quantity
    naohQty = Application.InputBox( _
        "Enter expected NAOH totes for " & Format(deliveryDate, "yyyy-mm-dd") & ":", _
        "Upcoming Delivery - NAOH Quantity", _
        0, Type:=1)
    
    If VarType(naohQty) = vbBoolean And naohQty = False Then Exit Sub
    If Not IsNumeric(naohQty) Or CLng(naohQty) < 0 Then
        MsgBox "Please enter a valid non-negative number.", vbExclamation
        Exit Sub
    End If
    naohQty = CLng(naohQty)
    
    ' At least one quantity should be > 0
    If h2o2Qty = 0 And naohQty = 0 Then
        MsgBox "At least one chemical quantity must be greater than 0.", vbExclamation
        Exit Sub
    End If
    
    ' Find the target row for insertion (maintain sorted order by date)
    targetRow = FindInsertRowForDate(ws, CDate(deliveryDate))
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    ' Shift existing rows down if needed
    ShiftRowsDownFrom ws, targetRow
    
    ' Write the new delivery data with day abbreviation (date + more spaces + day)
    ' Use 5 spaces between date and weekday abbreviation for better spacing
    With ws
        .Range(COL_UPCOMING_DATE & targetRow).Value = Format(CDate(deliveryDate), "yyyy-mm-dd") & "     " & GetDayAbbreviation(CDate(deliveryDate))
        .Range(COL_UPCOMING_DATE & targetRow).HorizontalAlignment = xlCenter
        .Range(COL_H2O2_EXPECTED & targetRow).Value = h2o2Qty
        .Range(COL_NAOH_EXPECTED & targetRow).Value = naohQty
    End With
    
    ' Apply borders: medium left/right/bottom, thin inside vertical separators
    ApplyRowBorders ws, targetRow
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws
    
    MsgBox "Upcoming delivery added for " & Format(deliveryDate, "yyyy-mm-dd") & " (" & GetDayAbbreviation(CDate(deliveryDate)) & ")." & vbCrLf & _
           "H2O2: " & h2o2Qty & " totes" & vbCrLf & _
           "NAOH: " & naohQty & " totes", vbInformation
End Sub

' Remove upcoming delivery by date
Public Sub RemoveUpcomingDelivery()
    Dim ws As Worksheet
    Dim removeDate As Variant
    Dim removedCount As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & MAIN_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt for date to remove
    removeDate = Application.InputBox( _
        "Enter the delivery date to remove (e.g., 2026-01-15):" & vbCrLf & _
        "All deliveries for this date will be removed.", _
        "Remove Upcoming Delivery", _
        Format(Date, "yyyy-mm-dd"), _
        Type:=2)
    
    If VarType(removeDate) = vbBoolean And removeDate = False Then Exit Sub
    If Len(Trim(CStr(removeDate))) = 0 Then Exit Sub
    
    ' Validate and convert to date
    If Not IsDate(removeDate) Then
        MsgBox "Invalid date format. Please enter a valid date.", vbExclamation
        Exit Sub
    End If
    removeDate = CDate(removeDate)
    
    ' Remove all deliveries for this date
    removedCount = RemoveDeliveriesByDate(ws, CDate(removeDate))
    
    If removedCount > 0 Then
        MsgBox removedCount & " upcoming delivery(ies) removed for " & Format(removeDate, "yyyy-mm-dd") & ".", vbInformation
    Else
        MsgBox "No upcoming deliveries found for " & Format(removeDate, "yyyy-mm-dd") & ".", vbInformation
    End If
End Sub

' ===================== INTEGRATION FUNCTION =====================
' Called from InventoryEvents when a delivery is logged
' Returns True if a matching upcoming delivery was found and removed
Public Function TryRemoveMatchingUpcomingDelivery(chemID As String, deliveryQty As Long) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim todayDate As Date
    Dim cellDateValue As Variant
    Dim cellDateParsed As Date
    Dim expectedQty As Long
    Dim colToCheck As String
    
    TryRemoveMatchingUpcomingDelivery = False
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    ' Determine which column to check based on chemical
    If chemID = "H2O2" Then
        colToCheck = COL_H2O2_EXPECTED
    ElseIf chemID = "NAOH" Then
        colToCheck = COL_NAOH_EXPECTED
    Else
        Exit Function
    End If
    
    todayDate = Date
    lastRow = GetLastUpcomingRow(ws)
    
    If lastRow < UPCOMING_START_ROW Then Exit Function
    
    ' Search for matching delivery (same date AND same quantity for this chemical)
    For i = UPCOMING_START_ROW To lastRow
        cellDateValue = ws.Range(COL_UPCOMING_DATE & i).Value
        
        ' Parse the date from the cell (format: "yyyy-mm-dd     Day")
        cellDateParsed = ParseDateFromCell(CStr(cellDateValue))
        
        If cellDateParsed > 0 Then
            If Int(cellDateParsed) = Int(todayDate) Then
                expectedQty = 0
                On Error Resume Next
                expectedQty = CLng(ws.Range(colToCheck & i).Value)
                On Error GoTo 0
                
                If expectedQty = deliveryQty Then
                    ' Found a match! Check if we should remove entire row or just update
                    Dim otherCol As String
                    Dim otherQty As Long
                    
                    If colToCheck = COL_H2O2_EXPECTED Then
                        otherCol = COL_NAOH_EXPECTED
                    Else
                        otherCol = COL_H2O2_EXPECTED
                    End If
                    
                    otherQty = 0
                    On Error Resume Next
                    otherQty = CLng(ws.Range(otherCol & i).Value)
                    On Error GoTo 0
                    
                    ' === UNPROTECT before changes ===
                    UnprotectSheet ws
                    
                    If otherQty = 0 Then
                        ' No other chemical expected, remove entire row
                        ClearUpcomingRow ws, i
                        ShiftRowsUpFrom ws, i
                    Else
                        ' Other chemical still expected, just clear this one
                        ws.Range(colToCheck & i).Value = 0
                    End If
                    
                    ' === RE-PROTECT after changes ===
                    ProtectSheet ws
                    
                    TryRemoveMatchingUpcomingDelivery = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ===================== HELPER FUNCTIONS =====================

' Get abbreviated day name from a date
Private Function GetDayAbbreviation(dt As Date) As String
    Select Case Weekday(dt, vbSunday)
        Case 1: GetDayAbbreviation = "Sun"
        Case 2: GetDayAbbreviation = "Mon"
        Case 3: GetDayAbbreviation = "Tue"
        Case 4: GetDayAbbreviation = "Wed"
        Case 5: GetDayAbbreviation = "Thu"
        Case 6: GetDayAbbreviation = "Fri"
        Case 7: GetDayAbbreviation = "Sat"
    End Select
End Function

' Parse date from cell value that may contain "yyyy-mm-dd     Day" format
Private Function ParseDateFromCell(cellValue As String) As Date
    Dim datePart As String
    Dim spacePos As Long
    
    On Error GoTo ParseError
    
    If Len(Trim(cellValue)) = 0 Then
        ParseDateFromCell = 0
        Exit Function
    End If
    
    ' Find the first space (date is before it)
    spacePos = InStr(1, cellValue, " ")
    If spacePos > 0 Then
        datePart = Left$(cellValue, spacePos - 1)
    Else
        datePart = cellValue
    End If
    
    If IsDate(datePart) Then
        ParseDateFromCell = CDate(datePart)
    Else
        ParseDateFromCell = 0
    End If
    Exit Function
    
ParseError:
    ParseDateFromCell = 0
End Function

' Find the row where a new date should be inserted (maintains ascending sort)
Private Function FindInsertRowForDate(ws As Worksheet, newDate As Date) As Long
    Dim lastRow As Long, i As Long
    Dim cellDate As Date
    
    lastRow = GetLastUpcomingRow(ws)
    
    ' If no existing data, start at row 2
    If lastRow < UPCOMING_START_ROW Then
        FindInsertRowForDate = UPCOMING_START_ROW
        Exit Function
    End If
    
    ' Find where to insert (ascending order by date)
    For i = UPCOMING_START_ROW To lastRow
        cellDate = ParseDateFromCell(CStr(ws.Range(COL_UPCOMING_DATE & i).Value))
        If cellDate > 0 Then
            If newDate < cellDate Then
                FindInsertRowForDate = i
                Exit Function
            End If
        End If
    Next i
    
    ' New date is after all existing dates, add at end
    FindInsertRowForDate = lastRow + 1
End Function

' Get the last row with upcoming delivery data
Private Function GetLastUpcomingRow(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, COL_UPCOMING_DATE).End(xlUp).Row
    
    ' Make sure we're not counting the header
    If lastRow < UPCOMING_START_ROW Then
        GetLastUpcomingRow = UPCOMING_START_ROW - 1
    Else
        ' Verify the cell actually has data
        If IsEmpty(ws.Range(COL_UPCOMING_DATE & lastRow).Value) Or _
           Len(Trim(CStr(ws.Range(COL_UPCOMING_DATE & lastRow).Value))) = 0 Then
            GetLastUpcomingRow = UPCOMING_START_ROW - 1
        Else
            GetLastUpcomingRow = lastRow
        End If
    End If
End Function

' Shift rows down from a specific row to make room for insertion
Private Sub ShiftRowsDownFrom(ws As Worksheet, startRow As Long)
    Dim lastRow As Long, i As Long
    
    lastRow = GetLastUpcomingRow(ws)
    
    ' If startRow is beyond last row, no shifting needed
    If startRow > lastRow Then Exit Sub
    
    ' Shift from bottom to top to avoid overwriting
    For i = lastRow To startRow Step -1
        ' Copy values
        ws.Range(COL_UPCOMING_DATE & (i + 1)).Value = ws.Range(COL_UPCOMING_DATE & i).Value
        ws.Range(COL_UPCOMING_DATE & (i + 1)).HorizontalAlignment = xlCenter
        ws.Range(COL_H2O2_EXPECTED & (i + 1)).Value = ws.Range(COL_H2O2_EXPECTED & i).Value
        ws.Range(COL_NAOH_EXPECTED & (i + 1)).Value = ws.Range(COL_NAOH_EXPECTED & i).Value
        
        ' Apply borders to new location
        ApplyRowBorders ws, i + 1
    Next i
    
    ' Clear the original row (it will be overwritten with new data)
    ClearUpcomingRow ws, startRow
End Sub

' Shift rows up from a specific row after removal
Private Sub ShiftRowsUpFrom(ws As Worksheet, fromRow As Long)
    Dim lastRow As Long, i As Long
    
    lastRow = GetLastUpcomingRow(ws)
    
    ' If fromRow is at or beyond last row, just clear it
    If fromRow >= lastRow Then
        ClearUpcomingRow ws, fromRow
        Exit Sub
    End If
    
    ' Shift from top to bottom
    For i = fromRow To lastRow - 1
        ' Copy values from row below
        ws.Range(COL_UPCOMING_DATE & i).Value = ws.Range(COL_UPCOMING_DATE & (i + 1)).Value
        ws.Range(COL_UPCOMING_DATE & i).HorizontalAlignment = xlCenter
        ws.Range(COL_H2O2_EXPECTED & i).Value = ws.Range(COL_H2O2_EXPECTED & (i + 1)).Value
        ws.Range(COL_NAOH_EXPECTED & i).Value = ws.Range(COL_NAOH_EXPECTED & (i + 1)).Value
        
        ' Apply borders
        ApplyRowBorders ws, i
    Next i
    
    ' Clear the last row (now empty)
    ClearUpcomingRow ws, lastRow
End Sub

' Remove all deliveries for a specific date
Private Function RemoveDeliveriesByDate(ws As Worksheet, targetDate As Date) As Long
    Dim lastRow As Long, i As Long
    Dim cellDate As Date
    Dim removedCount As Long
    
    removedCount = 0
    
    ' === UNPROTECT before changes ===
    UnprotectSheet ws
    
    ' Process from bottom to top to handle shifting correctly
    lastRow = GetLastUpcomingRow(ws)
    
    i = UPCOMING_START_ROW
    Do While i <= GetLastUpcomingRow(ws)
        cellDate = ParseDateFromCell(CStr(ws.Range(COL_UPCOMING_DATE & i).Value))
        If cellDate > 0 Then
            If Int(cellDate) = Int(targetDate) Then
                ' Found a match, remove it
                ClearUpcomingRow ws, i
                ShiftRowsUpFrom ws, i
                removedCount = removedCount + 1
                ' Don't increment i, check the same row again (new data shifted here)
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    
    ' === RE-PROTECT after changes ===
    ProtectSheet ws
    
    RemoveDeliveriesByDate = removedCount
End Function

' Clear a single row of upcoming delivery data.
' IMPORTANT: do NOT remove the top outer border when clearing ? preserve any top edge.
Private Sub ClearUpcomingRow(ws As Worksheet, rowNum As Long)
    With ws.Range(COL_UPCOMING_DATE & rowNum & ":" & COL_NAOH_EXPECTED & rowNum)
        .ClearContents
        ' Remove left/right/bottom and inside vertical/horizontal separators but keep the top edge border
        On Error Resume Next
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        On Error GoTo 0
        ' Intentionally preserve .Borders(xlEdgeTop)
    End With
End Sub

' Apply borders to a row:
' - medium (thicker) border on left, right and bottom only (no top).
' - thin vertical separators between columns (so columns are visually separated).
' NOTE: This routine intentionally avoids clearing or touching xlEdgeTop to preserve the header bottom border.
Private Sub ApplyRowBorders(ws As Worksheet, rowNum As Long)
    Dim rng As Range
    Set rng = ws.Range(COL_UPCOMING_DATE & rowNum & ":" & COL_NAOH_EXPECTED & rowNum)
    
    ' Clear only the borders we will set (leave xlEdgeTop alone)
    On Error Resume Next
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlEdgeBottom).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    On Error GoTo 0
    
    ' Apply thin vertical separators between the columns (inside vertical)
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    
    ' Apply medium left edge
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    
    ' Apply medium right edge
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    
    ' Apply medium bottom edge
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(0, 0, 0)
    End With
    
End Sub





