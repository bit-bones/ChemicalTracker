VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Calculate()
    Const TRANSACTIONS_SHEET As String = "Transactions"
    Const TRANSACTIONS_TABLE As String = "tblTransactions"
    
    Dim wsTrans As Worksheet
    Dim lo As ListObject
    Dim dtRange As Range
    Dim latestDt As Variant
    
    On Error GoTo SafeExit
    
    Set wsTrans = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    Set lo = wsTrans.ListObjects(TRANSACTIONS_TABLE)
    
    If lo.DataBodyRange Is Nothing Then
        SetLastUpdatedCell Empty
        GoTo SafeExit
    End If
    
    On Error Resume Next
    Set dtRange = wsTrans.Range("tblTransactions[DateTime]")
    On Error GoTo SafeExit
    
    If dtRange Is Nothing Then
        GoTo SafeExit
    End If
    
    latestDt = Application.WorksheetFunction.Max(dtRange)
    
    If Not CellsEqual(Me.Range("A2").Value, latestDt) Then
        Application.EnableEvents = False
        UnprotectSheet Me
        SetLastUpdatedCell latestDt
        ProtectSheet Me
        Application.EnableEvents = True
    End If
    
SafeExit:
    On Error Resume Next
    Application.EnableEvents = True
End Sub

Private Sub SetLastUpdatedCell(val As Variant)
    If IsEmpty(val) Then
        Me.Range("A2").ClearContents
    Else
        Me.Range("A2").Value = val
    End If
End Sub

Private Function CellsEqual(v1 As Variant, v2 As Variant) As Boolean
    If IsDate(v1) And IsDate(v2) Then
        CellsEqual = (CDbl(CDate(v1)) = CDbl(CDate(v2)))
    Else
        CellsEqual = (v1 = v2)
    End If
End Function

