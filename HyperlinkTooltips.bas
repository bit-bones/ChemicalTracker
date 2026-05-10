Attribute VB_Name = "HyperlinkTooltips"

Option Explicit

Public Sub SetupInfoTooltips()
    Dim ws As Worksheet
    Dim count As Long
    
    count = 0
    
    For Each ws In ThisWorkbook.Worksheets
        count = count + SetupInfoTooltipsForSheet(ws)
    Next ws
    
    CleanupLegacyHyperlinkArtifacts
    
    MsgBox count & " info tooltips configured!" & vbCrLf & vbCrLf & _
           "- Hover over (i) icon = instant tooltip" & vbCrLf & _
           "- Click button = runs macro", vbInformation
End Sub

Private Function SetupInfoTooltipsForSheet(ws As Worksheet) As Long
    Dim shp As Shape
    Dim tipText As String
    Dim count As Long
    
    UnprotectSheet ws
    CleanupMainButtons ws
    
    For Each shp In ws.Shapes
        tipText = GetTooltipForInfoIcon(shp.Name)
        
        If Len(tipText) > 0 Then
            ' Remove any existing hyperlink
            On Error Resume Next
            shp.Hyperlink.Delete
            On Error GoTo 0
            
            ' Add hyperlink with tooltip - points to same cell (does nothing on click)
            ws.Hyperlinks.Add _
                Anchor:=shp, _
                Address:="", _
                SubAddress:="", _
                ScreenTip:=tipText
            
            count = count + 1
        End If
    Next shp
    
    ProtectSheet ws
    SetupInfoTooltipsForSheet = count
End Function

Private Sub CleanupMainButtons(ws As Worksheet)
    Dim shp As Shape
    Dim macroName As String
    
    ' Remove hyperlinks from main buttons and restore OnAction
    For Each shp In ws.Shapes
        macroName = GetMacroForButton(shp.Name)
        
        If Len(macroName) > 0 Then
            On Error Resume Next
            shp.Hyperlink.Delete
            On Error GoTo 0
            shp.OnAction = macroName
        End If
    Next shp
End Sub

Private Sub CleanupLegacyHyperlinkArtifacts()
    Dim nm As Name

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("HyperlinkHelper").Delete
    Application.DisplayAlerts = True

    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 4) = "BTN_" Then
            nm.Delete
        End If
    Next nm
    On Error GoTo 0
End Sub

Private Function GetMacroForButton(shapeName As String) As String
    Select Case shapeName
        Case "btnUndoMyLastTransaction":  GetMacroForButton = "UndoMyLastTransaction"
        Case "btnManualSnapshot": GetMacroForButton = "ManualSnapshot"
        Case "btnDeliveryH2O2": GetMacroForButton = "Delivery_H2O2"
        Case "btnEmptyH2O2": GetMacroForButton = "Empty_H2O2"
        Case "btnDeliveryNaOH": GetMacroForButton = "Delivery_NaOH"
        Case "btnOpenNaOH":  GetMacroForButton = "Open_NaOH"
        Case "btnEmptyNaOH": GetMacroForButton = "Empty_NaOH"
        Case "btnFixCountsAddAdjustments": GetMacroForButton = "FixCounts_AddAdjustments"
        Case "btnAddUpcomingDelivery": GetMacroForButton = "AddUpcomingDelivery"
        Case "btnRemoveUpcomingDelivery": GetMacroForButton = "RemoveUpcomingDelivery"
        Case Else: GetMacroForButton = ""
    End Select
End Function

Private Function GetTooltipForInfoIcon(iconName As String) As String
    Select Case iconName
        Case "Info_Undo"
            GetTooltipForInfoIcon = " - UNDO - " & vbLf & _
                                    "Remove your most recent transaction."

        Case "Info_ManualSnapshot"
            GetTooltipForInfoIcon = " - MANUAL SNAPSHOT - " & vbLf & _
                                    "Create today's daily snapshot without logging a transaction."

        Case "Info_H2o2Delivery"
            GetTooltipForInfoIcon = " - H2O2 DELIVERY - " & vbLf & _
                                    "Receive full hydrogen peroxide totes and remove empties picked up."
        
        Case "Info_H2o2Empty"
            GetTooltipForInfoIcon = " - H2O2 EMPTY - " & vbLf & _
                                    "Convert full hydrogen peroxide totes to empty."
        
        Case "Info_NaohDelivery"
            GetTooltipForInfoIcon = " - NaOH DELIVERY - " & vbLf & _
                                    "Receive full sodium hydroxide totes and remove empties picked up."
        
        Case "Info_NaohPartial"
            GetTooltipForInfoIcon = " - NaOH PARTIAL - " & vbLf & _
                                    "Convert full sodium hydroxide totes to partial."
        
        Case "Info_NaohEmpty"
            GetTooltipForInfoIcon = " - NaOH EMPTY - " & vbLf & _
                                    "Convert partial sodium hydroxide totes to empty."
        
        Case "Info_FixCounts"
            GetTooltipForInfoIcon = " - FIX COUNTS - " & vbLf & _
                                    "Manually adjust inventory counts by entering the exact quantity of totes in stock." & vbLf & _
                                    "Only to be used if inventory does not match what has physically been counted."
        
        Case "Info_AddDelivery"
            GetTooltipForInfoIcon = " - ADD DELIVERY - " & vbLf & _
                                    "Add an upcoming delivery by entering date and totes expected." & vbLf & _
                                    "The list will automatically sort soonest to latest (top to bottom)."
        
        Case "Info_RemoveDelivery"
            GetTooltipForInfoIcon = " - REMOVE DELIVERY - " & vbLf & _
                                    "Remove delivery(s) by entering their scheduled date." & vbLf & _
                                    "Otherwise, delivery transactions should automatically" & vbLf & _
                                    "remove the upcoming delivery if date and quantity match."
        
        Case Else
            GetTooltipForInfoIcon = ""
    End Select
End Function

Public Sub RemoveInfoTooltips()
    Dim ws As Worksheet
    Dim count As Long
    
    count = 0
    
    For Each ws In ThisWorkbook.Worksheets
        count = count + RemoveInfoTooltipsForSheet(ws)
    Next ws
    
    MsgBox count & " info tooltips removed.", vbInformation
End Sub

Private Function RemoveInfoTooltipsForSheet(ws As Worksheet) As Long
    Dim shp As Shape
    Dim count As Long
    
    UnprotectSheet ws
    
    ' Remove hyperlinks from info icons
    For Each shp In ws.Shapes
        If Left(shp.Name, 5) = "Info_" Then
            On Error Resume Next
            shp.Hyperlink.Delete
            On Error GoTo 0
            count = count + 1
        End If
    Next shp
    
    CleanupMainButtons ws
    ProtectSheet ws
    
    RemoveInfoTooltipsForSheet = count
End Function

Public Sub DiagnoseShapes()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim msg As String
    Dim filePath As String
    Dim fileNum As Integer
    
    msg = "SHAPE DIAGNOSIS" & vbCrLf
    msg = msg & "Date: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    msg = msg & "================================================" & vbCrLf & vbCrLf
    
    For Each ws In ThisWorkbook.Worksheets
        msg = msg & "Worksheet: " & ws.Name & vbCrLf
        For Each shp In ws.Shapes
            If Left(shp.Name, 3) = "btn" Or Left(shp.Name, 5) = "Info_" Then
                msg = msg & "  Shape: " & shp.Name & vbCrLf
                msg = msg & "    OnAction: [" & shp.OnAction & "]" & vbCrLf

                On Error Resume Next
                Dim hlTip As String
                hlTip = shp.Hyperlink.ScreenTip
                If Err.Number = 0 Then
                    msg = msg & "    Hyperlink ScreenTip: [" & Left(hlTip, 40) & "... ]" & vbCrLf
                Else
                    msg = msg & "    Hyperlink:  NONE" & vbCrLf
                End If
                Err.Clear
                On Error GoTo 0

                msg = msg & vbCrLf
            End If
        Next shp
    Next ws
    
    filePath = Environ("USERPROFILE") & "\Desktop\ShapeDiagnosis.txt"
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, msg
    Close #fileNum
    
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
End Sub



