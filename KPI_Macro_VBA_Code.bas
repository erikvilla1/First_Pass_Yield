Option Explicit

' =============================================================================
' MANUFACTURING KPI FAILURE ANALYSIS - PURE VBA (NO PYTHON REQUIRED)
' =============================================================================
' Works for anyone - just export data from Power BI and run this macro!
' No installation, no paths, no Python needed.
' =============================================================================

' Configuration - Minimum sample thresholds
Private Const THRESHOLD_WORST_PARTS As Long = 50
Private Const THRESHOLD_PART_REV_STAGE As Long = 30
Private Const THRESHOLD_STATIONS As Long = 100
Private Const THRESHOLD_STATION_STAGE As Long = 50
Private Const THRESHOLD_HOTSPOTS As Long = 20
Private Const THRESHOLD_DEPT_VIEW As Long = 20
Private Const THRESHOLD_RECENT_RISK As Long = 10
Private Const RECENT_DAYS As Long = 60

' Column indices (will be set dynamically)
Private colPartNo As Long
Private colRevision As Long
Private colSerialNo As Long
Private colProductFamily As Long
Private colFF As Long
Private colDateTested As Long
Private colPassFail As Long
Private colProcessStage As Long
Private colStation As Long

Public Sub Generate_KPI_Report()

    Dim inputFile As Variant
    Dim outputFile As Variant
    Dim wbInput As Workbook
    Dim wbOutput As Workbook
    Dim wsData As Worksheet
    Dim lastRow As Long, lastCol As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler

    ' Pick input Excel file
    inputFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xls;*.xlsm), *.xlsx;*.xls;*.xlsm", _
        Title:="Select the Power BI Failure Data Export")

    If inputFile = False Then
        MsgBox "No file selected. Operation cancelled.", vbInformation
        GoTo Cleanup
    End If

    ' Pick output location
    outputFile = Application.GetSaveAsFilename( _
        InitialFileName:=Replace(CStr(inputFile), ".xlsx", "_KPI_Report.xlsx"), _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Save KPI Report As")

    If outputFile = False Then
        MsgBox "No output file chosen. Operation cancelled.", vbInformation
        GoTo Cleanup
    End If

    Application.StatusBar = "Loading data..."

    ' Open input file
    Set wbInput = Workbooks.Open(CStr(inputFile), ReadOnly:=True)
    Set wsData = wbInput.Sheets(1)

    ' Find data range
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "No data found in the file!", vbExclamation
        wbInput.Close False
        GoTo Cleanup
    End If

    ' Map column indices
    If Not MapColumns(wsData) Then
        MsgBox "Required columns not found! Make sure the file has:" & vbCrLf & _
               "Part No, Pass or Fail, Date Tested" & vbCrLf & vbCrLf & _
               "Optional: Revision, Serial No, Product Family, FF (Focus Factory), Process Stage, Station" & vbCrLf & vbCrLf & _
               "Found columns:" & vbCrLf & _
               "Part No: " & colPartNo & ", Pass or Fail: " & colPassFail & ", Date Tested: " & colDateTested, vbCritical
        wbInput.Close False
        GoTo Cleanup
    End If

    ' Create output workbook
    Set wbOutput = Workbooks.Add

    Application.StatusBar = "Generating KPI Report..."

    ' Generate all KPI sheets with error handling for each
    On Error Resume Next
    
    GenerateSummarySheet wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in Summary: " & Err.Description: Err.Clear
    
    GenerateWorstParts wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in WorstParts: " & Err.Description: Err.Clear
    
    GenerateWorstPartRevStage wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in PartRevStage: " & Err.Description: Err.Clear
    
    GenerateWorstStations wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in Stations: " & Err.Description: Err.Clear
    
    GenerateWorstStationStage wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in StationStage: " & Err.Description: Err.Clear
    
    GenerateHotspots wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in Hotspots: " & Err.Description: Err.Clear

    GenerateDepartmentView wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in DeptView: " & Err.Description: Err.Clear

    GenerateRecentRisk wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in RecentRisk: " & Err.Description: Err.Clear
    
    GenerateHighRiskWatchlist wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in Watchlist: " & Err.Description: Err.Clear

    GenerateFFSummary wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in FFSummary: " & Err.Description: Err.Clear

    GenerateDeptFailureSource wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in DeptFailure: " & Err.Description: Err.Clear

    ' NEW: Hierarchical drill-down by Focus Factory
    GenerateFF_ProcessStage wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in FF_ProcessStage: " & Err.Description: Err.Clear
    
    GenerateFF_Stage_Station wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in FF_Stage_Station: " & Err.Description: Err.Clear
    
    GenerateFF_Stage_Station_Part wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in FF_Full_Drilldown: " & Err.Description: Err.Clear

    GenerateFFTopFailingStationStage wbOutput, wsData, lastRow
    If Err.Number <> 0 Then Debug.Print "Error in TopStationStage: " & Err.Description: Err.Clear
    
    On Error GoTo ErrHandler

    ' Delete default empty sheets
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In wbOutput.Worksheets
        If ws.Name Like "Sheet*" And ws.UsedRange.Cells.Count <= 1 Then
            ws.Delete
        End If
    Next ws
    On Error GoTo ErrHandler

    ' Save output
    Application.StatusBar = "Saving report..."
    wbOutput.SaveAs CStr(outputFile), xlOpenXMLWorkbook

    ' Close input
    wbInput.Close False

    Application.StatusBar = False
    MsgBox "KPI Report generated successfully!" & vbCrLf & vbCrLf & CStr(outputFile), vbInformation, "Success"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Error occurred at: " & Erl, vbCritical, "Error"
    On Error Resume Next
    If Not wbInput Is Nothing Then wbInput.Close False
    Resume Cleanup
End Sub

' =============================================================================
' COLUMN MAPPING (UPDATED: FF = Focus Factory = Department)
' =============================================================================
Private Function MapColumns(ws As Worksheet) As Boolean
    Dim col As Long
    Dim header As String

    colPartNo = 0: colRevision = 0: colSerialNo = 0
    colProductFamily = 0: colFF = 0: colDateTested = 0
    colPassFail = 0: colProcessStage = 0: colStation = 0

    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = LCase(Trim(ws.Cells(1, col).Value))

        Select Case header
            Case "part no": colPartNo = col
            Case "revision": colRevision = col
            Case "serial no": colSerialNo = col

            ' Product Family is NOT department (entire product)
            Case "product family": colProductFamily = col

            ' Focus Factory is Department
            Case "ff", "focus factory", "focus_factory", "focusfactory": colFF = col

            Case "date tested": colDateTested = col
            Case "pass or fail", "pass/fail", "pass fail": colPassFail = col
            Case "process stage": colProcessStage = col
            Case "station": colStation = col
        End Select
    Next col

    MapColumns = (colPartNo > 0 And colPassFail > 0 And colDateTested > 0)
End Function

' =============================================================================
' SUMMARY SHEET
' =============================================================================
Private Sub GenerateSummarySheet(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim totalRecords As Long, totalFails As Long, totalPass As Long
    Dim minDate As Date, maxDate As Date
    Dim i As Long
    Dim passFailVal As String

    Set ws = wbOut.Worksheets.Add
    ws.Name = "0_Summary"

    totalRecords = lastRow - 1

    For i = 2 To lastRow
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))
        If passFailVal = "FAIL" Then
            totalFails = totalFails + 1
        ElseIf passFailVal = "PASS" Then
            totalPass = totalPass + 1
        End If

        If IsDate(wsData.Cells(i, colDateTested).Value) Then
            If minDate = 0 Or wsData.Cells(i, colDateTested).Value < minDate Then
                minDate = wsData.Cells(i, colDateTested).Value
            End If
            If wsData.Cells(i, colDateTested).Value > maxDate Then
                maxDate = wsData.Cells(i, colDateTested).Value
            End If
        End If
    Next i

    ws.Cells(1, 1).Value = "MANUFACTURING KPI REPORT"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 16

    ws.Cells(3, 1).Value = "Report Generated:"
    ws.Cells(3, 2).Value = Now

    ws.Cells(5, 1).Value = "=== DATA OVERVIEW ==="
    ws.Cells(6, 1).Value = "Total Records:"
    ws.Cells(6, 2).Value = totalRecords
    ws.Cells(7, 1).Value = "Total Fails:"
    ws.Cells(7, 2).Value = totalFails
    ws.Cells(8, 1).Value = "Total Pass:"
    ws.Cells(8, 2).Value = totalPass
    ws.Cells(9, 1).Value = "Overall Fail Rate:"
    If totalRecords > 0 Then
        ws.Cells(9, 2).Value = Format(totalFails / totalRecords, "0.00%")
    Else
        ws.Cells(9, 2).Value = "N/A"
    End If
    ws.Cells(10, 1).Value = "Overall Yield:"
    If totalRecords > 0 Then
        ws.Cells(10, 2).Value = Format(totalPass / totalRecords, "0.00%")
    Else
        ws.Cells(10, 2).Value = "N/A"
    End If

    ws.Cells(12, 1).Value = "Date Range:"
    ws.Cells(12, 2).Value = Format(minDate, "yyyy-mm-dd") & " to " & Format(maxDate, "yyyy-mm-dd")

    ws.Cells(14, 1).Value = "=== THRESHOLDS ==="
    ws.Cells(15, 1).Value = "Worst Parts: N >= " & THRESHOLD_WORST_PARTS
    ws.Cells(16, 1).Value = "Part-Rev-Stage: N >= " & THRESHOLD_PART_REV_STAGE
    ws.Cells(17, 1).Value = "Worst Stations: N >= " & THRESHOLD_STATIONS
    ws.Cells(18, 1).Value = "Station-Stage: N >= " & THRESHOLD_STATION_STAGE
    ws.Cells(19, 1).Value = "Hotspots: N >= " & THRESHOLD_HOTSPOTS
    ws.Cells(20, 1).Value = "Department View: N >= " & THRESHOLD_DEPT_VIEW
    ws.Cells(21, 1).Value = "Recent Risk (60 days): N >= " & THRESHOLD_RECENT_RISK

    ws.Columns("A:B").AutoFit
End Sub

' =============================================================================
' WORST PARTS
' =============================================================================
Private Sub GenerateWorstParts(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant
    Dim partNo As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim arr As Variant

    Application.StatusBar = "Generating Worst Parts..."

    Set dict = CreateObject("Scripting.Dictionary")

    ' Collect data
    For i = 2 To lastRow
        partNo = Trim(wsData.Cells(i, colPartNo).Value)
        If partNo = "" Then partNo = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        If Not dict.Exists(partNo) Then
            dict(partNo) = Array(0, 0) ' tests, fails
        End If

        arr = dict(partNo)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(partNo) = arr
    Next i

    ' Filter by threshold and build results
    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 5)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_WORST_PARTS Then
            resultCount = resultCount + 1
            results(resultCount, 1) = key
            results(resultCount, 2) = tests
            results(resultCount, 3) = fails
            results(resultCount, 4) = tests - fails
            results(resultCount, 5) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    ' Sort by fail rate descending
    SortResults results, resultCount, 5

    ' Create sheet
    Set ws = wbOut.Worksheets.Add
    ws.Name = "1_Worst_Parts"

    ' Headers
    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Part No"
    ws.Cells(1, 3).Value = "Total_Tests"
    ws.Cells(1, 4).Value = "Total_Fails"
    ws.Cells(1, 5).Value = "Total_Pass"
    ws.Cells(1, 6).Value = "Fail_Rate_Pct"
    FormatHeader ws, 6

    ' Data
    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
    Next r

    ws.Columns("A:F").AutoFit
End Sub

' =============================================================================
' WORST PART-REV-STAGE
' =============================================================================
Private Sub GenerateWorstPartRevStage(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim partNo As String, revision As String, stage As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating Worst Part-Rev-Stage..."

    ' Check required columns
    If colRevision = 0 Or colProcessStage = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        partNo = Trim(wsData.Cells(i, colPartNo).Value)
        If partNo = "" Then partNo = "UNKNOWN"
        revision = Trim(wsData.Cells(i, colRevision).Value)
        If revision = "" Then revision = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = partNo & "|" & revision & "|" & stage

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 7)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_PART_REV_STAGE Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0) ' Part No
            results(resultCount, 2) = parts(1) ' Revision
            results(resultCount, 3) = parts(2) ' Stage
            results(resultCount, 4) = tests
            results(resultCount, 5) = fails
            results(resultCount, 6) = tests - fails
            results(resultCount, 7) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 7

    Set ws = wbOut.Worksheets.Add
    ws.Name = "2_Worst_Part_Rev_Stage"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Part No"
    ws.Cells(1, 3).Value = "Revision"
    ws.Cells(1, 4).Value = "Process Stage"
    ws.Cells(1, 5).Value = "Total_Tests"
    ws.Cells(1, 6).Value = "Total_Fails"
    ws.Cells(1, 7).Value = "Total_Pass"
    ws.Cells(1, 8).Value = "Fail_Rate_Pct"
    FormatHeader ws, 8

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
    Next r

    ws.Columns("A:H").AutoFit
End Sub

' =============================================================================
' WORST STATIONS
' =============================================================================
Private Sub GenerateWorstStations(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant
    Dim station As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim arr As Variant

    Application.StatusBar = "Generating Worst Stations..."

    ' Check required columns
    If colStation = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        If Not dict.Exists(station) Then
            dict(station) = Array(0, 0)
        End If

        arr = dict(station)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(station) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 5)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_STATIONS Then
            resultCount = resultCount + 1
            results(resultCount, 1) = key
            results(resultCount, 2) = tests
            results(resultCount, 3) = fails
            results(resultCount, 4) = tests - fails
            results(resultCount, 5) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 5

    Set ws = wbOut.Worksheets.Add
    ws.Name = "3_Worst_Stations"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Station"
    ws.Cells(1, 3).Value = "Total_Tests"
    ws.Cells(1, 4).Value = "Total_Fails"
    ws.Cells(1, 5).Value = "Total_Pass"
    ws.Cells(1, 6).Value = "Fail_Rate_Pct"
    FormatHeader ws, 6

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
    Next r

    ws.Columns("A:F").AutoFit
End Sub

' =============================================================================
' WORST STATION-STAGE
' =============================================================================
Private Sub GenerateWorstStationStage(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim station As String, stage As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating Worst Station-Stage..."

    ' Check required columns
    If colStation = 0 Or colProcessStage = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = station & "|" & stage

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 6)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_STATION_STAGE Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0)
            results(resultCount, 2) = parts(1)
            results(resultCount, 3) = tests
            results(resultCount, 4) = fails
            results(resultCount, 5) = tests - fails
            results(resultCount, 6) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 6

    Set ws = wbOut.Worksheets.Add
    ws.Name = "4_Worst_Station_Stage"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Station"
    ws.Cells(1, 3).Value = "Process Stage"
    ws.Cells(1, 4).Value = "Total_Tests"
    ws.Cells(1, 5).Value = "Total_Fails"
    ws.Cells(1, 6).Value = "Total_Pass"
    ws.Cells(1, 7).Value = "Fail_Rate_Pct"
    FormatHeader ws, 7

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
    Next r

    ws.Columns("A:G").AutoFit
End Sub

' =============================================================================
' HOTSPOTS
' =============================================================================
Private Sub GenerateHotspots(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim station As String, partNo As String, stage As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating Hotspots..."

    ' Check required columns
    If colStation = 0 Or colProcessStage = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"
        partNo = Trim(wsData.Cells(i, colPartNo).Value)
        If partNo = "" Then partNo = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = station & "|" & partNo & "|" & stage

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 7)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_HOTSPOTS Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0)
            results(resultCount, 2) = parts(1)
            results(resultCount, 3) = parts(2)
            results(resultCount, 4) = tests
            results(resultCount, 5) = fails
            results(resultCount, 6) = tests - fails
            results(resultCount, 7) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 7

    Set ws = wbOut.Worksheets.Add
    ws.Name = "5_Hotspots"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Station"
    ws.Cells(1, 3).Value = "Part No"
    ws.Cells(1, 4).Value = "Process Stage"
    ws.Cells(1, 5).Value = "Total_Tests"
    ws.Cells(1, 6).Value = "Total_Fails"
    ws.Cells(1, 7).Value = "Total_Pass"
    ws.Cells(1, 8).Value = "Fail_Rate_Pct"
    FormatHeader ws, 8

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
    Next r

    ws.Columns("A:H").AutoFit
End Sub

' =============================================================================
' DEPARTMENT VIEW (UPDATED: uses Focus Factory as Department)
' =============================================================================
Private Sub GenerateDepartmentView(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim deptFF As String, partNo As String, revision As String, stage As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating Department View (Focus Factory)..."

    ' Focus Factory is Department
    If colFF = 0 Or colRevision = 0 Or colProcessStage = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        deptFF = Trim(wsData.Cells(i, colFF).Value)
        If deptFF = "" Then deptFF = "UNKNOWN"

        partNo = Trim(wsData.Cells(i, colPartNo).Value)
        If partNo = "" Then partNo = "UNKNOWN"

        revision = Trim(wsData.Cells(i, colRevision).Value)
        If revision = "" Then revision = "UNKNOWN"

        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"

        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = deptFF & "|" & partNo & "|" & revision & "|" & stage

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 8)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_DEPT_VIEW Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0) ' Focus Factory
            results(resultCount, 2) = parts(1) ' Part No
            results(resultCount, 3) = parts(2) ' Revision
            results(resultCount, 4) = parts(3) ' Stage
            results(resultCount, 5) = tests
            results(resultCount, 6) = fails
            results(resultCount, 7) = tests - fails
            results(resultCount, 8) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 8

    Set ws = wbOut.Worksheets.Add
    ws.Name = "6_FocusFactory_View"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory (Department)"
    ws.Cells(1, 3).Value = "Part No"
    ws.Cells(1, 4).Value = "Revision"
    ws.Cells(1, 5).Value = "Process Stage"
    ws.Cells(1, 6).Value = "Total_Tests"
    ws.Cells(1, 7).Value = "Total_Fails"
    ws.Cells(1, 8).Value = "Total_Pass"
    ws.Cells(1, 9).Value = "Fail_Rate_Pct"
    FormatHeader ws, 9

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
        ws.Cells(r + 1, 9).Value = results(r, 8)
    Next r

    ws.Columns("A:I").AutoFit
End Sub

' =============================================================================
' RECENT RISK
' =============================================================================
Private Sub GenerateRecentRisk(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim partNo As String, revision As String, stage As String, station As String, passFailVal As String
    Dim testDate As Date, maxDate As Date, cutoffDate As Date
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim riskLevel As String
    Dim failRate As Double
    Dim arr As Variant

    Application.StatusBar = "Generating Recent Risk..."

    ' Check required columns
    If colRevision = 0 Or colProcessStage = 0 Or colStation = 0 Then Exit Sub

    ' Find max date
    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, colDateTested).Value) Then
            testDate = wsData.Cells(i, colDateTested).Value
            If testDate > maxDate Then maxDate = testDate
        End If
    Next i

    cutoffDate = maxDate - RECENT_DAYS

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, colDateTested).Value) Then
            testDate = wsData.Cells(i, colDateTested).Value
            If testDate >= cutoffDate Then
                partNo = Trim(wsData.Cells(i, colPartNo).Value)
                If partNo = "" Then partNo = "UNKNOWN"
                revision = Trim(wsData.Cells(i, colRevision).Value)
                If revision = "" Then revision = "UNKNOWN"
                stage = Trim(wsData.Cells(i, colProcessStage).Value)
                If stage = "" Then stage = "UNKNOWN"
                station = Trim(wsData.Cells(i, colStation).Value)
                If station = "" Then station = "UNKNOWN"
                passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

                keyStr = partNo & "|" & revision & "|" & stage & "|" & station

                If Not dict.Exists(keyStr) Then
                    dict(keyStr) = Array(0, 0)
                End If

                arr = dict(keyStr)
                arr(0) = arr(0) + 1
                If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
                dict(keyStr) = arr
            End If
        End If
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 9)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_RECENT_RISK Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            failRate = Round(fails / tests * 100, 2)

            If failRate >= 50 Then
                riskLevel = "CRITICAL"
            ElseIf failRate >= 30 Then
                riskLevel = "HIGH"
            ElseIf failRate >= 15 Then
                riskLevel = "MEDIUM"
            Else
                riskLevel = "LOW"
            End If

            results(resultCount, 1) = parts(0)
            results(resultCount, 2) = parts(1)
            results(resultCount, 3) = parts(2)
            results(resultCount, 4) = parts(3)
            results(resultCount, 5) = tests
            results(resultCount, 6) = fails
            results(resultCount, 7) = tests - fails
            results(resultCount, 8) = failRate
            results(resultCount, 9) = riskLevel
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 8

    Set ws = wbOut.Worksheets.Add
    ws.Name = "7_Recent_Risk_60d"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Part No"
    ws.Cells(1, 3).Value = "Revision"
    ws.Cells(1, 4).Value = "Process Stage"
    ws.Cells(1, 5).Value = "Station"
    ws.Cells(1, 6).Value = "Total_Tests"
    ws.Cells(1, 7).Value = "Total_Fails"
    ws.Cells(1, 8).Value = "Total_Pass"
    ws.Cells(1, 9).Value = "Fail_Rate_Pct"
    ws.Cells(1, 10).Value = "Risk_Level"
    FormatHeader ws, 10

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
        ws.Cells(r + 1, 9).Value = results(r, 8)
        ws.Cells(r + 1, 10).Value = results(r, 9)

        ' Color code risk level
        Select Case results(r, 9)
            Case "CRITICAL"
                ws.Cells(r + 1, 10).Interior.Color = RGB(255, 107, 107)
            Case "HIGH"
                ws.Cells(r + 1, 10).Interior.Color = RGB(255, 165, 0)
            Case "MEDIUM"
                ws.Cells(r + 1, 10).Interior.Color = RGB(255, 255, 0)
        End Select
    Next r

    ws.Columns("A:J").AutoFit
End Sub

Private Sub GenerateHighRiskWatchlist(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    ' Placeholder: can be expanded later with a risk-score formula if desired.
End Sub

' =============================================================================
' FF SUMMARY (Focus Factory breakdown)
' =============================================================================
Private Sub GenerateFFSummary(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant
    Dim ff As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim arr As Variant

    Application.StatusBar = "Generating FF Summary..."

    If colFF = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        ff = Trim(wsData.Cells(i, colFF).Value)
        If ff = "" Then ff = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        If Not dict.Exists(ff) Then
            dict(ff) = Array(0, 0)
        End If

        arr = dict(ff)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(ff) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 5)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        resultCount = resultCount + 1
        results(resultCount, 1) = key
        results(resultCount, 2) = tests
        results(resultCount, 3) = fails
        results(resultCount, 4) = tests - fails
        results(resultCount, 5) = Round(fails / tests * 100, 2)
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 5

    Set ws = wbOut.Worksheets.Add
    ws.Name = "8_FF_Summary"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory"
    ws.Cells(1, 3).Value = "Total_Tests"
    ws.Cells(1, 4).Value = "Total_Fails"
    ws.Cells(1, 5).Value = "Total_Pass"
    ws.Cells(1, 6).Value = "Fail_Rate_Pct"
    FormatHeader ws, 6

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
    Next r

    ws.Columns("A:F").AutoFit
End Sub

' =============================================================================
' DEPT FAILURE SOURCE (UPDATED: uses FF as Department)
' =============================================================================
Private Sub GenerateDeptFailureSource(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dictDept As Object
    Dim i As Long, r As Long
    Dim key As Variant
    Dim deptFF As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim arr As Variant

    Application.StatusBar = "Generating Department Summary (Focus Factory)..."

    If colFF = 0 Then Exit Sub

    Set dictDept = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        deptFF = Trim(wsData.Cells(i, colFF).Value)
        If deptFF = "" Then deptFF = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        If Not dictDept.Exists(deptFF) Then
            dictDept(deptFF) = Array(0, 0)
        End If

        arr = dictDept(deptFF)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dictDept(deptFF) = arr
    Next i

    If dictDept.Count = 0 Then Exit Sub
    ReDim results(1 To dictDept.Count, 1 To 5)
    resultCount = 0

    For Each key In dictDept.Keys
        arr = dictDept(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= 50 Then
            resultCount = resultCount + 1
            results(resultCount, 1) = key
            results(resultCount, 2) = tests
            results(resultCount, 3) = fails
            results(resultCount, 4) = tests - fails
            results(resultCount, 5) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 5

    Set ws = wbOut.Worksheets.Add
    ws.Name = "9_FocusFactory_Summary"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory (Department)"
    ws.Cells(1, 3).Value = "Total_Tests"
    ws.Cells(1, 4).Value = "Total_Fails"
    ws.Cells(1, 5).Value = "Total_Pass"
    ws.Cells(1, 6).Value = "Fail_Rate_Pct"
    FormatHeader ws, 6

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
    Next r

    ws.Columns("A:F").AutoFit
End Sub

' =============================================================================
' NEW: TOP FAILING STATION + PROCESS STAGE PER FOCUS FACTORY (BY HIGHEST FAIL %)
' =============================================================================
Private Sub GenerateFFTopFailingStationStage(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dictAgg As Object
    Dim dictBest As Object
    Dim i As Long

    Dim deptFF As String, station As String, stage As String, passFailVal As String
    Dim keyStr As String, deptKey As String
    Dim arr As Variant

    Application.StatusBar = "Generating Top Failing Station+Stage per Focus Factory..."

    ' Must have Focus Factory (Department), Station, and Process Stage
    If colFF = 0 Or colStation = 0 Or colProcessStage = 0 Then Exit Sub

    Set dictAgg = CreateObject("Scripting.Dictionary")  ' Dept|Station|Stage -> tests,fails
    Set dictBest = CreateObject("Scripting.Dictionary") ' Dept -> bestStation, bestStage, tests, fails, failRate

    ' Aggregate Dept|Station|Stage
    For i = 2 To lastRow
        deptFF = Trim(wsData.Cells(i, colFF).Value)
        If deptFF = "" Then deptFF = "UNKNOWN"

        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"

        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"

        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = deptFF & "|" & station & "|" & stage

        If Not dictAgg.Exists(keyStr) Then
            dictAgg(keyStr) = Array(0, 0) ' tests, fails
        End If

        arr = dictAgg(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dictAgg(keyStr) = arr
    Next i

    ' Find best Station+Stage per Dept by FAIL % (tie-break: fails desc, tests desc)
    Dim key As Variant, parts() As String
    Dim tests As Long, fails As Long
    Dim failRate As Double
    Dim bestArr As Variant

    For Each key In dictAgg.Keys
        arr = dictAgg(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= THRESHOLD_DEPT_VIEW And fails > 0 Then
            parts = Split(CStr(key), "|")
            deptKey = parts(0)
            station = parts(1)
            stage = parts(2)

            failRate = Round((fails / tests) * 100, 2)

            If Not dictBest.Exists(deptKey) Then
                dictBest(deptKey) = Array(station, stage, tests, fails, failRate)
            Else
                bestArr = dictBest(deptKey)

                If failRate > bestArr(4) _
                   Or (failRate = bestArr(4) And fails > bestArr(3)) _
                   Or (failRate = bestArr(4) And fails = bestArr(3) And tests > bestArr(2)) Then
                    dictBest(deptKey) = Array(station, stage, tests, fails, failRate)
                End If
            End If
        End If
    Next key

    If dictBest.Count = 0 Then Exit Sub

    ' Output (one row per department)
    Set ws = wbOut.Worksheets.Add
    ws.Name = "12_TopStationStage_by_FF"

    ws.Cells(1, 1).Value = "Focus Factory (Department)"
    ws.Cells(1, 2).Value = "Top Failing Station"
    ws.Cells(1, 3).Value = "Top Failing Process Stage"
    ws.Cells(1, 4).Value = "Total_Tests"
    ws.Cells(1, 5).Value = "Total_Fails"
    ws.Cells(1, 6).Value = "Total_Pass"
    ws.Cells(1, 7).Value = "Fail_Rate_Pct"
    FormatHeader ws, 7

    Dim r As Long: r = 2
    Dim dept As Variant

    For Each dept In dictBest.Keys
        bestArr = dictBest(dept)
        ws.Cells(r, 1).Value = dept
        ws.Cells(r, 2).Value = bestArr(0)                 ' Station
        ws.Cells(r, 3).Value = bestArr(1)                 ' Stage
        ws.Cells(r, 4).Value = bestArr(2)                 ' Tests
        ws.Cells(r, 5).Value = bestArr(3)                 ' Fails
        ws.Cells(r, 6).Value = bestArr(2) - bestArr(3)    ' Pass
        ws.Cells(r, 7).Value = bestArr(4)                 ' Fail Rate %
        r = r + 1
    Next dept

    ' Sort by Fail Rate desc (highest risk first) - only if we have data rows
    If r > 2 Then
        On Error Resume Next
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("G2:G" & r - 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange ws.Range("A1:G" & r - 1)
            .Header = xlYes
            .Apply
        End With
        On Error GoTo 0
    End If

    ws.Columns("A:G").AutoFit
End Sub

' =============================================================================
' FF DRILL-DOWN: Focus Factory -> Process Stage breakdown
' Shows which test type (HP/PT/BI/FT/SYS) is worst in each department
' =============================================================================
Private Sub GenerateFF_ProcessStage(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim ff As String, stage As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating FF + Process Stage breakdown..."

    If colFF = 0 Or colProcessStage = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        ff = Trim(wsData.Cells(i, colFF).Value)
        If ff = "" Then ff = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = ff & "|" & stage

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 6)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= 10 Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0) ' FF
            results(resultCount, 2) = parts(1) ' Stage
            results(resultCount, 3) = tests
            results(resultCount, 4) = fails
            results(resultCount, 5) = tests - fails
            results(resultCount, 6) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 6

    Set ws = wbOut.Worksheets.Add
    ws.Name = "10_FF_by_TestType"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory"
    ws.Cells(1, 3).Value = "Test Type (Process Stage)"
    ws.Cells(1, 4).Value = "Total_Tests"
    ws.Cells(1, 5).Value = "Total_Fails"
    ws.Cells(1, 6).Value = "Total_Pass"
    ws.Cells(1, 7).Value = "Fail_Rate_Pct"
    FormatHeader ws, 7

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
    Next r

    ws.Columns("A:G").AutoFit
End Sub

' =============================================================================
' FF DRILL-DOWN: Focus Factory -> Process Stage -> Station
' Shows which station is worst for each test type in each department
' =============================================================================
Private Sub GenerateFF_Stage_Station(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim ff As String, stage As String, station As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating FF + Stage + Station breakdown..."

    If colFF = 0 Or colProcessStage = 0 Or colStation = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        ff = Trim(wsData.Cells(i, colFF).Value)
        If ff = "" Then ff = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = ff & "|" & stage & "|" & station

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 7)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= 10 Then
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0) ' FF
            results(resultCount, 2) = parts(1) ' Stage
            results(resultCount, 3) = parts(2) ' Station
            results(resultCount, 4) = tests
            results(resultCount, 5) = fails
            results(resultCount, 6) = tests - fails
            results(resultCount, 7) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 7

    Set ws = wbOut.Worksheets.Add
    ws.Name = "11_FF_Stage_Station"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory"
    ws.Cells(1, 3).Value = "Test Type (Process Stage)"
    ws.Cells(1, 4).Value = "Station"
    ws.Cells(1, 5).Value = "Total_Tests"
    ws.Cells(1, 6).Value = "Total_Fails"
    ws.Cells(1, 7).Value = "Total_Pass"
    ws.Cells(1, 8).Value = "Fail_Rate_Pct"
    FormatHeader ws, 8

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
    Next r

    ws.Columns("A:H").AutoFit
End Sub

' =============================================================================
' FF DRILL-DOWN: Focus Factory -> Process Stage -> Station -> Part
' Full drill-down showing worst parts at each station/test type/department
' =============================================================================
Private Sub GenerateFF_Stage_Station_Part(wbOut As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim i As Long, r As Long
    Dim key As Variant, keyStr As String
    Dim ff As String, stage As String, station As String, partNo As String, passFailVal As String
    Dim tests As Long, fails As Long
    Dim results() As Variant
    Dim resultCount As Long
    Dim parts() As String
    Dim arr As Variant

    Application.StatusBar = "Generating FF + Stage + Station + Part breakdown..."

    If colFF = 0 Or colProcessStage = 0 Or colStation = 0 Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        ff = Trim(wsData.Cells(i, colFF).Value)
        If ff = "" Then ff = "UNKNOWN"
        stage = Trim(wsData.Cells(i, colProcessStage).Value)
        If stage = "" Then stage = "UNKNOWN"
        station = Trim(wsData.Cells(i, colStation).Value)
        If station = "" Then station = "UNKNOWN"
        partNo = Trim(wsData.Cells(i, colPartNo).Value)
        If partNo = "" Then partNo = "UNKNOWN"
        passFailVal = UCase(Trim(wsData.Cells(i, colPassFail).Value))

        keyStr = ff & "|" & stage & "|" & station & "|" & partNo

        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(0, 0)
        End If

        arr = dict(keyStr)
        arr(0) = arr(0) + 1
        If passFailVal = "FAIL" Then arr(1) = arr(1) + 1
        dict(keyStr) = arr
    Next i

    If dict.Count = 0 Then Exit Sub
    ReDim results(1 To dict.Count, 1 To 8)
    resultCount = 0

    For Each key In dict.Keys
        arr = dict(key)
        tests = arr(0)
        fails = arr(1)

        If tests >= 5 And fails > 0 Then  ' Lower threshold, must have failures
            resultCount = resultCount + 1
            parts = Split(key, "|")
            results(resultCount, 1) = parts(0) ' FF
            results(resultCount, 2) = parts(1) ' Stage
            results(resultCount, 3) = parts(2) ' Station
            results(resultCount, 4) = parts(3) ' Part No
            results(resultCount, 5) = tests
            results(resultCount, 6) = fails
            results(resultCount, 7) = tests - fails
            results(resultCount, 8) = Round(fails / tests * 100, 2)
        End If
    Next key

    If resultCount = 0 Then Exit Sub

    SortResults results, resultCount, 8

    Set ws = wbOut.Worksheets.Add
    ws.Name = "12_FF_Full_Drilldown"

    ws.Cells(1, 1).Value = "Rank"
    ws.Cells(1, 2).Value = "Focus Factory"
    ws.Cells(1, 3).Value = "Test Type (Process Stage)"
    ws.Cells(1, 4).Value = "Station"
    ws.Cells(1, 5).Value = "Part No"
    ws.Cells(1, 6).Value = "Total_Tests"
    ws.Cells(1, 7).Value = "Total_Fails"
    ws.Cells(1, 8).Value = "Total_Pass"
    ws.Cells(1, 9).Value = "Fail_Rate_Pct"
    FormatHeader ws, 9

    For r = 1 To resultCount
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = results(r, 1)
        ws.Cells(r + 1, 3).Value = results(r, 2)
        ws.Cells(r + 1, 4).Value = results(r, 3)
        ws.Cells(r + 1, 5).Value = results(r, 4)
        ws.Cells(r + 1, 6).Value = results(r, 5)
        ws.Cells(r + 1, 7).Value = results(r, 6)
        ws.Cells(r + 1, 8).Value = results(r, 7)
        ws.Cells(r + 1, 9).Value = results(r, 8)
    Next r

    ws.Columns("A:I").AutoFit
End Sub

' =============================================================================
' SORT + HEADER FORMATTING HELPERS
' =============================================================================
Private Sub SortResults(ByRef results() As Variant, ByVal count As Long, ByVal sortCol As Long)
    ' Simple bubble sort by sortCol (descending)
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim cols As Long

    If count < 2 Then Exit Sub

    cols = UBound(results, 2)

    For i = 1 To count - 1
        For j = i + 1 To count
            If results(j, sortCol) > results(i, sortCol) Then
                ' Swap rows
                For k = 1 To cols
                    temp = results(i, k)
                    results(i, k) = results(j, k)
                    results(j, k) = temp
                Next k
            End If
        Next j
    Next i
End Sub

Private Sub FormatHeader(ws As Worksheet, cols As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, cols))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
End Sub