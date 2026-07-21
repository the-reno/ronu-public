Attribute VB_Name = "RatesWorkbookSetup"
Option Explicit

' ============================================================================
' Rates analysis workbook setup
'
' Import this module together with Rates_Analysis_Engine.bas.
'
' Public macros:
'   CreateRatesWorkbook       Create or repair the workbook without deleting
'                             existing Inputs, Curve, or Frontier_Settings data.
'   ResetRatesOutputs         Clear calculated output sheets only.
'   RebuildRatesLayout        Reapply the standard layout without clearing data.
'
' User-maintained sheets:
'   Inputs
'   Curve
'   Frontier_Settings
'
' Calculated sheets:
'   Transactions
'   Daily_Accrual
'   Rolling_Results
'   Tenor_Analysis
'   Monthly_Returns
'   Portfolio_Analysis
'   Out_of_Sample
'   Chart_Data
'   Dashboard
'   Test_Results
'
' Curve headers:
'   Date | ON | 1M | 2M | 3M | 6M
'
' Rates are ordinary numbers. 4.31 means 4.31 percent.
' No cells are merged.
' ============================================================================

Private Const COLOR_NAVY As Long = 3809035
Private Const COLOR_INPUT As Long = 13434879
Private Const COLOR_PALE As Long = 16448250
Private Const COLOR_GREEN_PALE As Long = 15198183

Public Sub CreateRatesWorkbook()
    BuildRatesWorkbook True
End Sub

Public Sub EnsureRatesWorkbook()
    BuildRatesWorkbook False
End Sub

Public Sub RebuildRatesLayout()
    BuildRatesWorkbook False
End Sub

Public Sub ResetRatesOutputs()
    Dim sheetNames As Variant
    Dim item As Variant
    Dim ws As Worksheet

    sheetNames = OutputSheetNames()

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo Fail

    For Each item In sheetNames
        Set ws = GetOrCreateRatesSheet(CStr(item))
        ClearCalculatedSheet ws
    Next item

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub BuildRatesWorkbook(ByVal showMessage As Boolean)
    Dim oldCalculation As XlCalculation
    Dim stageName As String

    On Error GoTo Fail

    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stageName = "creating input sheets"
    SetupInputs
    SetupCurve
    SetupFrontierSettings

    stageName = "creating calculated sheets"
    SetupCalculatedSheet "Transactions", "Transaction Schedule", 14
    SetupCalculatedSheet "Daily_Accrual", "Daily Accrual Ledger", 20
    SetupCalculatedSheet "Rolling_Results", "Rolling Investment Results", 15
    SetupCalculatedSheet "Tenor_Analysis", "Daily Tenor Scenario Analysis", 31
    SetupCalculatedSheet "Monthly_Returns", "Monthly Economic Returns", 12
    SetupCalculatedSheet "Portfolio_Analysis", _
                         "Static-Sleeve Efficient Frontier", 43
    SetupCalculatedSheet "Out_of_Sample", _
                         "Out-of-Sample Frontier Validation", 22
    SetupCalculatedSheet "Chart_Data", "Chart Data", 35
    SetupCalculatedSheet "Dashboard", _
                         "Historical Cash Investment Analysis", 17
    SetupCalculatedSheet "Test_Results", "Model Validation Results", 4

    stageName = "ordering sheets"
    OrderRatesSheets
    ApplyTabColors

CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        MsgBox "Rates workbook is ready." & vbCrLf & _
               "Existing Inputs, Curve, and Frontier_Settings data were preserved." & _
               vbCrLf & "Replace the Curve data and run RunRatesAnalysis.", _
               vbInformation
    End If
    Exit Sub

Fail:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Workbook setup stopped during " & stageName & "." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub SetupInputs()
    Dim ws As Worksheet

    Set ws = GetOrCreateRatesSheet("Inputs")

    ws.Range("A1:H16").ClearFormats
    ws.Range("A1:H1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Rates Analysis Model - Inputs"
    RatesStyleTitle ws.Range("A1:H1")

    RatesWriteRow ws.Range("A4"), Array("Model Input", "Value")
    RatesStyleHeader ws.Range("A4:B4")

    ws.Range("A5").Value = "Analysis Start Date"
    ws.Range("A6").Value = "Analysis End Date"
    ws.Range("A7").Value = "Initial Notional ($)"
    ws.Range("A8").Value = "Frontier Weight Step (percentage points)"
    ws.Range("A9").Value = "Validation Split Date"
    ws.Range("A10").Value = "Day Count"
    ws.Range("A11").Value = "Rate Input Convention"
    ws.Range("A12").Value = "Rolling Convention"
    ws.Range("A13").Value = "Portfolio Convention"

    If Len(Trim$(CStr(ws.Range("B5").Value2))) = 0 Then
        ws.Range("B5").Value = DateSerial(2023, 1, 3)
    End If
    If Len(Trim$(CStr(ws.Range("B6").Value2))) = 0 Then
        ws.Range("B6").Value = DateSerial(2026, 7, 3)
    End If
    If Len(Trim$(CStr(ws.Range("B7").Value2))) = 0 Then
        ws.Range("B7").Value = 100000000#
    End If
    If Len(Trim$(CStr(ws.Range("B8").Value2))) = 0 Then
        ws.Range("B8").Value = 5#
    End If
    If Len(Trim$(CStr(ws.Range("B9").Value2))) = 0 Then
        ws.Range("B9").Value = DateSerial(2024, 12, 31)
    End If

    ws.Range("B10").Value = "ACT/360"
    ws.Range("B11").Value = "4.31 means 4.31 percent"
    ws.Range("B12").Value = _
        "Each new term starts from the prior actual maturity"
    ws.Range("B13").Value = _
        "Static sleeves; no implicit monthly rebalancing"

    RatesWriteRow ws.Range("D4"), Array("Instructions")
    RatesStyleHeader ws.Range("D4:H4")

    ws.Range("D5").Value = _
        "1. Replace Curve!A:F and retain Date, ON, 1M, 2M, 3M, 6M headers."
    ws.Range("D6").Value = _
        "2. Only Inputs!B5 through Inputs!B6 are processed."
    ws.Range("D7").Value = _
        "3. Enter rates as numbers such as 4.31, not Excel percentages."
    ws.Range("D8").Value = _
        "4. Frontier step 5 means five-percentage-point weights."
    ws.Range("D9").Value = _
        "5. The validation split separates estimation and test periods."
    ws.Range("D10").Value = _
        "6. Frontier constraints are maintained on Frontier_Settings."
    ws.Range("D11").Value = _
        "7. Run RunRatesAnalysis after changing dates, curve, or constraints."
    ws.Range("D12").Value = _
        "8. Rerunning the engine clears outputs but preserves all input data."

    ws.Range("B5:B9").Interior.Color = COLOR_INPUT
    ws.Range("B5:B9").Font.Color = RGB(0, 0, 255)
    ws.Range("B5:B6").NumberFormat = "mm/dd/yyyy"
    ws.Range("B9").NumberFormat = "mm/dd/yyyy"
    ws.Range("B7").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("B8").NumberFormat = "0.0"

    ws.Range("D5:H12").Interior.Color = COLOR_PALE
    ws.Range("D5:H12").WrapText = True

    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D:H").ColumnWidth = 18
    ws.Rows("5:13").RowHeight = 25
    ws.Rows(1).RowHeight = 28
End Sub

Private Sub SetupCurve()
    Dim ws As Worksheet

    Set ws = GetOrCreateRatesSheet("Curve")

    If FindRatesHeaderRow(ws) = 0 Then
        RatesWriteRow ws.Range("A1"), _
            Array("Date", "ON", "1M", "2M", "3M", "6M")
    End If

    RatesStyleHeader ws.Range("A1:F1")
    ws.Range("A2:F10000").Font.Color = RGB(0, 0, 255)
    ws.Columns("A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:F").NumberFormat = "0.0000"
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B:F").ColumnWidth = 12
    ws.Rows(1).RowHeight = 24
End Sub

Private Sub SetupFrontierSettings()
    Dim ws As Worksheet

    Set ws = GetOrCreateRatesSheet("Frontier_Settings")

    ws.Range("A1:F16").ClearFormats
    ws.Range("A1:F1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Efficient Frontier - Treasury Constraints"
    RatesStyleTitle ws.Range("A1:F1")

    RatesWriteRow ws.Range("A4"), _
        Array("Constraint / Setting", "Value", "Input Convention")
    RatesStyleHeader ws.Range("A4:C4")

    ws.Range("A5").Value = "Use Treasury Constraints"
    ws.Range("A6").Value = "Minimum Available Within 30 Days"
    ws.Range("A7").Value = "Minimum Available Within 60 Days"
    ws.Range("A8").Value = "Maximum 6M Allocation"
    ws.Range("A9").Value = "Maximum Single-Tenor Allocation"
    ws.Range("A10").Value = "Maximum Weighted-Average Maturity"
    ws.Range("A11").Value = "Minimum ON Allocation"

    If Len(Trim$(CStr(ws.Range("B5").Value2))) = 0 Then
        ws.Range("B5").Value = "YES"
    End If
    If Len(Trim$(CStr(ws.Range("B6").Value2))) = 0 Then
        ws.Range("B6").Value = 40#
    End If
    If Len(Trim$(CStr(ws.Range("B7").Value2))) = 0 Then
        ws.Range("B7").Value = 60#
    End If
    If Len(Trim$(CStr(ws.Range("B8").Value2))) = 0 Then
        ws.Range("B8").Value = 30#
    End If
    If Len(Trim$(CStr(ws.Range("B9").Value2))) = 0 Then
        ws.Range("B9").Value = 60#
    End If
    If Len(Trim$(CStr(ws.Range("B10").Value2))) = 0 Then
        ws.Range("B10").Value = 3#
    End If
    If Len(Trim$(CStr(ws.Range("B11").Value2))) = 0 Then
        ws.Range("B11").Value = 0#
    End If

    ws.Range("C5").Value = "YES or NO"
    ws.Range("C6").Value = "Number: 40 means 40 percent"
    ws.Range("C7").Value = "Number: 60 means 60 percent"
    ws.Range("C8").Value = "Number: 30 means 30 percent"
    ws.Range("C9").Value = "Number: 60 means 60 percent"
    ws.Range("C10").Value = "Months"
    ws.Range("C11").Value = "Number: 0 means 0 percent"

    ws.Range("B5:B11").Interior.Color = COLOR_INPUT
    ws.Range("B5:B11").Font.Color = RGB(0, 0, 255)
    ws.Range("B6:B9").NumberFormat = "0.0"
    ws.Range("B10").NumberFormat = "0.0"
    ws.Range("B11").NumberFormat = "0.0"

    ws.Range("E4:F4").Interior.Color = COLOR_NAVY
    ws.Range("E4").Value = "Model Design"
    RatesStyleHeader ws.Range("E4:F4")

    ws.Range("E5").Value = "Portfolio method"
    ws.Range("F5").Value = "Static tenor sleeves"
    ws.Range("E6").Value = "Primary return measure"
    ws.Range("F6").Value = "Incremental annual return vs ON"
    ws.Range("E7").Value = "Primary risk measure"
    ws.Range("F7").Value = "Aligned monthly earnings volatility"
    ws.Range("E8").Value = "Additional risk"
    ws.Range("F8").Value = "Common-sample reset volatility"
    ws.Range("E9").Value = "Downside benchmark"
    ws.Range("F9").Value = "Monthly performance vs ON"
    ws.Range("E10").Value = "Frontier validation"
    ws.Range("F10").Value = "Estimation period and out-of-sample period"

    ws.Range("E5:F10").Interior.Color = COLOR_GREEN_PALE
    ws.Range("E5:F10").WrapText = True

    ws.Columns("A").ColumnWidth = 42
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 31
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 42
    ws.Rows("5:11").RowHeight = 24
    ws.Rows(1).RowHeight = 28
End Sub

Private Sub SetupCalculatedSheet(ByVal sheetName As String, _
                                 ByVal titleText As String, _
                                 ByVal columnCount As Long)
    Dim ws As Worksheet

    Set ws = GetOrCreateRatesSheet(sheetName)

    ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount)).Interior.Color = _
        COLOR_NAVY
    ws.Range("A1").Value = titleText
    RatesStyleTitle ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount))
    ws.Rows(1).RowHeight = 28
    ws.Columns(1).ColumnWidth = 15
End Sub

Private Function InputSheetNames() As Variant
    InputSheetNames = Array("Inputs", "Curve", "Frontier_Settings")
End Function

Private Function OutputSheetNames() As Variant
    OutputSheetNames = Array( _
        "Transactions", _
        "Daily_Accrual", _
        "Rolling_Results", _
        "Tenor_Analysis", _
        "Monthly_Returns", _
        "Portfolio_Analysis", _
        "Out_of_Sample", _
        "Chart_Data", _
        "Dashboard", _
        "Test_Results")
End Function

Private Sub OrderRatesSheets()
    Dim sheetNames As Variant
    Dim index As Long

    sheetNames = Array( _
        "Inputs", _
        "Curve", _
        "Frontier_Settings", _
        "Transactions", _
        "Daily_Accrual", _
        "Rolling_Results", _
        "Tenor_Analysis", _
        "Monthly_Returns", _
        "Portfolio_Analysis", _
        "Out_of_Sample", _
        "Chart_Data", _
        "Dashboard", _
        "Test_Results")

    For index = UBound(sheetNames) To LBound(sheetNames) Step -1
        ThisWorkbook.Worksheets(CStr(sheetNames(index))).Move _
            Before:=ThisWorkbook.Worksheets(1)
    Next index
End Sub

Private Sub ApplyTabColors()
    ThisWorkbook.Worksheets("Inputs").Tab.Color = RGB(255, 192, 0)
    ThisWorkbook.Worksheets("Curve").Tab.Color = RGB(255, 192, 0)
    ThisWorkbook.Worksheets("Frontier_Settings").Tab.Color = RGB(255, 192, 0)

    ThisWorkbook.Worksheets("Dashboard").Tab.Color = RGB(0, 176, 80)
    ThisWorkbook.Worksheets("Test_Results").Tab.Color = RGB(112, 173, 71)
End Sub

Private Sub ClearCalculatedSheet(ByVal ws As Worksheet)
    Dim chartObject As ChartObject

    On Error Resume Next
    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
    On Error GoTo 0

    ws.Cells.Clear
End Sub

Public Function GetOrCreateRatesSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateRatesSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateRatesSheet Is Nothing Then
        Set GetOrCreateRatesSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateRatesSheet.Name = sheetName
    End If
End Function

Public Function RatesSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    RatesSheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FindRatesHeaderRow(ByVal ws As Worksheet) As Long
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim foundDate As Boolean
    Dim foundON As Boolean
    Dim valueText As String

    For rowNumber = 1 To 20
        foundDate = False
        foundON = False

        For columnNumber = 1 To 20
            valueText = UCase$(Trim$(CStr( _
                ws.Cells(rowNumber, columnNumber).Value2)))

            If valueText = "DATE" Then foundDate = True
            If valueText = "ON" Then foundON = True
        Next columnNumber

        If foundDate And foundON Then
            FindRatesHeaderRow = rowNumber
            Exit Function
        End If
    Next rowNumber
End Function

Public Sub RatesWriteRow(ByVal firstCell As Range, ByVal values As Variant)
    Dim matrix() As Variant
    Dim itemCount As Long
    Dim index As Long

    itemCount = UBound(values) - LBound(values) + 1
    ReDim matrix(1 To 1, 1 To itemCount)

    For index = 1 To itemCount
        matrix(1, index) = values(LBound(values) + index - 1)
    Next index

    firstCell.Resize(1, itemCount).Value2 = matrix
End Sub

Public Sub RatesStyleTitle(ByVal target As Range)
    With target
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub RatesStyleHeader(ByVal target As Range)
    With target
        .Interior.Color = COLOR_NAVY
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 215, 220)
        .Borders.Weight = xlThin
    End With
End Sub
