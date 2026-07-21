Attribute VB_Name = "RatesWorkbookSetup"
Option Explicit

' ============================================================================
' Historical Cash Investment Analysis - Workbook Setup
'
' Import this module together with Rates_Analysis_Engine.bas.
'
' Public macros:
'   CreateRatesWorkbook   Rebuilds the workbook structure and clears old data.
'   EnsureRatesWorkbook   Creates missing sheets without clearing Inputs/Curve.
'
' Curve input convention:
'   Date | ON | 1M | 2M | 3M | 6M
'   Rates are ordinary numbers: 4.31 means 4.31 percent.
'
' No cells are merged anywhere in this workbook.
' ============================================================================

Private Const COLOR_NAVY As Long = 6049350       ' RGB(70, 77, 94)
Private Const COLOR_HEADER As Long = 3809035     ' RGB(11, 31, 58)
Private Const COLOR_INPUT As Long = 16777164     ' RGB(255, 255, 204)
Private Const COLOR_PALE As Long = 16185078      ' RGB(246, 248, 250)

Public Sub CreateRatesWorkbook()
    BuildRatesWorkbookStructure True, True
End Sub

Public Sub EnsureRatesWorkbook()
    BuildRatesWorkbookStructure False, False
End Sub

Private Sub BuildRatesWorkbookStructure(ByVal resetWorkbook As Boolean, _
                                         ByVal showMessage As Boolean)

    Dim oldCalc As XlCalculation

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    SetupInputs resetWorkbook
    SetupCurve resetWorkbook

    SetupOutputSheet "Data_Quality", "Data Quality and Controls", "C", _
        Array("Check", "Result", "Status"), resetWorkbook

    SetupOutputSheet "Transactions", "Transaction Schedule", "N", _
        Array("Tenor", "Transaction ID", "Target Start Date", "Actual Start Date", _
              "Rate Observation Date", "Rate Used", "Target Roll Date", _
              "Actual Roll Date", "Transaction Days", "Opening Notional ($)", _
              "Period Interest ($)", "Closing Notional ($)", "Status", _
              "Adjustment Flag"), resetWorkbook

    SetupOutputSheet "Daily_Accrual", "Daily Accrual Ledger", "T", _
        Array("Accrual Date", "Tenor", "Transaction ID", "Transaction Start Date", _
              "Target Roll Date", "Actual Roll Date", "Rate Observation Date", _
              "Rate Used", "Opening Notional ($)", "Daily Interest ($)", _
              "Cumulative Period Interest ($)", "Full Period Interest ($)", _
              "Interest Paid Today ($)", "Economic Balance ($)", "Days Accrued", _
              "Transaction Days", "Days to Roll", "Roll Flag", "Status", _
              "Adjustment Flag"), resetWorkbook

    SetupOutputSheet "Premium_Analysis", "Historical Curve and Term Premium", "T", _
        Array("Date", "ON", "1M", "2M", "3M", "6M", "1M Premium (bps)", _
              "2M Premium (bps)", "3M Premium (bps)", "6M Premium (bps)"), _
              resetWorkbook

    SetupOutputSheet "Rolling_Results", "Rolling Investment Results", "O", _
        Array("Date", "ON", "1M", "2M", "3M", "6M"), resetWorkbook

    SetupOutputSheet "Daily_Rolling_Reset", "Daily Rolling Reinvestment Analysis", "AB", _
        Array("Tenor", "Start Date", "Start Rate Date", "Start Rate", _
              "Target Maturity", "Actual Maturity", "Maturity Rate Date", _
              "Maturity Rate", "Reset Change (bps)", "Actual Days", _
              "Next-Cycle Dollar Impact ($)", "Direction"), resetWorkbook

    SetupOutputSheet "Monthly_Returns", "Monthly Economic Returns", "K", _
        Array("Month End", "ON", "1M", "2M", "3M", "6M"), resetWorkbook

    SetupOutputSheet "Portfolio_Analysis", "Historical Maturity Diversification", "AC", _
        Array("ON Weight", "1M Weight", "2M Weight", "3M Weight", "6M Weight", _
              "Annualized Return", "Annualized Volatility (bps)", _
              "Return / Volatility", "WAM (Months)", "Available <=30D", _
              "Available <=60D", "Available <=90D", "Available <=180D"), _
              resetWorkbook

    SetupOutputSheet "Chart_Data", "Chart Data", "AO", Empty, resetWorkbook
    SetupOutputSheet "Dashboard", "Historical Cash Investment Analysis", "Q", Empty, resetWorkbook
    SetupOutputSheet "Methodology", "Methodology and Definitions", "H", Empty, resetWorkbook

    SetupOutputSheet "Test_Results", "Model Validation Results", "D", _
        Array("Test", "Actual", "Expected / Tolerance", "Status"), resetWorkbook

    OrderSheets
    ApplyWorkbookView

CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        MsgBox "Workbook structure created." & vbCrLf & _
               "Paste the curve into the Curve sheet and run RunRatesAnalysis.", _
               vbInformation
    End If
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Workbook setup stopped: " & Err.Description, vbCritical

End Sub

Private Sub SetupInputs(ByVal resetSheet As Boolean)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Inputs")

    If resetSheet Then ClearSheetCompletely ws

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        ws.Range("A1:H1").Interior.Color = COLOR_HEADER
        ws.Range("A1").Value = "Rates Analysis Model - Inputs"
        ApplyTitle ws.Range("A1:H1")

        WriteRow ws.Range("A4"), Array("Model Input", "Value")
        ApplyHeader ws.Range("A4:B4")

        ws.Range("A5").Value = "Analysis Start Date"
        ws.Range("A6").Value = "Analysis End Date"
        ws.Range("A7").Value = "Initial Notional ($)"
        ws.Range("A8").Value = "Frontier Weight Step (number)"
        ws.Range("A9").Value = "Day Count"
        ws.Range("A10").Value = "Rate Input Convention"
        ws.Range("A11").Value = "Maturity Convention"

        ws.Range("B5").Value = DateSerial(2023, 1, 3)
        ws.Range("B6").Value = DateSerial(2026, 7, 3)
        ws.Range("B7").Value = 100000000#
        ws.Range("B8").Value = 10#
        ws.Range("B9").Value = "ACT/360"
        ws.Range("B10").Value = "4.31 means 4.31 percent"
        ws.Range("B11").Value = "Latest curve date on or before target"

        ws.Range("D4:H4").Interior.Color = COLOR_HEADER
        ws.Range("D4").Value = "Instructions"
        ApplyHeader ws.Range("D4:H4")

        ws.Range("D5").Value = "1. Paste the curve into Curve!A:F and keep the existing headers."
        ws.Range("D6").Value = "2. Enter start and end dates in B5:B6. Only that period is processed."
        ws.Range("D7").Value = "3. Enter rates as numbers, for example 4.31, not 4.31%."
        ws.Range("D8").Value = "4. Run RunRatesAnalysis from the second BAS module."
        ws.Range("D9").Value = "5. Default curve period is January 2023 through July 2026."
        ws.Range("D10").Value = "6. Weight step 10 means ten-percentage-point portfolio increments."

        ws.Range("B5:B8").Interior.Color = COLOR_INPUT
        ws.Range("B5:B8").Font.Color = RGB(0, 0, 255)
        ws.Range("B5:B6").NumberFormat = "mm/dd/yyyy"
        ws.Range("B7").NumberFormat = "$#,##0;[Red]($#,##0);-"
        ws.Range("B8").NumberFormat = "0.0"

        ws.Range("D5:H10").Interior.Color = COLOR_PALE
        ws.Range("D5:H10").WrapText = True

        ws.Columns("A").ColumnWidth = 31
        ws.Columns("B").ColumnWidth = 28
        ws.Columns("C").ColumnWidth = 3
        ws.Columns("D:H").ColumnWidth = 17
        ws.Rows("5:11").RowHeight = 24
        ws.Rows(1).RowHeight = 28
    End If

End Sub

Private Sub SetupCurve(ByVal resetSheet As Boolean)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Curve")

    If resetSheet Then ClearSheetCompletely ws

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        WriteRow ws.Range("A1"), Array("Date", "ON", "1M", "2M", "3M", "6M")
        ApplyHeader ws.Range("A1:F1")

        ws.Range("A2:F10000").Font.Color = RGB(0, 0, 255)
        ws.Range("A:A").NumberFormat = "mm/dd/yyyy"
        ws.Range("B:F").NumberFormat = "0.0000"

        ws.Columns("A").ColumnWidth = 13
        ws.Columns("B:F").ColumnWidth = 12
        ws.Rows(1).RowHeight = 24
        FreezeAt ws, "A2"
    End If

End Sub

Private Sub SetupOutputSheet(ByVal sheetName As String, ByVal titleText As String, _
                             ByVal finalColumn As String, ByVal headers As Variant, _
                             ByVal resetSheet As Boolean)

    Dim ws As Worksheet
    Dim headerCount As Long

    Set ws = GetOrCreateSheet(sheetName)
    If resetSheet Then ClearSheetCompletely ws

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        ws.Range("A1:" & finalColumn & "1").Interior.Color = COLOR_HEADER
        ws.Range("A1").Value = titleText
        ApplyTitle ws.Range("A1:" & finalColumn & "1")
        ws.Rows(1).RowHeight = 28

        If Not IsEmpty(headers) Then
            headerCount = UBound(headers) - LBound(headers) + 1
            WriteRow ws.Range("A3"), headers
            ApplyHeader ws.Range("A3").Resize(1, headerCount)
            ws.Rows(3).RowHeight = 30
            FreezeAt ws, "A4"
        End If

        ws.Columns("A:" & finalColumn).ColumnWidth = 12
        ws.Columns("A").ColumnWidth = 15
    End If

End Sub

Private Sub ApplyTitle(ByVal target As Range)
    With target
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub ApplyRatesHeader(ByVal target As Range)
    ApplyHeader target
End Sub

Private Sub ApplyHeader(ByVal target As Range)
    With target
        .Interior.Color = COLOR_HEADER
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

Public Sub WriteRatesRow(ByVal firstCell As Range, ByVal values As Variant)
    WriteRow firstCell, values
End Sub

Private Sub WriteRow(ByVal firstCell As Range, ByVal values As Variant)

    Dim outputData() As Variant
    Dim itemCount As Long
    Dim i As Long

    itemCount = UBound(values) - LBound(values) + 1
    ReDim outputData(1 To 1, 1 To itemCount)

    For i = 1 To itemCount
        outputData(1, i) = values(LBound(values) + i - 1)
    Next i

    firstCell.Resize(1, itemCount).Value = outputData

End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If

End Function

Private Sub ClearSheetCompletely(ByVal ws As Worksheet)

    Dim chartObject As ChartObject

    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject

    ws.Cells.Clear
    ws.Cells.ClearFormats

End Sub

Private Sub FreezeAt(ByVal ws As Worksheet, ByVal cellAddress As String)

    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(cellAddress).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

End Sub

Private Sub OrderSheets()

    Dim sheetNames As Variant
    Dim i As Long

    sheetNames = Array("Inputs", "Curve", "Data_Quality", "Transactions", _
                       "Daily_Accrual", "Premium_Analysis", "Rolling_Results", _
                       "Daily_Rolling_Reset", "Monthly_Returns", _
                       "Portfolio_Analysis", "Chart_Data", "Dashboard", _
                       "Methodology", "Test_Results")

    For i = UBound(sheetNames) To LBound(sheetNames) Step -1
        ThisWorkbook.Worksheets(CStr(sheetNames(i))).Move _
            Before:=ThisWorkbook.Worksheets(1)
    Next i

End Sub

Private Sub ApplyWorkbookView()

    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws
    ThisWorkbook.Worksheets("Inputs").Activate
    On Error GoTo 0

End Sub
