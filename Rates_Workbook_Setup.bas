Attribute VB_Name = "RatesWorkbookSetup"
Option Explicit

' ============================================================================
' Rates model workbook setup
'
' Import this module together with Rates_Analysis_Engine.bas.
'
' Public macros:
'   CreateRatesWorkbook  - creates a clean workbook structure
'   EnsureRatesWorkbook  - creates missing sheets without clearing Inputs/Curve
'
' Curve headers:
'   Date | ON | 1M | 2M | 3M | 6M
'
' Rate convention:
'   Rates are ordinary numbers. 4.31 means 4.31 percent.
'
' No cells are merged by this module.
' ============================================================================

Private Const COLOR_NAVY As Long = 3809035
Private Const COLOR_INPUT As Long = 13434879
Private Const COLOR_PALE As Long = 16448250

Public Sub CreateRatesWorkbook()
    BuildRatesWorkbook True, True
End Sub

Public Sub EnsureRatesWorkbook()
    BuildRatesWorkbook False, False
End Sub

Private Sub BuildRatesWorkbook(ByVal resetWorkbook As Boolean, _
                               ByVal showMessage As Boolean)
    Dim oldCalculation As XlCalculation
    Dim stageName As String

    On Error GoTo Fail

    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stageName = "removing legacy model sheets"
    DeleteLegacySheets

    stageName = "creating Inputs"
    SetupInputs resetWorkbook

    stageName = "creating Curve"
    SetupCurve resetWorkbook

    stageName = "creating output sheets"
    SetupOutputSheet "Transactions", "Transaction Schedule", 14, resetWorkbook
    SetupOutputSheet "Daily_Accrual", "Daily Accrual Ledger", 20, resetWorkbook
    SetupOutputSheet "Rolling_Results", "Rolling Investment Results", 15, resetWorkbook
    SetupOutputSheet "Tenor_Analysis", "Daily Tenor Scenario Analysis", 28, resetWorkbook
    SetupOutputSheet "Monthly_Returns", "Monthly Economic Returns", 11, resetWorkbook
    SetupOutputSheet "Portfolio_Analysis", "Historical Maturity Diversification", 15, resetWorkbook
    SetupOutputSheet "Chart_Data", "Chart Data", 40, resetWorkbook
    SetupOutputSheet "Dashboard", "Historical Cash Investment Analysis", 17, resetWorkbook
    SetupOutputSheet "Test_Results", "Model Validation Results", 4, resetWorkbook

    stageName = "applying sheet order"
    OrderModelSheets

CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If showMessage Then
        MsgBox "Workbook structure created." & vbCrLf & _
               "Replace the Curve data, set Inputs!B5:B8 and run RunRatesAnalysis.", _
               vbInformation
    End If
    Exit Sub

Fail:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Workbook setup stopped during " & stageName & ": " & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub SetupInputs(ByVal resetSheet As Boolean)
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet("Inputs")
    If resetSheet Then ClearSheet ws

    If Len(Trim$(CStr(ws.Range("A1").Value2))) = 0 Then
        ws.Range("A1:H1").Interior.Color = COLOR_NAVY
        ws.Range("A1").Value = "Rates Analysis Model - Inputs"
        StyleTitle ws.Range("A1:H1")

        WriteRow ws.Range("A4"), Array("Model Input", "Value")
        StyleHeader ws.Range("A4:B4")

        ws.Range("A5").Value = "Analysis Start Date"
        ws.Range("A6").Value = "Analysis End Date"
        ws.Range("A7").Value = "Initial Notional ($)"
        ws.Range("A8").Value = "Frontier Weight Step (number)"
        ws.Range("A9").Value = "Day Count"
        ws.Range("A10").Value = "Rate Input Convention"
        ws.Range("A11").Value = "Term Maturity Convention"
        ws.Range("A12").Value = "Rolling Convention"

        ws.Range("B5").Value = DateSerial(2023, 1, 3)
        ws.Range("B6").Value = DateSerial(2026, 7, 3)
        ws.Range("B7").Value = 100000000#
        ws.Range("B8").Value = 10#
        ws.Range("B9").Value = "ACT/360"
        ws.Range("B10").Value = "4.31 means 4.31 percent"
        ws.Range("B11").Value = "Latest curve date on or before target"
        ws.Range("B12").Value = "Each new term starts from the prior actual maturity"

        WriteRow ws.Range("D4"), Array("Instructions")
        StyleHeader ws.Range("D4:H4")
        ws.Range("D5").Value = "1. Replace Curve!A:F and retain the headers."
        ws.Range("D6").Value = "2. Only dates from Inputs!B5 through Inputs!B6 are processed."
        ws.Range("D7").Value = "3. Enter rates as numbers such as 4.31, not Excel percentages."
        ws.Range("D8").Value = "4. Run RunRatesAnalysis."
        ws.Range("D9").Value = "5. The default source period is January 2023 through July 2026."
        ws.Range("D10").Value = "6. Weight step 10 means ten-percentage-point increments."
        ws.Range("D11").Value = "7. Detail uses all starts; comparison statistics use common start dates."
        ws.Range("D12").Value = "8. Frontier risk uses aligned monthly strategy returns."

        ws.Range("B5:B8").Interior.Color = COLOR_INPUT
        ws.Range("B5:B8").Font.Color = RGB(0, 0, 255)
        ws.Range("B5:B6").NumberFormat = "mm/dd/yyyy"
        ws.Range("B7").NumberFormat = "$#,##0;[Red]($#,##0);-"
        ws.Range("B8").NumberFormat = "0.0"

        ws.Range("D5:H12").Interior.Color = COLOR_PALE
        ws.Range("D5:H12").WrapText = True

        ws.Columns("A").ColumnWidth = 36
        ws.Columns("B").ColumnWidth = 42
        ws.Columns("C").ColumnWidth = 3
        ws.Columns("D:H").ColumnWidth = 18
        ws.Rows("5:12").RowHeight = 24
        ws.Rows(1).RowHeight = 28
    End If
End Sub

Private Sub SetupCurve(ByVal resetSheet As Boolean)
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet("Curve")
    If resetSheet Then ClearSheet ws

    If Len(Trim$(CStr(ws.Range("A1").Value2))) = 0 Then
        WriteRow ws.Range("A1"), Array("Date", "ON", "1M", "2M", "3M", "6M")
        StyleHeader ws.Range("A1:F1")
        ws.Range("A2:F10000").Font.Color = RGB(0, 0, 255)
        ws.Columns("A").NumberFormat = "mm/dd/yyyy"
        ws.Columns("B:F").NumberFormat = "0.0000"
        ws.Columns("A").ColumnWidth = 14
        ws.Columns("B:F").ColumnWidth = 12
        ws.Rows(1).RowHeight = 24
    End If
End Sub

Private Sub SetupOutputSheet(ByVal sheetName As String, _
                             ByVal titleText As String, _
                             ByVal columnCount As Long, _
                             ByVal resetSheet As Boolean)
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(sheetName)
    If resetSheet Then ClearSheet ws

    If Len(Trim$(CStr(ws.Range("A1").Value2))) = 0 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount)).Interior.Color = COLOR_NAVY
        ws.Range("A1").Value = titleText
        StyleTitle ws.Range(ws.Cells(1, 1), ws.Cells(1, columnCount))
        ws.Rows(1).RowHeight = 28
        ws.Columns(1).ColumnWidth = 15
    End If
End Sub

Private Sub DeleteLegacySheets()
    Dim names As Variant
    Dim item As Variant

    names = Array("Data_Quality", "Premium_Analysis", "Daily_Rolling_Reset", _
                  "Methodology", "Swap_Data", "Swap_Analysis")

    Application.DisplayAlerts = False
    For Each item In names
        If SheetExists(CStr(item)) Then ThisWorkbook.Worksheets(CStr(item)).Delete
    Next item
    Application.DisplayAlerts = True
End Sub

Private Sub OrderModelSheets()
    Dim names As Variant
    Dim i As Long

    names = Array("Inputs", "Curve", "Transactions", "Daily_Accrual", _
                  "Rolling_Results", "Tenor_Analysis", "Monthly_Returns", _
                  "Portfolio_Analysis", "Chart_Data", "Dashboard", "Test_Results")

    For i = UBound(names) To LBound(names) Step -1
        ThisWorkbook.Worksheets(CStr(names(i))).Move _
            Before:=ThisWorkbook.Worksheets(1)
    Next i
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

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub ClearSheet(ByVal ws As Worksheet)
    Dim chartObject As ChartObject

    On Error Resume Next
    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
    On Error GoTo 0

    ws.Cells.Clear
End Sub

Public Sub RatesWriteRow(ByVal firstCell As Range, ByVal values As Variant)
    WriteRow firstCell, values
End Sub

Private Sub WriteRow(ByVal firstCell As Range, ByVal values As Variant)
    Dim matrix() As Variant
    Dim itemCount As Long
    Dim i As Long

    itemCount = UBound(values) - LBound(values) + 1
    ReDim matrix(1 To 1, 1 To itemCount)

    For i = 1 To itemCount
        matrix(1, i) = values(LBound(values) + i - 1)
    Next i

    firstCell.Resize(1, itemCount).Value2 = matrix
End Sub

Public Sub RatesStyleHeader(ByVal target As Range)
    StyleHeader target
End Sub

Private Sub StyleTitle(ByVal target As Range)
    With target
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StyleHeader(ByVal target As Range)
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
