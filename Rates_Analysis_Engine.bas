Attribute VB_Name = "RatesAnalysisEngine"
Option Explicit

' ============================================================================
' Historical Cash Investment Analysis - Calculation Engine
'
' Import together with Rates_Workbook_Setup.bas.
'
' Main macro:
'   RunRatesAnalysis
'
' Validation macro:
'   ValidateRatesAnalysis
'
' Curve input:
'   Date | ON | 1M | 2M | 3M | 6M
'   Rates are ordinary numbers. Example: 4.31 means 4.31 percent.
'
' Core conventions:
'   - The model processes only Inputs!B5 through Inputs!B6.
'   - Effective start is the first available curve date on or after B5.
'   - Analysis end is exactly B6 and must not exceed the final curve date.
'   - ACT/360 simple interest with interest reinvested at maturity.
'   - ON matures on the next available curve date.
'   - Term maturities use the latest curve date on or before target maturity.
'   - Daily reset analysis uses every eligible curve date as a valid start.
'   - No cells are merged by this module.
' ============================================================================

Private Const TENOR_COUNT As Long = 5
Private Const DAY_COUNT As Double = 360#
Private Const COLOR_NAVY As Long = 3809035
Private Const COLOR_BLUE As Long = 12026821
Private Const COLOR_TEAL As Long = 9413930
Private Const COLOR_GOLD As Long = 2857691
Private Const COLOR_PURPLE As Long = 12736661
Private Const COLOR_GREEN As Long = 6134316
Private Const COLOR_RED As Long = 5394116
Private Const COLOR_GRAY As Long = 10790052
Private Const COLOR_PALE As Long = 16185078

Private gCurveDates() As Double
Private gRates() As Double
Private gCurveCount As Long

Private gRequestedStart As Double
Private gStartDate As Double
Private gEndDate As Double
Private gNotional As Double
Private gWeightStepNumber As Double
Private gWeightUnits As Long
Private gNumDays As Long

Private gBalance() As Double
Private gDailyInterest() As Double
Private gOpeningPrincipal() As Double
Private gCompletedTransactions() As Long
Private gEndingValue() As Double
Private gTotalInterest() As Double
Private gAnnualizedReturnPct() As Double
Private gAverageRateNumber() As Double
Private gPrincipalDaySum() As Double
Private gInterestDaySum() As Double

Private gDailyRows() As Variant
Private gDailyRowCount As Long
Private gTransactionRows As Collection
Private gResetRows As Collection
Private gResetSummary(1 To TENOR_COUNT, 1 To 14) As Variant

Private gMonthEndDates() As Double
Private gMonthlyReturns() As Double
Private gMonthlyReturnCount As Long
Private gTenorAnnualReturnPct() As Double
Private gTenorAnnualVolBps() As Double

Private gPortfolio() As Double
Private gPortfolioCount As Long
Private gFrontier() As Double
Private gFrontierCount As Long

Public Sub RunRatesAnalysis()

    Dim oldCalculation As XlCalculation

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    oldCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Preparing workbook structure..."
    EnsureRatesWorkbook
    ClearOutputData

    Application.StatusBar = "Loading curve and inputs..."
    LoadCurveData
    LoadInputs
    ValidateInputsAndCurve
    InitializeModelArrays

    Application.StatusBar = "Calculating rolling strategies..."
    BuildAllStrategies
    WriteDataQuality
    WriteTransactions
    WriteDailyAccrual

    Application.StatusBar = "Calculating curve premium and reset risk..."
    BuildPremiumAnalysis
    BuildDailyRollingReset

    Application.StatusBar = "Calculating returns and portfolio frontier..."
    BuildRollingResults
    BuildMonthlyReturns
    BuildPortfolioAnalysis

    Application.StatusBar = "Building dashboard and charts..."
    BuildChartData
    BuildDashboard
    BuildMethodology
    WriteValidationResults

    If Not ValidationPassed Then
        Err.Raise vbObjectError + 900, , _
            "One or more model validation tests failed. Review Test_Results."
    End If

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Rates analysis completed for " & Format$(gStartDate, "dd-mmm-yyyy") & _
           " through " & Format$(gEndDate, "dd-mmm-yyyy") & ".", vbInformation
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Rates analysis stopped: " & Err.Description, vbCritical

End Sub

Public Sub ValidateRatesAnalysis()

    On Error GoTo Fail

    If Not SheetExists("Test_Results") Then
        MsgBox "Run RunRatesAnalysis first.", vbExclamation
        Exit Sub
    End If

    If ValidationPassed Then
        MsgBox "All current validation tests passed.", vbInformation
    Else
        MsgBox "At least one validation test failed. Review Test_Results.", vbExclamation
    End If
    Exit Sub

Fail:
    MsgBox "Validation stopped: " & Err.Description, vbCritical

End Sub

' ============================================================================
' Input and curve loading
' ============================================================================

Private Sub LoadInputs()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Inputs")

    If Not IsDate(ws.Range("B5").Value) Then
        Err.Raise vbObjectError + 101, , "Inputs!B5 must contain the analysis start date."
    End If

    If Not IsDate(ws.Range("B6").Value) Then
        Err.Raise vbObjectError + 102, , "Inputs!B6 must contain the analysis end date."
    End If

    If Not IsNumeric(ws.Range("B7").Value) Or CDbl(ws.Range("B7").Value) <= 0 Then
        Err.Raise vbObjectError + 103, , "Inputs!B7 must contain a positive notional."
    End If

    If Not IsNumeric(ws.Range("B8").Value) Then
        Err.Raise vbObjectError + 104, , "Inputs!B8 must contain a numeric frontier step."
    End If

    gRequestedStart = CDbl(CDate(ws.Range("B5").Value))
    gEndDate = CDbl(CDate(ws.Range("B6").Value))
    gNotional = CDbl(ws.Range("B7").Value)
    gWeightStepNumber = CDbl(ws.Range("B8").Value)

End Sub

Private Sub LoadCurveData()

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dateColumn As Long
    Dim tenorColumns(1 To TENOR_COUNT) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim t As Long
    Dim dateValue As Variant
    Dim rateValue As Variant
    Dim currentDate As Double

    Set ws = ThisWorkbook.Worksheets("Curve")

    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then
        Err.Raise vbObjectError + 110, , _
            "Curve headers were not found. Required: Date, ON, 1M, 2M, 3M, 6M."
    End If

    dateColumn = FindHeaderColumn(ws, headerRow, "Date")
    tenorColumns(1) = FindHeaderColumn(ws, headerRow, "ON")
    tenorColumns(2) = FindHeaderColumn(ws, headerRow, "1M")
    tenorColumns(3) = FindHeaderColumn(ws, headerRow, "2M")
    tenorColumns(4) = FindHeaderColumn(ws, headerRow, "3M")
    tenorColumns(5) = FindHeaderColumn(ws, headerRow, "6M")

    If dateColumn = 0 Then
        Err.Raise vbObjectError + 111, , "Curve header Date is missing."
    End If

    For t = 1 To TENOR_COUNT
        If tenorColumns(t) = 0 Then
            Err.Raise vbObjectError + 112, , _
                "Curve header " & TenorName(t) & " is missing."
        End If
    Next t

    lastRow = ws.Cells(ws.Rows.Count, dateColumn).End(xlUp).Row
    If lastRow <= headerRow Then
        Err.Raise vbObjectError + 113, , "The Curve sheet contains no data."
    End If

    ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, 6)).Sort _
        Key1:=ws.Cells(headerRow + 1, dateColumn), Order1:=xlAscending, Header:=xlYes

    gCurveCount = 0
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, dateColumn).Value))) > 0 Then
            gCurveCount = gCurveCount + 1
        End If
    Next r

    If gCurveCount = 0 Then
        Err.Raise vbObjectError + 114, , "The Curve sheet contains no valid rows."
    End If

    ReDim gCurveDates(1 To gCurveCount)
    ReDim gRates(1 To gCurveCount, 1 To TENOR_COUNT)

    i = 0
    For r = headerRow + 1 To lastRow

        dateValue = ws.Cells(r, dateColumn).Value
        If Len(Trim$(CStr(dateValue))) = 0 Then GoTo NextCurveRow

        If IsDate(dateValue) Then
            currentDate = CDbl(CDate(dateValue))
        ElseIf IsNumeric(dateValue) And CDbl(dateValue) > 0 Then
            currentDate = CDbl(dateValue)
        Else
            Err.Raise vbObjectError + 115, , _
                "Invalid curve date on row " & r & "."
        End If

        i = i + 1
        gCurveDates(i) = currentDate

        For t = 1 To TENOR_COUNT
            rateValue = ws.Cells(r, tenorColumns(t)).Value
            If Not IsNumeric(rateValue) Then
                Err.Raise vbObjectError + 116, , _
                    "Non-numeric " & TenorName(t) & " rate on curve row " & r & "."
            End If
            gRates(i, t) = CDbl(rateValue)
        Next t

NextCurveRow:
    Next r

End Sub

Private Sub ValidateInputsAndCurve()

    Dim i As Long
    Dim unitsExact As Double

    If gRequestedStart > gEndDate Then
        Err.Raise vbObjectError + 120, , "Analysis start date is after the end date."
    End If

    For i = 2 To gCurveCount
        If gCurveDates(i) <= gCurveDates(i - 1) Then
            Err.Raise vbObjectError + 121, , _
                "Curve dates must be unique and strictly increasing."
        End If
    Next i

    If gRequestedStart > gCurveDates(gCurveCount) Then
        Err.Raise vbObjectError + 122, , _
            "Analysis start date is after the final curve date."
    End If

    If gEndDate > gCurveDates(gCurveCount) Then
        Err.Raise vbObjectError + 123, , _
            "Analysis end date cannot exceed the final curve date."
    End If

    gStartDate = FirstCurveDateOnOrAfter(gRequestedStart)
    If gStartDate = 0 Or gStartDate > gEndDate Then
        Err.Raise vbObjectError + 124, , _
            "No curve date exists within the requested analysis period."
    End If

    If gWeightStepNumber < 5# Or gWeightStepNumber > 50# Then
        Err.Raise vbObjectError + 125, , _
            "Frontier step must be between 5 and 50. Enter 10 for ten-percent increments."
    End If

    unitsExact = 100# / gWeightStepNumber
    gWeightUnits = CLng(Round(unitsExact, 0))

    If Abs(gWeightUnits - unitsExact) > 0.0000001 Then
        Err.Raise vbObjectError + 126, , _
            "Frontier step must divide 100 exactly. Examples: 5, 10, 20 or 25."
    End If

    If gWeightUnits > 20 Then
        Err.Raise vbObjectError + 127, , _
            "Frontier step creates too many portfolios. Use 5 or a larger number."
    End If

    gNumDays = CLng(gEndDate - gStartDate) + 1
    If gNumDays < 60 Then
        Err.Raise vbObjectError + 128, , _
            "The selected period must contain at least 60 calendar days."
    End If

End Sub

Private Sub InitializeModelArrays()

    ReDim gBalance(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gDailyInterest(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gOpeningPrincipal(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gCompletedTransactions(1 To TENOR_COUNT)
    ReDim gEndingValue(1 To TENOR_COUNT)
    ReDim gTotalInterest(1 To TENOR_COUNT)
    ReDim gAnnualizedReturnPct(1 To TENOR_COUNT)
    ReDim gAverageRateNumber(1 To TENOR_COUNT)
    ReDim gPrincipalDaySum(1 To TENOR_COUNT)
    ReDim gInterestDaySum(1 To TENOR_COUNT)

    ReDim gDailyRows(1 To TENOR_COUNT * gNumDays, 1 To 20)
    gDailyRowCount = 0

    Set gTransactionRows = New Collection
    Set gResetRows = New Collection

End Sub

' ============================================================================
' Rolling strategy calculation
' ============================================================================

Private Sub BuildAllStrategies()

    Dim t As Long

    For t = 1 To TENOR_COUNT
        BuildOneStrategy t
    Next t

End Sub

Private Sub BuildOneStrategy(ByVal tenorIndex As Long)

    Dim currentStart As Double
    Dim targetRoll As Double
    Dim actualRoll As Double
    Dim rateDate As Double
    Dim rateValue As Double
    Dim openingPrincipal As Double
    Dim closingPrincipal As Double
    Dim fullPeriodInterest As Double
    Dim priorInterestPaid As Double
    Dim dailyInterestValue As Double
    Dim transactionDays As Long
    Dim transactionID As Long
    Dim completed As Boolean
    Dim accrualEnd As Double
    Dim currentDay As Double
    Dim dayIndex As Long
    Dim daysAccrued As Long
    Dim adjustmentFlag As String
    Dim statusText As String
    Dim rollFlag As String

    currentStart = gStartDate
    openingPrincipal = gNotional
    transactionID = 1
    priorInterestPaid = 0#

    Do While currentStart <= gEndDate

        rateDate = CurveDateOnOrBefore(currentStart)
        rateValue = RateOnOrBefore(currentStart, tenorIndex)

        If tenorIndex = 1 Then
            targetRoll = currentStart + 1
            actualRoll = NextCurveDateAfter(currentStart)
        Else
            targetRoll = AddMonthsAnchored(gStartDate, transactionID * TenorMonths(tenorIndex))
            If targetRoll <= gCurveDates(gCurveCount) Then
                actualRoll = CurveDateOnOrBefore(targetRoll)
            Else
                actualRoll = 0
            End If
        End If

        If actualRoll > 0 And actualRoll <= currentStart Then
            Err.Raise vbObjectError + 200, , _
                "Invalid roll date for " & TenorName(tenorIndex) & "."
        End If

        completed = (actualRoll > 0 And actualRoll <= gEndDate)

        If actualRoll > 0 Then
            transactionDays = CLng(actualRoll - currentStart)
        Else
            transactionDays = CLng(targetRoll - currentStart)
        End If

        dailyInterestValue = openingPrincipal * (rateValue / 100#) / DAY_COUNT
        fullPeriodInterest = dailyInterestValue * transactionDays
        closingPrincipal = openingPrincipal + fullPeriodInterest

        adjustmentFlag = vbNullString
        If rateDate <> currentStart Then
            adjustmentFlag = "Rate from prior curve date"
        End If
        If actualRoll > 0 And actualRoll <> targetRoll Then
            If Len(adjustmentFlag) > 0 Then adjustmentFlag = adjustmentFlag & "; "
            adjustmentFlag = adjustmentFlag & "Maturity adjusted to prior curve date"
        End If

        If completed Then
            statusText = "COMPLETED"
        Else
            statusText = "OPEN AT ANALYSIS END"
        End If

        AddTransactionRow tenorIndex, transactionID, currentStart, rateDate, rateValue, _
                          targetRoll, actualRoll, transactionDays, openingPrincipal, _
                          fullPeriodInterest, closingPrincipal, statusText, adjustmentFlag

        If actualRoll > 0 Then
            accrualEnd = Application.WorksheetFunction.Min(gEndDate, actualRoll - 1)
        Else
            accrualEnd = gEndDate
        End If

        currentDay = currentStart
        Do While currentDay <= accrualEnd

            dayIndex = CLng(currentDay - gStartDate)
            daysAccrued = CLng(currentDay - currentStart) + 1

            gOpeningPrincipal(tenorIndex, dayIndex) = openingPrincipal
            gDailyInterest(tenorIndex, dayIndex) = dailyInterestValue
            gBalance(tenorIndex, dayIndex) = openingPrincipal + dailyInterestValue * daysAccrued

            gPrincipalDaySum(tenorIndex) = gPrincipalDaySum(tenorIndex) + openingPrincipal
            gInterestDaySum(tenorIndex) = gInterestDaySum(tenorIndex) + dailyInterestValue

            If currentDay = currentStart Then
                If transactionID = 1 Then
                    rollFlag = "START"
                Else
                    rollFlag = "ROLL / NEW DEAL"
                End If
            Else
                rollFlag = vbNullString
            End If

            AddDailyRow tenorIndex, transactionID, currentDay, currentStart, _
                        targetRoll, actualRoll, rateDate, rateValue, openingPrincipal, _
                        dailyInterestValue, dailyInterestValue * daysAccrued, _
                        fullPeriodInterest, IIf(currentDay = currentStart, priorInterestPaid, 0#), _
                        gBalance(tenorIndex, dayIndex), daysAccrued, transactionDays, _
                        DaysToRoll(currentDay, actualRoll, targetRoll), rollFlag, _
                        statusText, adjustmentFlag

            currentDay = currentDay + 1
        Loop

        If Not completed Then Exit Do

        gCompletedTransactions(tenorIndex) = gCompletedTransactions(tenorIndex) + 1
        priorInterestPaid = fullPeriodInterest
        openingPrincipal = closingPrincipal
        currentStart = actualRoll
        transactionID = transactionID + 1

    Loop

    gEndingValue(tenorIndex) = gBalance(tenorIndex, gNumDays - 1)
    gTotalInterest(tenorIndex) = gEndingValue(tenorIndex) - gNotional
    gAnnualizedReturnPct(tenorIndex) = _
        ((gEndingValue(tenorIndex) / gNotional) ^ (DAY_COUNT / gNumDays) - 1#) * 100#

    If gPrincipalDaySum(tenorIndex) > 0 Then
        gAverageRateNumber(tenorIndex) = _
            gInterestDaySum(tenorIndex) * DAY_COUNT / gPrincipalDaySum(tenorIndex) * 100#
    End If

End Sub

Private Sub AddTransactionRow(ByVal tenorIndex As Long, ByVal transactionID As Long, _
                              ByVal actualStart As Double, ByVal rateDate As Double, _
                              ByVal rateValue As Double, ByVal targetRoll As Double, _
                              ByVal actualRoll As Double, ByVal transactionDays As Long, _
                              ByVal openingPrincipal As Double, _
                              ByVal periodInterest As Double, _
                              ByVal closingPrincipal As Double, _
                              ByVal statusText As String, ByVal adjustmentFlag As String)

    Dim rowData As Variant

    rowData = Array(TenorName(tenorIndex), transactionID, actualStart, actualStart, _
                    rateDate, rateValue, targetRoll, BlankIfZero(actualRoll), _
                    transactionDays, openingPrincipal, periodInterest, closingPrincipal, _
                    statusText, adjustmentFlag)

    gTransactionRows.Add rowData

End Sub

Private Sub AddDailyRow(ByVal tenorIndex As Long, ByVal transactionID As Long, _
                        ByVal accrualDate As Double, ByVal transactionStart As Double, _
                        ByVal targetRoll As Double, ByVal actualRoll As Double, _
                        ByVal rateDate As Double, ByVal rateValue As Double, _
                        ByVal openingPrincipal As Double, ByVal dailyInterestValue As Double, _
                        ByVal cumulativeInterest As Double, ByVal fullPeriodInterest As Double, _
                        ByVal interestPaidToday As Double, ByVal economicBalance As Double, _
                        ByVal daysAccrued As Long, ByVal transactionDays As Long, _
                        ByVal daysToRollValue As Long, ByVal rollFlag As String, _
                        ByVal statusText As String, ByVal adjustmentFlag As String)

    gDailyRowCount = gDailyRowCount + 1

    gDailyRows(gDailyRowCount, 1) = accrualDate
    gDailyRows(gDailyRowCount, 2) = TenorName(tenorIndex)
    gDailyRows(gDailyRowCount, 3) = transactionID
    gDailyRows(gDailyRowCount, 4) = transactionStart
    gDailyRows(gDailyRowCount, 5) = targetRoll
    gDailyRows(gDailyRowCount, 6) = BlankIfZero(actualRoll)
    gDailyRows(gDailyRowCount, 7) = rateDate
    gDailyRows(gDailyRowCount, 8) = rateValue
    gDailyRows(gDailyRowCount, 9) = openingPrincipal
    gDailyRows(gDailyRowCount, 10) = dailyInterestValue
    gDailyRows(gDailyRowCount, 11) = cumulativeInterest
    gDailyRows(gDailyRowCount, 12) = fullPeriodInterest
    gDailyRows(gDailyRowCount, 13) = interestPaidToday
    gDailyRows(gDailyRowCount, 14) = economicBalance
    gDailyRows(gDailyRowCount, 15) = daysAccrued
    gDailyRows(gDailyRowCount, 16) = transactionDays
    gDailyRows(gDailyRowCount, 17) = daysToRollValue
    gDailyRows(gDailyRowCount, 18) = rollFlag
    gDailyRows(gDailyRowCount, 19) = statusText
    gDailyRows(gDailyRowCount, 20) = adjustmentFlag

End Sub

Private Function DaysToRoll(ByVal currentDate As Double, ByVal actualRoll As Double, _
                            ByVal targetRoll As Double) As Long
    If actualRoll > 0 Then
        DaysToRoll = Application.WorksheetFunction.Max(0, CLng(actualRoll - currentDate))
    Else
        DaysToRoll = Application.WorksheetFunction.Max(0, CLng(targetRoll - currentDate))
    End If
End Function

' ============================================================================
' Output: data quality, transactions and daily accrual
' ============================================================================

Private Sub WriteDataQuality()

    Dim ws As Worksheet
    Dim rows As Collection
    Dim matrix As Variant
    Dim t As Long

    Set ws = ThisWorkbook.Worksheets("Data_Quality")
    ClearBody ws, "A4:C100"
    Set rows = New Collection

    rows.Add Array("Curve observations", gCurveCount, "PASS")
    rows.Add Array("Curve first date", gCurveDates(1), "PASS")
    rows.Add Array("Curve final date", gCurveDates(gCurveCount), "PASS")
    rows.Add Array("Requested start", gRequestedStart, "PASS")
    rows.Add Array("Effective start", gStartDate, "PASS")
    rows.Add Array("Analysis end", gEndDate, "PASS")
    rows.Add Array("Calendar days processed", gNumDays, "PASS")
    rows.Add Array("Rate convention", "4.31 means 4.31 percent", "PASS")

    For t = 1 To TENOR_COUNT
        rows.Add Array(TenorName(t) & " completed transactions", _
                       gCompletedTransactions(t), "PASS")
    Next t

    matrix = CollectionToMatrix(rows, 3)
    ws.Range("A4").Resize(rows.Count, 3).Value = matrix

    ws.Range("B5:B9").NumberFormat = "mm/dd/yyyy"
    ws.Columns("A:C").AutoFit
    ws.Columns("A").ColumnWidth = 34

End Sub

Private Sub WriteTransactions()

    Dim ws As Worksheet
    Dim matrix As Variant
    Dim rowCount As Long

    Set ws = ThisWorkbook.Worksheets("Transactions")
    ClearBody ws, "A4:N10000"

    rowCount = gTransactionRows.Count
    If rowCount = 0 Then Exit Sub

    matrix = CollectionToMatrix(gTransactionRows, 14)
    ws.Range("A4").Resize(rowCount, 14).Value = matrix

    ws.Range("C4:E" & rowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("G4:H" & rowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("F4:F" & rowCount + 3).NumberFormat = "0.0000"
    ws.Range("J4:L" & rowCount + 3).NumberFormat = _
        "$#,##0;[Red]($#,##0);-"

    ws.Columns("A:N").AutoFit
    ws.Columns("N").ColumnWidth = 40

End Sub

Private Sub WriteDailyAccrual()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Daily_Accrual")
    ClearBody ws, "A4:T50000"

    If gDailyRowCount = 0 Then Exit Sub

    ws.Range("A4").Resize(gDailyRowCount, 20).Value = gDailyRows

    ws.Range("A4:A" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("D4:G" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("H4:H" & gDailyRowCount + 3).NumberFormat = "0.0000"
    ws.Range("I4:N" & gDailyRowCount + 3).NumberFormat = _
        "$#,##0;[Red]($#,##0);-"

    ws.Columns("A:T").AutoFit
    ws.Columns("T").ColumnWidth = 40

End Sub

' ============================================================================
' Curve premium and daily rolling reset analysis
' ============================================================================

Private Sub BuildPremiumAnalysis()

    Dim ws As Worksheet
    Dim firstIndex As Long
    Dim finalIndex As Long
    Dim rowCount As Long
    Dim outputData() As Variant
    Dim summaryData(1 To 4, 1 To 9) As Variant
    Dim i As Long
    Dim r As Long
    Dim t As Long
    Dim countValue As Long
    Dim totalValue As Double
    Dim values() As Double
    Dim currentIndex As Long

    Set ws = ThisWorkbook.Worksheets("Premium_Analysis")
    ClearBody ws, "A4:T10000"

    firstIndex = FindIndexOnOrAfter(gStartDate)
    finalIndex = FindIndexOnOrBefore(gEndDate)
    rowCount = finalIndex - firstIndex + 1

    ReDim outputData(1 To rowCount, 1 To 10)

    r = 0
    For i = firstIndex To finalIndex
        r = r + 1
        outputData(r, 1) = gCurveDates(i)
        For t = 1 To TENOR_COUNT
            outputData(r, t + 1) = gRates(i, t)
        Next t
        For t = 2 To TENOR_COUNT
            outputData(r, t + 5) = (gRates(i, t) - gRates(i, 1)) * 100#
        Next t
    Next i

    ws.Range("A4").Resize(rowCount, 10).Value = outputData
    ws.Range("A4:A" & rowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("B4:F" & rowCount + 3).NumberFormat = "0.0000"
    ws.Range("G4:J" & rowCount + 3).NumberFormat = "0.0;[Red](0.0);-"

    WriteRatesRow ws.Range("L3"), _
        Array("Tenor", "Average (bps)", "Median (bps)", "Minimum (bps)", _
              "Maximum (bps)", "Positive Days", "Current (bps)", _
              "5th Percentile", "95th Percentile")
    ApplyRatesHeader ws.Range("L3:T3")

    currentIndex = FindIndexOnOrBefore(gEndDate)

    For t = 2 To TENOR_COUNT
        ReDim values(1 To rowCount)
        totalValue = 0#
        countValue = 0

        For i = firstIndex To finalIndex
            countValue = countValue + 1
            values(countValue) = (gRates(i, t) - gRates(i, 1)) * 100#
            totalValue = totalValue + values(countValue)
        Next i

        SortDoubleArray values

        summaryData(t - 1, 1) = TenorName(t)
        summaryData(t - 1, 2) = totalValue / countValue
        summaryData(t - 1, 3) = PercentileSorted(values, 0.5)
        summaryData(t - 1, 4) = values(LBound(values))
        summaryData(t - 1, 5) = values(UBound(values))
        summaryData(t - 1, 6) = PositiveShare(values)
        summaryData(t - 1, 7) = (gRates(currentIndex, t) - gRates(currentIndex, 1)) * 100#
        summaryData(t - 1, 8) = PercentileSorted(values, 0.05)
        summaryData(t - 1, 9) = PercentileSorted(values, 0.95)
    Next t

    ws.Range("L4:T7").Value = summaryData
    ws.Range("M4:P7").NumberFormat = "0.0;[Red](0.0);-"
    ws.Range("Q4:Q7").NumberFormat = "0.0%"
    ws.Range("R4:T7").NumberFormat = "0.0;[Red](0.0);-"

    ws.Columns("A:T").AutoFit

End Sub

Private Sub BuildDailyRollingReset()

    Dim ws As Worksheet
    Dim startIndex As Long
    Dim i As Long
    Dim t As Long
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim maturityIndex As Long
    Dim startRate As Double
    Dim maturityRate As Double
    Dim resetChange As Double
    Dim actualDays As Long
    Dim dollarImpact As Double
    Dim directionText As String
    Dim rowData As Variant
    Dim detailMatrix As Variant
    Dim summaryMatrix(1 To TENOR_COUNT, 1 To 14) As Variant

    Set ws = ThisWorkbook.Worksheets("Daily_Rolling_Reset")
    ClearBody ws, "A4:AB50000"
    Set gResetRows = New Collection

    startIndex = FindIndexOnOrAfter(gStartDate)

    For t = 1 To TENOR_COUNT
        For i = startIndex To gCurveCount

            If gCurveDates(i) > gEndDate Then Exit For

            If t = 1 Then
                actualMaturity = NextCurveDateAfter(gCurveDates(i))
                targetMaturity = gCurveDates(i) + 1
                If actualMaturity = 0 Or actualMaturity > gEndDate Then Exit For
            Else
                targetMaturity = AddMonthsAnchored(gCurveDates(i), TenorMonths(t))
                If targetMaturity > gEndDate Then Exit For
                maturityIndex = FindIndexOnOrBefore(targetMaturity)
                actualMaturity = gCurveDates(maturityIndex)
                If actualMaturity <= gCurveDates(i) Then GoTo NextResetStart
            End If

            startRate = gRates(i, t)
            maturityIndex = FindIndexOnOrBefore(actualMaturity)
            maturityRate = gRates(maturityIndex, t)
            resetChange = (maturityRate - startRate) * 100#
            actualDays = CLng(actualMaturity - gCurveDates(i))
            dollarImpact = gNotional * resetChange / 10000# * actualDays / DAY_COUNT

            If resetChange > 0 Then
                directionText = "Higher"
            ElseIf resetChange < 0 Then
                directionText = "Lower"
            Else
                directionText = "Unchanged"
            End If

            rowData = Array(TenorName(t), gCurveDates(i), gCurveDates(i), startRate, _
                            targetMaturity, actualMaturity, actualMaturity, _
                            maturityRate, resetChange, actualDays, dollarImpact, _
                            directionText)
            gResetRows.Add rowData

NextResetStart:
        Next i
    Next t

    If gResetRows.Count > 0 Then
        detailMatrix = CollectionToMatrix(gResetRows, 12)
        ws.Range("A4").Resize(gResetRows.Count, 12).Value = detailMatrix
        ws.Range("B4:C" & gResetRows.Count + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("E4:G" & gResetRows.Count + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("D4:D" & gResetRows.Count + 3).NumberFormat = "0.0000"
        ws.Range("H4:H" & gResetRows.Count + 3).NumberFormat = "0.0000"
        ws.Range("I4:I" & gResetRows.Count + 3).NumberFormat = "0.0;[Red](0.0);-"
        ws.Range("K4:K" & gResetRows.Count + 3).NumberFormat = _
            "$#,##0;[Red]($#,##0);-"
    End If

    WriteRatesRow ws.Range("O3"), _
        Array("Tenor", "Daily Starting Scenarios", "Average Start Rate", _
              "Average Maturity Rate", "Average Reset (bps)", _
              "Reset Volatility (bps)", "Median Reset (bps)", _
              "5th Percentile (bps)", "95th Percentile (bps)", _
              "Worst Decline (bps)", "Largest Increase (bps)", _
              "Positive Resets", "Average Dollar Impact ($)", _
              "Worst Dollar Impact ($)")
    ApplyRatesHeader ws.Range("O3:AB3")

    CalculateResetSummary

    For t = 1 To TENOR_COUNT
        For i = 1 To 14
            summaryMatrix(t, i) = gResetSummary(t, i)
        Next i
    Next t

    ws.Range("O4:AB8").Value = summaryMatrix
    ws.Range("P4:P8").NumberFormat = "0"
    ws.Range("Q4:R8").NumberFormat = "0.0000"
    ws.Range("S4:Y8").NumberFormat = "0.0;[Red](0.0);-"
    ws.Range("Z4:Z8").NumberFormat = "0.0%"
    ws.Range("AA4:AB8").NumberFormat = "$#,##0;[Red]($#,##0);-"

    ws.Columns("A:AB").AutoFit
    ws.Columns("AA:AB").ColumnWidth = 19

End Sub

Private Sub CalculateResetSummary()

    Dim t As Long
    Dim i As Long
    Dim rowData As Variant
    Dim countValue As Long
    Dim values() As Double
    Dim impacts() As Double
    Dim startRates() As Double
    Dim maturityRates() As Double
    Dim sumValues As Double
    Dim sumImpacts As Double
    Dim sumStart As Double
    Dim sumMaturity As Double

    For t = 1 To TENOR_COUNT

        countValue = 0
        For i = 1 To gResetRows.Count
            rowData = gResetRows(i)
            If CStr(rowData(0)) = TenorName(t) Then countValue = countValue + 1
        Next i

        If countValue = 0 Then
            Err.Raise vbObjectError + 300, , _
                "No complete daily reset scenarios for " & TenorName(t) & "."
        End If

        ReDim values(1 To countValue)
        ReDim impacts(1 To countValue)
        ReDim startRates(1 To countValue)
        ReDim maturityRates(1 To countValue)

        countValue = 0
        For i = 1 To gResetRows.Count
            rowData = gResetRows(i)
            If CStr(rowData(0)) = TenorName(t) Then
                countValue = countValue + 1
                startRates(countValue) = CDbl(rowData(3))
                maturityRates(countValue) = CDbl(rowData(7))
                values(countValue) = CDbl(rowData(8))
                impacts(countValue) = CDbl(rowData(10))
                sumStart = sumStart + startRates(countValue)
                sumMaturity = sumMaturity + maturityRates(countValue)
                sumValues = sumValues + values(countValue)
                sumImpacts = sumImpacts + impacts(countValue)
            End If
        Next i

        SortDoubleArray values
        SortDoubleArray impacts

        gResetSummary(t, 1) = TenorName(t)
        gResetSummary(t, 2) = countValue
        gResetSummary(t, 3) = sumStart / countValue
        gResetSummary(t, 4) = sumMaturity / countValue
        gResetSummary(t, 5) = sumValues / countValue
        gResetSummary(t, 6) = SampleStdDev(values)
        gResetSummary(t, 7) = PercentileSorted(values, 0.5)
        gResetSummary(t, 8) = PercentileSorted(values, 0.05)
        gResetSummary(t, 9) = PercentileSorted(values, 0.95)
        gResetSummary(t, 10) = values(LBound(values))
        gResetSummary(t, 11) = values(UBound(values))
        gResetSummary(t, 12) = PositiveShare(values)
        gResetSummary(t, 13) = sumImpacts / countValue
        gResetSummary(t, 14) = impacts(LBound(impacts))

        sumStart = 0#
        sumMaturity = 0#
        sumValues = 0#
        sumImpacts = 0#

    Next t

End Sub

' ============================================================================
' Rolling results and aligned monthly returns
' ============================================================================

Private Sub BuildRollingResults()

    Dim ws As Worksheet
    Dim growthData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 8) As Variant
    Dim d As Long
    Dim t As Long

    Set ws = ThisWorkbook.Worksheets("Rolling_Results")
    ClearBody ws, "A4:O10000"

    ReDim growthData(1 To gNumDays, 1 To 6)

    For d = 0 To gNumDays - 1
        growthData(d + 1, 1) = gStartDate + d
        For t = 1 To TENOR_COUNT
            growthData(d + 1, t + 1) = gBalance(t, d)
        Next t
    Next d

    ws.Range("A4").Resize(gNumDays, 6).Value = growthData
    ws.Range("A4:A" & gNumDays + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("B4:F" & gNumDays + 3).NumberFormat = _
        "$#,##0;[Red]($#,##0);-"

    WriteRatesRow ws.Range("H3"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", _
              "Total Return", "Annualized Return", "Average Invested Rate", _
              "Completed Transactions", "Incremental Interest vs ON ($)")
    ApplyRatesHeader ws.Range("H3:O3")

    For t = 1 To TENOR_COUNT
        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gEndingValue(t)
        summaryData(t, 3) = gTotalInterest(t)
        summaryData(t, 4) = (gEndingValue(t) / gNotional - 1#) * 100#
        summaryData(t, 5) = gAnnualizedReturnPct(t)
        summaryData(t, 6) = gAverageRateNumber(t)
        summaryData(t, 7) = gCompletedTransactions(t)
        summaryData(t, 8) = gTotalInterest(t) - gTotalInterest(1)
    Next t

    ws.Range("H4:O8").Value = summaryData
    ws.Range("I4:J8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("K4:M8").NumberFormat = "0.0000\%"
    ws.Range("N4:N8").NumberFormat = "0"
    ws.Range("O4:O8").NumberFormat = "$#,##0;[Red]($#,##0);-"

    ws.Columns("A:O").AutoFit

End Sub

Private Sub BuildMonthlyReturns()

    Dim ws As Worksheet
    Dim allMonthEnds() As Double
    Dim monthEndCount As Long
    Dim currentMonthEnd As Double
    Dim outputData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 4) As Variant
    Dim i As Long
    Dim t As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim productValue As Double
    Dim values() As Double

    Set ws = ThisWorkbook.Worksheets("Monthly_Returns")
    ClearBody ws, "A4:K500"

    currentMonthEnd = DateSerial(Year(CDate(gStartDate)), Month(CDate(gStartDate)) + 1, 0)

    Do While currentMonthEnd <= gEndDate
        monthEndCount = monthEndCount + 1
        If monthEndCount = 1 Then
            ReDim allMonthEnds(1 To 1)
        Else
            ReDim Preserve allMonthEnds(1 To monthEndCount)
        End If
        allMonthEnds(monthEndCount) = currentMonthEnd
        currentMonthEnd = DateSerial(Year(CDate(currentMonthEnd)), _
                                     Month(CDate(currentMonthEnd)) + 2, 0)
    Loop

    If monthEndCount < 3 Then
        Err.Raise vbObjectError + 400, , _
            "At least three month-end observations are required."
    End If

    gMonthlyReturnCount = monthEndCount - 1
    ReDim gMonthEndDates(1 To monthEndCount)
    ReDim gMonthlyReturns(1 To gMonthlyReturnCount, 1 To TENOR_COUNT)
    ReDim gTenorAnnualReturnPct(1 To TENOR_COUNT)
    ReDim gTenorAnnualVolBps(1 To TENOR_COUNT)
    ReDim outputData(1 To gMonthlyReturnCount, 1 To 6)

    For i = 1 To monthEndCount
        gMonthEndDates(i) = allMonthEnds(i)
    Next i

    For i = 1 To gMonthlyReturnCount
        outputData(i, 1) = allMonthEnds(i + 1)
        startIndex = CLng(allMonthEnds(i) - gStartDate)
        endIndex = CLng(allMonthEnds(i + 1) - gStartDate)

        For t = 1 To TENOR_COUNT
            gMonthlyReturns(i, t) = gBalance(t, endIndex) / gBalance(t, startIndex) - 1#
            outputData(i, t + 1) = gMonthlyReturns(i, t) * 100#
        Next t
    Next i

    ws.Range("A4").Resize(gMonthlyReturnCount, 6).Value = outputData
    ws.Range("A4:A" & gMonthlyReturnCount + 3).NumberFormat = "mmm-yy"
    ws.Range("B4:F" & gMonthlyReturnCount + 3).NumberFormat = "0.0000\%"

    WriteRatesRow ws.Range("H3"), _
        Array("Tenor", "Annualized Return", "Annualized Volatility (bps)", _
              "WAM (Months)")
    ApplyRatesHeader ws.Range("H3:K3")

    For t = 1 To TENOR_COUNT
        ReDim values(1 To gMonthlyReturnCount)
        productValue = 1#

        For i = 1 To gMonthlyReturnCount
            values(i) = gMonthlyReturns(i, t)
            productValue = productValue * (1# + values(i))
        Next i

        gTenorAnnualReturnPct(t) = _
            (productValue ^ (12# / gMonthlyReturnCount) - 1#) * 100#
        gTenorAnnualVolBps(t) = SampleStdDev(values) * Sqr(12#) * 10000#

        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gTenorAnnualReturnPct(t)
        summaryData(t, 3) = gTenorAnnualVolBps(t)
        summaryData(t, 4) = TenorMonths(t)
    Next t

    ws.Range("H4:K8").Value = summaryData
    ws.Range("I4:I8").NumberFormat = "0.0000\%"
    ws.Range("J4:J8").NumberFormat = "0.00"
    ws.Range("K4:K8").NumberFormat = "0.0"
    ws.Columns("A:K").AutoFit

End Sub

' ============================================================================
' Portfolio combinations and efficient frontier
' ============================================================================

Private Sub BuildPortfolioAnalysis()

    Dim ws As Worksheet
    Dim combinationCount As Long
    Dim w0 As Long
    Dim w1 As Long
    Dim w2 As Long
    Dim w3 As Long
    Dim w4 As Long
    Dim rowIndex As Long
    Dim weight(1 To TENOR_COUNT) As Double
    Dim monthlyPortfolio() As Double
    Dim i As Long
    Dim t As Long
    Dim productValue As Double
    Dim annualReturnPct As Double
    Dim annualVolBps As Double
    Dim ratioValue As Double
    Dim portfolioOutput() As Variant
    Dim selectedOutput(1 To 5, 1 To 10) As Variant
    Dim frontierOutput() As Variant
    Dim p As Long
    Dim rank As Long
    Dim segmentText As String
    Dim descriptionText As String
    Dim allocationText As String
    Dim minVolRow As Long
    Dim maxRatioRow As Long

    Set ws = ThisWorkbook.Worksheets("Portfolio_Analysis")
    ClearBody ws, "A4:AC50000"

    combinationCount = CombinationCount(gWeightUnits + 4, 4)
    ReDim gPortfolio(1 To combinationCount, 1 To 13)
    ReDim monthlyPortfolio(1 To gMonthlyReturnCount)

    rowIndex = 0

    For w0 = 0 To gWeightUnits
        For w1 = 0 To gWeightUnits - w0
            For w2 = 0 To gWeightUnits - w0 - w1
                For w3 = 0 To gWeightUnits - w0 - w1 - w2

                    w4 = gWeightUnits - w0 - w1 - w2 - w3
                    weight(1) = w0 / gWeightUnits
                    weight(2) = w1 / gWeightUnits
                    weight(3) = w2 / gWeightUnits
                    weight(4) = w3 / gWeightUnits
                    weight(5) = w4 / gWeightUnits

                    For i = 1 To gMonthlyReturnCount
                        monthlyPortfolio(i) = 0#
                        For t = 1 To TENOR_COUNT
                            monthlyPortfolio(i) = monthlyPortfolio(i) + _
                                                  weight(t) * gMonthlyReturns(i, t)
                        Next t
                    Next i

                    productValue = 1#
                    For i = 1 To gMonthlyReturnCount
                        productValue = productValue * (1# + monthlyPortfolio(i))
                    Next i

                    annualReturnPct = _
                        (productValue ^ (12# / gMonthlyReturnCount) - 1#) * 100#
                    annualVolBps = SampleStdDev(monthlyPortfolio) * Sqr(12#) * 10000#

                    If annualVolBps > 0 Then
                        ratioValue = (annualReturnPct / 100#) / (annualVolBps / 10000#)
                    Else
                        ratioValue = 0#
                    End If

                    rowIndex = rowIndex + 1
                    For t = 1 To TENOR_COUNT
                        gPortfolio(rowIndex, t) = weight(t)
                    Next t
                    gPortfolio(rowIndex, 6) = annualReturnPct
                    gPortfolio(rowIndex, 7) = annualVolBps
                    gPortfolio(rowIndex, 8) = ratioValue
                    gPortfolio(rowIndex, 9) = weight(2) + 2# * weight(3) + _
                                                3# * weight(4) + 6# * weight(5)
                    gPortfolio(rowIndex, 10) = weight(1) + weight(2)
                    gPortfolio(rowIndex, 11) = gPortfolio(rowIndex, 10) + weight(3)
                    gPortfolio(rowIndex, 12) = gPortfolio(rowIndex, 11) + weight(4)
                    gPortfolio(rowIndex, 13) = 1#

                Next w3
            Next w2
        Next w1
    Next w0

    gPortfolioCount = rowIndex
    QuickSortPortfolio 1, gPortfolioCount
    BuildFrontier

    ReDim portfolioOutput(1 To gPortfolioCount, 1 To 13)
    For p = 1 To gPortfolioCount
        For t = 1 To 5
            portfolioOutput(p, t) = gPortfolio(p, t) * 100#
        Next t
        portfolioOutput(p, 6) = gPortfolio(p, 6)
        portfolioOutput(p, 7) = gPortfolio(p, 7)
        portfolioOutput(p, 8) = gPortfolio(p, 8)
        portfolioOutput(p, 9) = gPortfolio(p, 9)
        For t = 10 To 13
            portfolioOutput(p, t) = gPortfolio(p, t) * 100#
        Next t
    Next p

    ws.Range("A4").Resize(gPortfolioCount, 13).Value = portfolioOutput
    ws.Range("A4:E" & gPortfolioCount + 3).NumberFormat = "0.0\%"
    ws.Range("F4:F" & gPortfolioCount + 3).NumberFormat = "0.0000\%"
    ws.Range("G4:G" & gPortfolioCount + 3).NumberFormat = "0.00"
    ws.Range("H4:H" & gPortfolioCount + 3).NumberFormat = "0.00x"
    ws.Range("I4:I" & gPortfolioCount + 3).NumberFormat = "0.0"
    ws.Range("J4:M" & gPortfolioCount + 3).NumberFormat = "0.0\%"

    WriteRatesRow ws.Range("O3"), _
        Array("Portfolio", "ON Weight", "1M Weight", "2M Weight", _
              "3M Weight", "6M Weight", "Annualized Return", _
              "Annualized Volatility (bps)", "WAM (Months)", "Description")
    ApplyRatesHeader ws.Range("O3:X3")

    minVolRow = 1
    maxRatioRow = 1
    For p = 2 To gPortfolioCount
        If gPortfolio(p, 7) < gPortfolio(minVolRow, 7) Then minVolRow = p
        If gPortfolio(p, 8) > gPortfolio(maxRatioRow, 8) Then maxRatioRow = p
    Next p

    FillSelectedPortfolio selectedOutput, 1, "100% ON", FindPortfolioRow(1#, 0#, 0#, 0#, 0#), _
        "Immediate-liquidity reference portfolio; not necessarily efficient."
    FillSelectedPortfolio selectedOutput, 2, "Minimum Volatility", minVolRow, _
        "Lowest historical earnings volatility among tested allocations."
    FillSelectedPortfolio selectedOutput, 3, "Closest to Equal Weight", FindClosestEqualPortfolioRow(), _
        "Tested allocation closest to equal capital across all tenors."
    FillSelectedPortfolio selectedOutput, 4, "Maximum Return / Volatility", maxRatioRow, _
        "Highest historical return per unit of earnings volatility."
    FillSelectedPortfolio selectedOutput, 5, "100% 6M", FindPortfolioRow(0#, 0#, 0#, 0#, 1#), _
        "Longest-maturity reference portfolio."

    ws.Range("O4:X8").Value = selectedOutput
    ws.Range("P4:T8").NumberFormat = "0.0\%"
    ws.Range("U4:U8").NumberFormat = "0.0000\%"
    ws.Range("V4:V8").NumberFormat = "0.00"
    ws.Range("W4:W8").NumberFormat = "0.0"
    ws.Range("X4:X8").WrapText = True

    WriteRatesRow ws.Range("O11"), _
        Array("Annualized Volatility (bps)", "Annualized Return", _
              "ON Weight", "1M Weight", "2M Weight", "3M Weight", _
              "6M Weight", "WAM (Months)", "Available <=30D", _
              "Available <=60D", "Available <=90D", "Frontier Rank", _
              "Frontier Segment", "Description", "Allocation Summary")
    ApplyRatesHeader ws.Range("O11:AC11")

    ReDim frontierOutput(1 To gFrontierCount, 1 To 15)

    For rank = 1 To gFrontierCount
        segmentText = FrontierSegment(rank, gFrontierCount)
        descriptionText = FrontierDescription(segmentText)
        allocationText = AllocationDescription(rank)

        frontierOutput(rank, 1) = gFrontier(rank, 7)
        frontierOutput(rank, 2) = gFrontier(rank, 6)
        For t = 1 To TENOR_COUNT
            frontierOutput(rank, t + 2) = gFrontier(rank, t) * 100#
        Next t
        frontierOutput(rank, 8) = gFrontier(rank, 9)
        frontierOutput(rank, 9) = gFrontier(rank, 10) * 100#
        frontierOutput(rank, 10) = gFrontier(rank, 11) * 100#
        frontierOutput(rank, 11) = gFrontier(rank, 12) * 100#
        frontierOutput(rank, 12) = rank
        frontierOutput(rank, 13) = segmentText
        frontierOutput(rank, 14) = descriptionText
        frontierOutput(rank, 15) = allocationText
    Next rank

    ws.Range("O12").Resize(gFrontierCount, 15).Value = frontierOutput
    ws.Range("O12:O" & gFrontierCount + 11).NumberFormat = "0.00"
    ws.Range("P12:P" & gFrontierCount + 11).NumberFormat = "0.0000\%"
    ws.Range("Q12:U" & gFrontierCount + 11).NumberFormat = "0.0\%"
    ws.Range("V12:V" & gFrontierCount + 11).NumberFormat = "0.0"
    ws.Range("W12:Y" & gFrontierCount + 11).NumberFormat = "0.0\%"
    ws.Range("AB12:AC" & gFrontierCount + 11).WrapText = True

    ws.Columns("A:AC").AutoFit
    ws.Columns("X").ColumnWidth = 42
    ws.Columns("AB:AC").ColumnWidth = 48

End Sub

Private Sub BuildFrontier()

    Dim temporary() As Double
    Dim bestReturn As Double
    Dim i As Long
    Dim c As Long

    ReDim temporary(1 To gPortfolioCount, 1 To 13)
    bestReturn = -1E+99
    gFrontierCount = 0

    For i = 1 To gPortfolioCount
        If gPortfolio(i, 6) > bestReturn + 0.0000001 Then
            gFrontierCount = gFrontierCount + 1
            For c = 1 To 13
                temporary(gFrontierCount, c) = gPortfolio(i, c)
            Next c
            bestReturn = gPortfolio(i, 6)
        End If
    Next i

    ReDim gFrontier(1 To gFrontierCount, 1 To 13)
    For i = 1 To gFrontierCount
        For c = 1 To 13
            gFrontier(i, c) = temporary(i, c)
        Next c
    Next i

End Sub

Private Sub FillSelectedPortfolio(ByRef outputData() As Variant, ByVal outputRow As Long, _
                                  ByVal portfolioName As String, ByVal portfolioRow As Long, _
                                  ByVal descriptionText As String)

    Dim t As Long

    outputData(outputRow, 1) = portfolioName
    For t = 1 To TENOR_COUNT
        outputData(outputRow, t + 1) = gPortfolio(portfolioRow, t) * 100#
    Next t
    outputData(outputRow, 7) = gPortfolio(portfolioRow, 6)
    outputData(outputRow, 8) = gPortfolio(portfolioRow, 7)
    outputData(outputRow, 9) = gPortfolio(portfolioRow, 9)
    outputData(outputRow, 10) = descriptionText

End Sub


Private Function FindClosestEqualPortfolioRow() As Long

    Dim i As Long
    Dim t As Long
    Dim distanceValue As Double
    Dim bestDistance As Double

    bestDistance = 1E+99
    FindClosestEqualPortfolioRow = 1

    For i = 1 To gPortfolioCount
        distanceValue = 0#
        For t = 1 To TENOR_COUNT
            distanceValue = distanceValue + (gPortfolio(i, t) - 0.2) ^ 2
        Next t

        If distanceValue < bestDistance Then
            bestDistance = distanceValue
            FindClosestEqualPortfolioRow = i
        End If
    Next i

End Function

Private Function FindPortfolioRow(ByVal w1 As Double, ByVal w2 As Double, _
                                  ByVal w3 As Double, ByVal w4 As Double, _
                                  ByVal w5 As Double) As Long

    Dim i As Long

    For i = 1 To gPortfolioCount
        If Abs(gPortfolio(i, 1) - w1) < 0.0000001 And _
           Abs(gPortfolio(i, 2) - w2) < 0.0000001 And _
           Abs(gPortfolio(i, 3) - w3) < 0.0000001 And _
           Abs(gPortfolio(i, 4) - w4) < 0.0000001 And _
           Abs(gPortfolio(i, 5) - w5) < 0.0000001 Then
            FindPortfolioRow = i
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 500, , "Required reference portfolio was not found."

End Function

Private Function FrontierSegment(ByVal rank As Long, ByVal totalRanks As Long) As String

    Dim positionValue As Double

    If rank = 1 Then
        FrontierSegment = "Minimum Volatility"
        Exit Function
    End If

    If totalRanks <= 1 Then
        FrontierSegment = "Minimum Volatility"
        Exit Function
    End If

    positionValue = (rank - 1) / (totalRanks - 1)

    If positionValue <= 0.25 Then
        FrontierSegment = "Defensive"
    ElseIf positionValue <= 0.5 Then
        FrontierSegment = "Conservative"
    ElseIf positionValue <= 0.75 Then
        FrontierSegment = "Balanced"
    ElseIf positionValue < 1# Then
        FrontierSegment = "Return Oriented"
    Else
        FrontierSegment = "Maximum Historical Return"
    End If

End Function

Private Function FrontierDescription(ByVal segmentText As String) As String

    Select Case segmentText
        Case "Minimum Volatility"
            FrontierDescription = _
                "Lowest historical earnings volatility on the efficient frontier."
        Case "Defensive"
            FrontierDescription = _
                "Small volatility increase from the minimum-risk point for incremental historical return."
        Case "Conservative"
            FrontierDescription = _
                "Moderate return improvement while remaining in the lower half of frontier volatility."
        Case "Balanced"
            FrontierDescription = _
                "Middle of the frontier, balancing historical return, volatility and liquidity."
        Case "Return Oriented"
            FrontierDescription = _
                "Higher historical return with greater volatility or reduced short-term liquidity."
        Case Else
            FrontierDescription = _
                "Highest historical return among efficient tested allocations."
    End Select

End Function

Private Function AllocationDescription(ByVal frontierRow As Long) As String

    AllocationDescription = _
        Format$(gFrontier(frontierRow, 1), "0%") & " ON / " & _
        Format$(gFrontier(frontierRow, 2), "0%") & " 1M / " & _
        Format$(gFrontier(frontierRow, 3), "0%") & " 2M / " & _
        Format$(gFrontier(frontierRow, 4), "0%") & " 3M / " & _
        Format$(gFrontier(frontierRow, 5), "0%") & " 6M; WAM " & _
        Format$(gFrontier(frontierRow, 9), "0.0") & " months; " & _
        Format$(gFrontier(frontierRow, 10), "0%") & _
        " available within 30 days."

End Function

' ============================================================================
' Chart data and dashboard
' ============================================================================

Private Sub BuildChartData()

    Dim ws As Worksheet
    Dim firstCurveIndex As Long
    Dim finalCurveIndex As Long
    Dim i As Long
    Dim t As Long
    Dim chartRow As Long
    Dim previousMonth As Long
    Dim previousYear As Long
    Dim lastMonthIndex As Long
    Dim monthEndDate As Double
    Dim finalDateIncluded As Boolean
    Dim premiumSummary(1 To 4, 1 To 3) As Variant
    Dim incrementalData(1 To 4, 1 To 2) As Variant
    Dim riskData(1 To TENOR_COUNT, 1 To 3) As Variant
    Dim resetVolData(1 To 4, 1 To 2) As Variant
    Dim resetDistribution(1 To 4, 1 To 4) As Variant
    Dim resultSummary(1 To TENOR_COUNT, 1 To 6) As Variant
    Dim currentCurveIndex As Long

    Set ws = ThisWorkbook.Worksheets("Chart_Data")
    ws.Cells.Clear

    WriteRatesRow ws.Range("A1"), Array("Date", "ON", "1M", "2M", "3M", "6M")
    ApplyRatesHeader ws.Range("A1:F1")

    firstCurveIndex = FindIndexOnOrAfter(gStartDate)
    finalCurveIndex = FindIndexOnOrBefore(gEndDate)
    chartRow = 2
    previousMonth = -1
    previousYear = -1
    lastMonthIndex = firstCurveIndex

    For i = firstCurveIndex To finalCurveIndex
        If previousMonth = -1 Then
            previousMonth = Month(CDate(gCurveDates(i)))
            previousYear = Year(CDate(gCurveDates(i)))
            lastMonthIndex = i
        ElseIf Month(CDate(gCurveDates(i))) <> previousMonth Or _
               Year(CDate(gCurveDates(i))) <> previousYear Then

            ws.Cells(chartRow, 1).Value = Format$(gCurveDates(lastMonthIndex), "mmm-yy")
            ws.Cells(chartRow, 42).Value = gCurveDates(lastMonthIndex)
            For t = 1 To TENOR_COUNT
                ws.Cells(chartRow, t + 1).Value = gRates(lastMonthIndex, t)
            Next t
            chartRow = chartRow + 1

            previousMonth = Month(CDate(gCurveDates(i)))
            previousYear = Year(CDate(gCurveDates(i)))
            lastMonthIndex = i
        Else
            lastMonthIndex = i
        End If
    Next i

    ws.Cells(chartRow, 1).Value = Format$(gCurveDates(lastMonthIndex), "mmm-yy")
    ws.Cells(chartRow, 42).Value = gCurveDates(lastMonthIndex)
    For t = 1 To TENOR_COUNT
        ws.Cells(chartRow, t + 1).Value = gRates(lastMonthIndex, t)
    Next t

    If gCurveDates(lastMonthIndex) <> gEndDate Then
        currentCurveIndex = FindIndexOnOrBefore(gEndDate)
        If gCurveDates(currentCurveIndex) <> gCurveDates(lastMonthIndex) Then
            chartRow = chartRow + 1
            ws.Cells(chartRow, 1).Value = Format$(gCurveDates(currentCurveIndex), "mmm-yy")
            ws.Cells(chartRow, 42).Value = gCurveDates(currentCurveIndex)
            For t = 1 To TENOR_COUNT
                ws.Cells(chartRow, t + 1).Value = gRates(currentCurveIndex, t)
            Next t
        End If
    End If

    ws.Range("A2:A" & chartRow).NumberFormat = "@"
    ws.Range("AP2:AP" & chartRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("B2:F" & chartRow).NumberFormat = "0.0000"

    WriteRatesRow ws.Range("H1"), Array("Date", "1M", "2M", "3M", "6M")
    ApplyRatesHeader ws.Range("H1:L1")

    chartRow = 2
    For i = 1 To UBound(gMonthEndDates)
        monthEndDate = gMonthEndDates(i)
        ws.Cells(chartRow, 8).Value = Format$(monthEndDate, "mmm-yy")
        ws.Cells(chartRow, 43).Value = monthEndDate
        For t = 2 To TENOR_COUNT
            ws.Cells(chartRow, t + 7).Value = _
                (BalanceOnDate(t, monthEndDate) - BalanceOnDate(1, monthEndDate)) / 1000#
        Next t
        chartRow = chartRow + 1
        If monthEndDate = gEndDate Then finalDateIncluded = True
    Next i

    If Not finalDateIncluded Then
        ws.Cells(chartRow, 8).Value = Format$(gEndDate, "mmm-yy")
        ws.Cells(chartRow, 43).Value = gEndDate
        For t = 2 To TENOR_COUNT
            ws.Cells(chartRow, t + 7).Value = _
                (BalanceOnDate(t, gEndDate) - BalanceOnDate(1, gEndDate)) / 1000#
        Next t
    End If

    ws.Range("H2:H" & chartRow).NumberFormat = "@"
    ws.Range("AQ2:AQ" & chartRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("I2:L" & chartRow).NumberFormat = "$0.0;[Red]($0.0);-"

    WriteRatesRow ws.Range("N1"), Array("Tenor", "Historical Average", "Current")
    ApplyRatesHeader ws.Range("N1:P1")

    For t = 2 To TENOR_COUNT
        premiumSummary(t - 1, 1) = TenorName(t)
        premiumSummary(t - 1, 2) = ThisWorkbook.Worksheets("Premium_Analysis").Cells(t + 2, 13).Value
        premiumSummary(t - 1, 3) = ThisWorkbook.Worksheets("Premium_Analysis").Cells(t + 2, 18).Value
    Next t
    ws.Range("N2:P5").Value = premiumSummary

    WriteRatesRow ws.Range("R1"), Array("Tenor", "Incremental Interest vs ON ($000)")
    ApplyRatesHeader ws.Range("R1:S1")

    For t = 2 To TENOR_COUNT
        incrementalData(t - 1, 1) = TenorName(t)
        incrementalData(t - 1, 2) = (gTotalInterest(t) - gTotalInterest(1)) / 1000#
    Next t
    ws.Range("R2:S5").Value = incrementalData

    WriteRatesRow ws.Range("U1"), _
        Array("Tenor", "Annualized Volatility (bps)", "Return vs ON (bps)")
    ApplyRatesHeader ws.Range("U1:W1")

    For t = 1 To TENOR_COUNT
        riskData(t, 1) = TenorName(t)
        riskData(t, 2) = gTenorAnnualVolBps(t)
        riskData(t, 3) = (gTenorAnnualReturnPct(t) - gTenorAnnualReturnPct(1)) * 100#
    Next t
    ws.Range("U2:W6").Value = riskData

    WriteRatesRow ws.Range("Y1"), Array("Annualized Volatility (bps)", "Return vs ON (bps)")
    ApplyRatesHeader ws.Range("Y1:Z1")

    For i = 1 To gFrontierCount
        ws.Cells(i + 1, 25).Value = gFrontier(i, 7)
        ws.Cells(i + 1, 26).Value = (gFrontier(i, 6) - gTenorAnnualReturnPct(1)) * 100#
    Next i

    WriteRatesRow ws.Range("AB1"), Array("Tenor", "Reset Volatility (bps)")
    ApplyRatesHeader ws.Range("AB1:AC1")

    For t = 2 To TENOR_COUNT
        resetVolData(t - 1, 1) = TenorName(t)
        resetVolData(t - 1, 2) = gResetSummary(t, 6)
    Next t
    ws.Range("AB2:AC5").Value = resetVolData

    WriteRatesRow ws.Range("AE1"), _
        Array("Tenor", "5th Percentile", "Median", "95th Percentile")
    ApplyRatesHeader ws.Range("AE1:AH1")

    For t = 2 To TENOR_COUNT
        resetDistribution(t - 1, 1) = TenorName(t)
        resetDistribution(t - 1, 2) = gResetSummary(t, 8)
        resetDistribution(t - 1, 3) = gResetSummary(t, 7)
        resetDistribution(t - 1, 4) = gResetSummary(t, 9)
    Next t
    ws.Range("AE2:AH5").Value = resetDistribution

    WriteRatesRow ws.Range("AJ1"), _
        Array("Tenor", "Total Interest ($MM)", "Annualized Return", _
              "Earnings Volatility (bps)", "Reset Volatility (bps)", _
              "Reset Scenarios")
    ApplyRatesHeader ws.Range("AJ1:AO1")

    For t = 1 To TENOR_COUNT
        resultSummary(t, 1) = TenorName(t)
        resultSummary(t, 2) = gTotalInterest(t) / 1000000#
        resultSummary(t, 3) = gAnnualizedReturnPct(t)
        resultSummary(t, 4) = gTenorAnnualVolBps(t)
        resultSummary(t, 5) = gResetSummary(t, 6)
        resultSummary(t, 6) = gResetSummary(t, 2)
    Next t
    ws.Range("AJ2:AO6").Value = resultSummary

    ws.Columns("A:AO").AutoFit

End Sub

Private Sub BuildDashboard()

    Dim ws As Worksheet
    Dim bestTenor As Long
    Dim worstResetTenor As Long
    Dim t As Long
    Dim summaryData(1 To TENOR_COUNT, 1 To 6) As Variant
    Dim selectedRows(1 To 3) As Long
    Dim selectedLabels As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    DeleteAllCharts ws
    ws.Cells.Clear

    ws.Range("A1:Q1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Historical Cash Investment Analysis | Actual Curve"
    ws.Range("A1:Q1").Font.Bold = True
    ws.Range("A1:Q1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:Q1").Font.Size = 17
    ws.Rows(1).RowHeight = 28

    ws.Range("A2:Q2").Interior.Color = COLOR_NAVY
    ws.Range("A2").Value = "Input-date analysis, daily reinvestment risk and historical efficient frontier"
    ws.Range("A2:Q2").Font.Color = RGB(255, 255, 255)
    ws.Range("A2:Q2").Font.Italic = True

    bestTenor = 1
    worstResetTenor = 2
    For t = 2 To TENOR_COUNT
        If gEndingValue(t) > gEndingValue(bestTenor) Then bestTenor = t
        If gResetSummary(t, 6) > gResetSummary(worstResetTenor, 6) Then worstResetTenor = t
    Next t

    FormatDashboardCard ws.Range("A4:D6"), "Analysis period", _
        Format$(gStartDate, "dd-mmm-yyyy") & " to " & Format$(gEndDate, "dd-mmm-yyyy")
    FormatDashboardCard ws.Range("E4:H6"), "Initial cash", _
        "$" & Format$(gNotional / 1000000#, "0.0") & "MM"
    FormatDashboardCard ws.Range("J4:M6"), "Highest ending value", _
        TenorName(bestTenor) & " | $" & Format$(gEndingValue(bestTenor) / 1000000#, "0.000") & "MM"
    FormatDashboardCard ws.Range("N4:Q6"), "Highest tenor reset volatility", _
        TenorName(worstResetTenor) & " | " & Format$(gResetSummary(worstResetTenor, 6), "0.0") & " bps"

    AddDashboardCharts ws

    WriteRatesRow ws.Range("A62"), _
        Array("Tenor", "Total Interest ($MM)", "Annualized Return", _
              "Earnings Volatility (bps)", "Reset Volatility (bps)", _
              "Reset Scenarios")
    ApplyRatesHeader ws.Range("A62:F62")

    For t = 1 To TENOR_COUNT
        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gTotalInterest(t) / 1000000#
        summaryData(t, 3) = gAnnualizedReturnPct(t)
        summaryData(t, 4) = gTenorAnnualVolBps(t)
        summaryData(t, 5) = gResetSummary(t, 6)
        summaryData(t, 6) = gResetSummary(t, 2)
    Next t

    ws.Range("A63:F67").Value = summaryData
    ws.Range("B63:B67").NumberFormat = "$0.000;[Red]($0.000);-"
    ws.Range("C63:C67").NumberFormat = "0.0000\%"
    ws.Range("D63:E67").NumberFormat = "0.00"
    ws.Range("F63:F67").NumberFormat = "0"

    WriteRatesRow ws.Range("H62"), _
        Array("Frontier Point", "Description", "Allocation and Liquidity")
    ApplyRatesHeader ws.Range("H62:J62")

    selectedRows(1) = 1
    selectedRows(2) = Application.WorksheetFunction.Max(1, CLng((gFrontierCount + 1) / 2))
    selectedRows(3) = gFrontierCount
    selectedLabels = Array("Minimum volatility", "Balanced frontier", _
                           "Maximum historical return")

    For i = 1 To 3
        ws.Cells(62 + i, 8).Value = selectedLabels(i - 1)
        ws.Cells(62 + i, 9).Value = _
            FrontierDescription(FrontierSegment(selectedRows(i), gFrontierCount))
        ws.Cells(62 + i, 10).Value = AllocationDescription(selectedRows(i))
        ws.Range(ws.Cells(62 + i, 8), ws.Cells(62 + i, 10)).Interior.Color = RGB(246, 248, 250)
        ws.Range(ws.Cells(62 + i, 8), ws.Cells(62 + i, 10)).WrapText = True
        ws.Cells(62 + i, 8).Font.Bold = True
        ws.Rows(62 + i).RowHeight = 48
    Next i

    ws.Range("A70:Q70").Interior.Color = COLOR_NAVY
    ws.Range("A70").Value = "Interpretation"
    ApplyRatesHeader ws.Range("A70:Q70")

    ws.Range("A71").Value = _
        "The model uses only the input start and end dates. Daily reset risk compares the same-tenor rate at every eligible start date with the rate at actual maturity."
    ws.Range("A72").Value = _
        TenorName(bestTenor) & " produced the highest ending value. Reset volatility is shown separately from aligned monthly earnings volatility used by the frontier."
    ws.Range("A73").Value = _
        "The curve chart uses actual Excel dates and never includes observations after the selected analysis end."
    ws.Range("A71:Q73").WrapText = True
    ws.Range("A71:Q73").Interior.Color = RGB(246, 248, 250)

    ws.Columns("A:Q").ColumnWidth = 11
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 22
    ws.Columns("I").ColumnWidth = 48
    ws.Columns("J").ColumnWidth = 58
    ws.Rows("71:73").RowHeight = 28

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80

End Sub

Private Sub AddDashboardCharts(ByVal dashboard As Worksheet)

    Dim dataSheet As Worksheet
    Dim chartObject As ChartObject
    Dim chartItem As Chart
    Dim seriesItem As Series
    Dim lastRow As Long
    Dim i As Long
    Dim colors As Variant

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")
    colors = Array(COLOR_NAVY, COLOR_BLUE, COLOR_TEAL, COLOR_GOLD, COLOR_PURPLE)

    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row
    Set chartObject = AddChartBox(dashboard, 10, 105, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlLine
    chartItem.SetSourceData Source:=dataSheet.Range("A1:F" & lastRow)
    ConfigureChart chartItem, "Historical deposit rates | monthly observations", True
    chartItem.Axes(xlCategory).TickLabelSpacing = 6
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "0.0\%"
    For i = 1 To chartItem.SeriesCollection.Count
        chartItem.SeriesCollection(i).Format.Line.ForeColor.RGB = colors(i - 1)
        chartItem.SeriesCollection(i).Format.Line.Weight = 1.75
    Next i

    Set chartObject = AddChartBox(dashboard, 500, 105, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlColumnClustered
    chartItem.SetSourceData Source:=dataSheet.Range("N1:P5")
    ConfigureChart chartItem, "Term premium: historical average vs current", True
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "0.0"

    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 8).End(xlUp).Row
    Set chartObject = AddChartBox(dashboard, 10, 345, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlLine
    chartItem.SetSourceData Source:=dataSheet.Range("H1:L" & lastRow)
    ConfigureChart chartItem, "Cumulative interest versus ON ($000)", True
    chartItem.Axes(xlCategory).TickLabelSpacing = 6
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "$0.0;[Red]($0.0);-"

    Set chartObject = AddChartBox(dashboard, 500, 345, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlBarClustered
    chartItem.SetSourceData Source:=dataSheet.Range("R1:S5")
    ConfigureChart chartItem, "Final incremental interest versus ON ($000)", False
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "$0.0;[Red]($0.0);-"

    Set chartObject = AddChartBox(dashboard, 10, 585, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlColumnClustered
    chartItem.SetSourceData Source:=dataSheet.Range("AB1:AC5")
    ConfigureChart chartItem, "Daily rolling reset volatility | full tenor horizon", False
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "0.0"

    Set chartObject = AddChartBox(dashboard, 500, 585, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlColumnClustered
    chartItem.SetSourceData Source:=dataSheet.Range("AE1:AH5")
    ConfigureChart chartItem, "Reset distribution | 5th, median and 95th percentile", True
    chartItem.Axes(xlValue).TickLabels.NumberFormat = "0.0"

    Set chartObject = AddChartBox(dashboard, 10, 825, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlXYScatter
    ConfigureChart chartItem, "Earnings return versus volatility | aligned month ends", True

    For i = 2 To 6
        Set seriesItem = chartItem.SeriesCollection.NewSeries
        seriesItem.Name = dataSheet.Cells(i, 21).Value
        seriesItem.XValues = dataSheet.Range(dataSheet.Cells(i, 22), dataSheet.Cells(i, 22))
        seriesItem.Values = dataSheet.Range(dataSheet.Cells(i, 23), dataSheet.Cells(i, 23))
        seriesItem.MarkerStyle = xlMarkerStyleCircle
        seriesItem.MarkerSize = 8
        seriesItem.Format.Fill.ForeColor.RGB = colors(i - 2)
        seriesItem.Format.Line.ForeColor.RGB = colors(i - 2)
    Next i

    chartItem.Axes(xlCategory).HasTitle = True
    chartItem.Axes(xlCategory).AxisTitle.Text = "Annualized earnings volatility (bps)"
    chartItem.Axes(xlValue).HasTitle = True
    chartItem.Axes(xlValue).AxisTitle.Text = "Annualized return vs ON (bps)"

    Set chartObject = AddChartBox(dashboard, 500, 825, 470, 225)
    Set chartItem = chartObject.Chart
    chartItem.ChartType = xlXYScatterLines
    ConfigureChart chartItem, "Historical efficient frontier", False

    Set seriesItem = chartItem.SeriesCollection.NewSeries
    seriesItem.Name = "Efficient frontier"
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, 25).End(xlUp).Row
    seriesItem.XValues = dataSheet.Range("Y2:Y" & lastRow)
    seriesItem.Values = dataSheet.Range("Z2:Z" & lastRow)
    seriesItem.MarkerStyle = xlMarkerStyleCircle
    seriesItem.MarkerSize = 5
    seriesItem.Format.Line.ForeColor.RGB = COLOR_NAVY
    seriesItem.Format.Line.Weight = 1.75

    chartItem.Axes(xlCategory).HasTitle = True
    chartItem.Axes(xlCategory).AxisTitle.Text = "Annualized earnings volatility (bps)"
    chartItem.Axes(xlValue).HasTitle = True
    chartItem.Axes(xlValue).AxisTitle.Text = "Annualized return vs ON (bps)"

End Sub

Private Function AddChartBox(ByVal ws As Worksheet, ByVal leftValue As Double, _
                             ByVal topValue As Double, ByVal widthValue As Double, _
                             ByVal heightValue As Double) As ChartObject

    Set AddChartBox = ws.ChartObjects.Add(Left:=leftValue, Top:=topValue, _
                                         Width:=widthValue, Height:=heightValue)

End Function

Private Sub ConfigureChart(ByVal chartItem As Chart, ByVal titleText As String, _
                           ByVal showLegend As Boolean)

    With chartItem
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = showLegend
        If showLegend Then .Legend.Position = xlLegendPositionBottom
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

End Sub

Private Sub FormatDashboardCard(ByVal target As Range, ByVal labelText As String, _
                                ByVal valueText As String)

    target.Interior.Color = COLOR_PALE
    target.Borders.LineStyle = xlContinuous
    target.Borders.Color = RGB(220, 225, 230)
    target.Borders.Weight = xlThin

    target.Cells(1, 1).Value = labelText
    target.Cells(1, 1).Font.Bold = True
    target.Cells(1, 1).Font.Color = RGB(110, 120, 130)
    target.Cells(2, 1).Value = valueText
    target.Cells(2, 1).Font.Bold = True
    target.Cells(2, 1).Font.Size = 12

End Sub

Private Sub BuildMethodology()

    Dim ws As Worksheet
    Dim lines As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Methodology")
    ws.Cells.Clear

    ws.Range("A1:H1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Methodology and Definitions"
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:H1").Font.Size = 15

    lines = Array( _
        "1. The model reads only the input start and end dates in Inputs!B5:B6.", _
        "2. The effective start is the first available curve date on or after the input start.", _
        "3. Rates are entered as ordinary numbers. 4.31 means 4.31 percent.", _
        "4. Interest uses ACT/360 and is reinvested at completed maturities.", _
        "5. ON rolls on the next curve date. Term maturities use the latest curve date on or before target.", _
        "6. Daily rolling reset risk uses every eligible curve date as a valid investment start.", _
        "7. Reset volatility is the standard deviation of same-tenor full-horizon rate changes in bps.", _
        "8. Monthly earnings volatility uses aligned month-end balances and supports covariance/frontier analysis.", _
        "9. Frontier portfolios are long-only, non-leveraged and use the weight step entered in Inputs!B8.", _
        "10. Frontier descriptions summarize historical risk, maturity and liquidity; they are not recommendations.")

    For i = LBound(lines) To UBound(lines)
        ws.Cells(i + 3, 1).Value = lines(i)
        ws.Range(ws.Cells(i + 3, 1), ws.Cells(i + 3, 8)).WrapText = True
        ws.Rows(i + 3).RowHeight = 27
    Next i

    ws.Columns("A:H").ColumnWidth = 17

End Sub

' ============================================================================
' Validation
' ============================================================================

Private Sub WriteValidationResults()

    Dim ws As Worksheet
    Dim tests As Collection
    Dim matrix As Variant
    Dim t As Long
    Dim interestDifference As Double
    Dim chartMaxDate As Double

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    ClearBody ws, "A4:D100"
    Set tests = New Collection

    tests.Add Array("Curve observations", gCurveCount, "> 0", PassFail(gCurveCount > 0))
    tests.Add Array("Effective start within input range", gStartDate, _
                    gRequestedStart & " to " & gEndDate, _
                    PassFail(gStartDate >= gRequestedStart And gStartDate <= gEndDate))
    tests.Add Array("Daily ledger rows", gDailyRowCount, TENOR_COUNT * gNumDays, _
                    PassFail(gDailyRowCount = TENOR_COUNT * gNumDays))
    tests.Add Array("Transactions", gTransactionRows.Count, "> 0", _
                    PassFail(gTransactionRows.Count > 0))
    tests.Add Array("Daily reset scenarios", gResetRows.Count, "> 0", _
                    PassFail(gResetRows.Count > 0))
    tests.Add Array("Efficient frontier points", gFrontierCount, "> 0", _
                    PassFail(gFrontierCount > 0))
    tests.Add Array("Dashboard charts", _
                    ThisWorkbook.Worksheets("Dashboard").ChartObjects.Count, 8, _
                    PassFail(ThisWorkbook.Worksheets("Dashboard").ChartObjects.Count = 8))

    chartMaxDate = Application.WorksheetFunction.Max( _
        ThisWorkbook.Worksheets("Chart_Data").Range("AP:AP"))
    tests.Add Array("Curve chart maximum date", chartMaxDate, gEndDate, _
                    PassFail(chartMaxDate <= gEndDate))

    For t = 1 To TENOR_COUNT
        interestDifference = Abs(gInterestDaySum(t) - gTotalInterest(t))
        tests.Add Array(TenorName(t) & " interest reconciliation", _
                        interestDifference, 0.05, PassFail(interestDifference <= 0.05))
    Next t

    matrix = CollectionToMatrix(tests, 4)
    ws.Range("A4").Resize(tests.Count, 4).Value = matrix
    ws.Range("D4:D" & tests.Count + 3).Font.Bold = True
    ws.Columns("A:D").AutoFit
    ws.Columns("A").ColumnWidth = 42

End Sub

Private Function ValidationPassed() As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    ValidationPassed = True
    For r = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 4).Value))) <> "PASS" Then
            ValidationPassed = False
            Exit Function
        End If
    Next r

End Function

' ============================================================================
' Helpers: curve, dates, arrays, sorting and formatting
' ============================================================================

Private Function FindHeaderRow(ByVal ws As Worksheet) As Long

    Dim r As Long

    For r = 1 To 10
        If FindHeaderColumn(ws, r, "Date") > 0 And _
           FindHeaderColumn(ws, r, "ON") > 0 And _
           FindHeaderColumn(ws, r, "1M") > 0 And _
           FindHeaderColumn(ws, r, "2M") > 0 And _
           FindHeaderColumn(ws, r, "3M") > 0 And _
           FindHeaderColumn(ws, r, "6M") > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r

End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                  ByVal headerText As String) As Long

    Dim lastColumn As Long
    Dim c As Long

    lastColumn = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastColumn
        If UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value))) = UCase$(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

End Function

Private Function FindIndexOnOrBefore(ByVal targetDate As Double) As Long

    Dim lowValue As Long
    Dim highValue As Long
    Dim middleValue As Long
    Dim resultValue As Long

    lowValue = 1
    highValue = gCurveCount
    resultValue = 0

    Do While lowValue <= highValue
        middleValue = (lowValue + highValue) \ 2
        If gCurveDates(middleValue) <= targetDate Then
            resultValue = middleValue
            lowValue = middleValue + 1
        Else
            highValue = middleValue - 1
        End If
    Loop

    FindIndexOnOrBefore = resultValue

End Function

Private Function FindIndexOnOrAfter(ByVal targetDate As Double) As Long

    Dim lowValue As Long
    Dim highValue As Long
    Dim middleValue As Long
    Dim resultValue As Long

    lowValue = 1
    highValue = gCurveCount
    resultValue = 0

    Do While lowValue <= highValue
        middleValue = (lowValue + highValue) \ 2
        If gCurveDates(middleValue) >= targetDate Then
            resultValue = middleValue
            highValue = middleValue - 1
        Else
            lowValue = middleValue + 1
        End If
    Loop

    FindIndexOnOrAfter = resultValue

End Function

Private Function FirstCurveDateOnOrAfter(ByVal targetDate As Double) As Double

    Dim indexValue As Long
    indexValue = FindIndexOnOrAfter(targetDate)
    If indexValue > 0 Then FirstCurveDateOnOrAfter = gCurveDates(indexValue)

End Function

Private Function CurveDateOnOrBefore(ByVal targetDate As Double) As Double

    Dim indexValue As Long
    indexValue = FindIndexOnOrBefore(targetDate)
    If indexValue = 0 Then
        Err.Raise vbObjectError + 600, , _
            "No curve date exists on or before " & Format$(targetDate, "dd-mmm-yyyy") & "."
    End If
    CurveDateOnOrBefore = gCurveDates(indexValue)

End Function

Private Function RateOnOrBefore(ByVal targetDate As Double, ByVal tenorIndex As Long) As Double

    Dim indexValue As Long
    indexValue = FindIndexOnOrBefore(targetDate)
    If indexValue = 0 Then
        Err.Raise vbObjectError + 601, , "No rate exists on or before the requested date."
    End If
    RateOnOrBefore = gRates(indexValue, tenorIndex)

End Function

Private Function NextCurveDateAfter(ByVal targetDate As Double) As Double

    Dim indexValue As Long
    indexValue = FindIndexOnOrAfter(targetDate + 0.0000001)
    If indexValue > 0 Then NextCurveDateAfter = gCurveDates(indexValue)

End Function

Private Function AddMonthsAnchored(ByVal anchorDate As Double, ByVal monthsToAdd As Long) As Double

    Dim anchor As Date
    Dim firstTargetMonth As Date
    Dim anchorLastDay As Long
    Dim targetLastDay As Long
    Dim targetDay As Long

    anchor = CDate(anchorDate)
    firstTargetMonth = DateSerial(Year(anchor), Month(anchor) + monthsToAdd, 1)
    anchorLastDay = Day(DateSerial(Year(anchor), Month(anchor) + 1, 0))
    targetLastDay = Day(DateSerial(Year(firstTargetMonth), Month(firstTargetMonth) + 1, 0))

    If Day(anchor) = anchorLastDay Then
        targetDay = targetLastDay
    ElseIf Day(anchor) > targetLastDay Then
        targetDay = targetLastDay
    Else
        targetDay = Day(anchor)
    End If

    AddMonthsAnchored = CDbl(DateSerial(Year(firstTargetMonth), Month(firstTargetMonth), targetDay))

End Function

Private Function TenorName(ByVal tenorIndex As Long) As String

    Select Case tenorIndex
        Case 1: TenorName = "ON"
        Case 2: TenorName = "1M"
        Case 3: TenorName = "2M"
        Case 4: TenorName = "3M"
        Case 5: TenorName = "6M"
        Case Else: TenorName = "Unknown"
    End Select

End Function

Private Function TenorMonths(ByVal tenorIndex As Long) As Long

    Select Case tenorIndex
        Case 1: TenorMonths = 0
        Case 2: TenorMonths = 1
        Case 3: TenorMonths = 2
        Case 4: TenorMonths = 3
        Case 5: TenorMonths = 6
    End Select

End Function

Private Function BlankIfZero(ByVal valueNumber As Double) As Variant
    If valueNumber = 0 Then
        BlankIfZero = Empty
    Else
        BlankIfZero = valueNumber
    End If
End Function

Private Function BalanceOnDate(ByVal tenorIndex As Long, ByVal targetDate As Double) As Double

    Dim dayIndex As Long
    dayIndex = CLng(targetDate - gStartDate)

    If dayIndex < 0 Or dayIndex > gNumDays - 1 Then
        Err.Raise vbObjectError + 602, , "Requested balance date is outside the analysis period."
    End If

    BalanceOnDate = gBalance(tenorIndex, dayIndex)

End Function

Private Function CollectionToMatrix(ByVal rows As Collection, ByVal columnCount As Long) As Variant

    Dim outputData() As Variant
    Dim rowData As Variant
    Dim r As Long
    Dim c As Long

    ReDim outputData(1 To rows.Count, 1 To columnCount)

    For r = 1 To rows.Count
        rowData = rows(r)
        For c = 1 To columnCount
            outputData(r, c) = rowData(LBound(rowData) + c - 1)
        Next c
    Next r

    CollectionToMatrix = outputData

End Function

Private Function SampleStdDev(ByRef values() As Double) As Double

    Dim n As Long
    Dim i As Long
    Dim averageValue As Double
    Dim sumSquared As Double

    n = UBound(values) - LBound(values) + 1
    If n < 2 Then
        SampleStdDev = 0#
        Exit Function
    End If

    For i = LBound(values) To UBound(values)
        averageValue = averageValue + values(i)
    Next i
    averageValue = averageValue / n

    For i = LBound(values) To UBound(values)
        sumSquared = sumSquared + (values(i) - averageValue) ^ 2
    Next i

    SampleStdDev = Sqr(sumSquared / (n - 1))

End Function

Private Function PercentileSorted(ByRef sortedValues() As Double, _
                                  ByVal percentileValue As Double) As Double

    Dim n As Long
    Dim positionValue As Double
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim fractionValue As Double

    n = UBound(sortedValues) - LBound(sortedValues) + 1
    positionValue = (n - 1) * percentileValue
    lowerIndex = LBound(sortedValues) + Int(positionValue)
    upperIndex = LBound(sortedValues) + Application.WorksheetFunction.RoundUp(positionValue, 0)

    If lowerIndex = upperIndex Then
        PercentileSorted = sortedValues(lowerIndex)
    Else
        fractionValue = positionValue - Int(positionValue)
        PercentileSorted = sortedValues(lowerIndex) + _
                           fractionValue * (sortedValues(upperIndex) - sortedValues(lowerIndex))
    End If

End Function

Private Function PositiveShare(ByRef values() As Double) As Double

    Dim i As Long
    Dim positiveCount As Long

    For i = LBound(values) To UBound(values)
        If values(i) > 0 Then positiveCount = positiveCount + 1
    Next i

    PositiveShare = positiveCount / (UBound(values) - LBound(values) + 1)

End Function

Private Sub SortDoubleArray(ByRef values() As Double)
    QuickSortDouble values, LBound(values), UBound(values)
End Sub

Private Sub QuickSortDouble(ByRef values() As Double, ByVal first As Long, ByVal last As Long)

    Dim lowValue As Long
    Dim highValue As Long
    Dim pivotValue As Double
    Dim tempValue As Double

    lowValue = first
    highValue = last
    pivotValue = values((first + last) \ 2)

    Do While lowValue <= highValue
        Do While values(lowValue) < pivotValue
            lowValue = lowValue + 1
        Loop
        Do While values(highValue) > pivotValue
            highValue = highValue - 1
        Loop
        If lowValue <= highValue Then
            tempValue = values(lowValue)
            values(lowValue) = values(highValue)
            values(highValue) = tempValue
            lowValue = lowValue + 1
            highValue = highValue - 1
        End If
    Loop

    If first < highValue Then QuickSortDouble values, first, highValue
    If lowValue < last Then QuickSortDouble values, lowValue, last

End Sub

Private Sub QuickSortPortfolio(ByVal first As Long, ByVal last As Long)

    Dim lowValue As Long
    Dim highValue As Long
    Dim pivotVol As Double
    Dim pivotReturn As Double

    lowValue = first
    highValue = last
    pivotVol = gPortfolio((first + last) \ 2, 7)
    pivotReturn = gPortfolio((first + last) \ 2, 6)

    Do While lowValue <= highValue
        Do While PortfolioComesBefore(lowValue, pivotVol, pivotReturn)
            lowValue = lowValue + 1
        Loop
        Do While PortfolioComesAfter(highValue, pivotVol, pivotReturn)
            highValue = highValue - 1
        Loop

        If lowValue <= highValue Then
            SwapPortfolioRows lowValue, highValue
            lowValue = lowValue + 1
            highValue = highValue - 1
        End If
    Loop

    If first < highValue Then QuickSortPortfolio first, highValue
    If lowValue < last Then QuickSortPortfolio lowValue, last

End Sub

Private Function PortfolioComesBefore(ByVal rowIndex As Long, ByVal pivotVol As Double, _
                                      ByVal pivotReturn As Double) As Boolean

    If gPortfolio(rowIndex, 7) < pivotVol - 0.0000001 Then
        PortfolioComesBefore = True
    ElseIf Abs(gPortfolio(rowIndex, 7) - pivotVol) <= 0.0000001 And _
           gPortfolio(rowIndex, 6) > pivotReturn Then
        PortfolioComesBefore = True
    End If

End Function

Private Function PortfolioComesAfter(ByVal rowIndex As Long, ByVal pivotVol As Double, _
                                     ByVal pivotReturn As Double) As Boolean

    If gPortfolio(rowIndex, 7) > pivotVol + 0.0000001 Then
        PortfolioComesAfter = True
    ElseIf Abs(gPortfolio(rowIndex, 7) - pivotVol) <= 0.0000001 And _
           gPortfolio(rowIndex, 6) < pivotReturn Then
        PortfolioComesAfter = True
    End If

End Function

Private Sub SwapPortfolioRows(ByVal firstRow As Long, ByVal secondRow As Long)

    Dim c As Long
    Dim temporaryValue As Double

    If firstRow = secondRow Then Exit Sub

    For c = 1 To 13
        temporaryValue = gPortfolio(firstRow, c)
        gPortfolio(firstRow, c) = gPortfolio(secondRow, c)
        gPortfolio(secondRow, c) = temporaryValue
    Next c

End Sub

Private Function CombinationCount(ByVal n As Long, ByVal k As Long) As Long

    Dim i As Long
    Dim resultValue As Double

    If k > n - k Then k = n - k
    resultValue = 1#

    For i = 1 To k
        resultValue = resultValue * (n - k + i) / i
    Next i

    CombinationCount = CLng(Round(resultValue, 0))

End Function

Private Function PassFail(ByVal conditionValue As Boolean) As String
    If conditionValue Then
        PassFail = "PASS"
    Else
        PassFail = "FAIL"
    End If
End Function

Private Sub ClearOutputData()

    Dim sheetNames As Variant
    Dim i As Long
    Dim ws As Worksheet

    sheetNames = Array("Data_Quality", "Transactions", "Daily_Accrual", _
                       "Premium_Analysis", "Rolling_Results", _
                       "Daily_Rolling_Reset", "Monthly_Returns", _
                       "Portfolio_Analysis", "Chart_Data", "Dashboard", _
                       "Methodology", "Test_Results")

    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetNames(i)))
        If CStr(sheetNames(i)) = "Chart_Data" Or _
           CStr(sheetNames(i)) = "Dashboard" Or _
           CStr(sheetNames(i)) = "Methodology" Then
            ws.Cells.Clear
        Else
            ClearBody ws, "A4:AZ50000"
        End If
        DeleteAllCharts ws
    Next i

End Sub

Private Sub ClearBody(ByVal ws As Worksheet, ByVal addressText As String)
    ws.Range(addressText).Clear
End Sub

Private Sub DeleteAllCharts(ByVal ws As Worksheet)

    Dim chartObject As ChartObject

    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function
