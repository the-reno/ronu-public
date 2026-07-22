Attribute VB_Name = "RatesAnalysisEngine"
Option Explicit

' ============================================================================
' Rates analysis engine
'
' Import together with Rates_Workbook_Setup.bas.
'
' Main macro:
'   RunRatesAnalysis
'
' Validation macro:
'   ValidateRatesAnalysis
'
' Model design:
'   - Processes only Inputs!B5 through Inputs!B6.
'   - Rates are ordinary numbers: 4.31 means 4.31 percent.
'   - ACT/360 simple interest, reinvested at each actual maturity.
'   - Every new term starts from the prior actual maturity.
'   - Term maturity is the latest curve date on or before the target.
'   - Tenor detail uses every eligible curve-date start.
'   - Cross-tenor statistics use common start dates.
'   - Efficient-frontier portfolios are static tenor sleeves.
'   - No implicit monthly rebalancing is assumed.
'   - Return is shown as incremental annual return versus ON.
'   - Constrained and unconstrained frontiers are both retained.
'   - Out-of-sample results use a separate validation period.
'   - No cells are merged.
' ============================================================================
'
' MODULE SECTION MAP
'
'   SECTION 01 - Execution control and model orchestration
'   SECTION 02 - Inputs, curve loading, and validation
'   SECTION 03 - Rolling-strategy return analysis
'   SECTION 04 - Tenor-rate and reinvestment-volatility analysis
'   SECTION 05 - Earnings-return and earnings-volatility analysis
'   SECTION 06 - Static-sleeve efficient-frontier analysis
'   SECTION 07 - Out-of-sample validation
'   SECTION 08 - Output tables and worksheet presentation
'   SECTION 09 - Chart data and dashboard layout
'   SECTION 10 - Model validation and reconciliation
'   SECTION 11 - Shared calculation and worksheet helpers
'
' Each analytical concept is kept in a separate section so return,
' earnings volatility, reset volatility, and portfolio optimization are not
' mixed together in the code or in the workbook outputs.
' ============================================================================

Private Const TENOR_COUNT As Long = 5
Private Const DAY_COUNT As Double = 360#
Private Const CANDIDATE_COLUMNS As Long = 19
Private Const FRONTIER_OUTPUT_COLUMNS As Long = 21

Private Const COLOR_NAVY As Long = 3809035
Private Const COLOR_BLUE As Long = 12026821
Private Const COLOR_TEAL As Long = 9413930
Private Const COLOR_GOLD As Long = 2857691
Private Const COLOR_PURPLE As Long = 12736661
Private Const COLOR_GREEN As Long = 6134316
Private Const COLOR_RED As Long = 5394116
Private Const COLOR_GRAY As Long = 10790052
Private Const COLOR_PALE As Long = 16185078

Private gStage As String
Private gChartError As String
Private gChartCount As Long

Private gCurveDates() As Double
Private gRates() As Double
Private gCurveCount As Long

Private gRequestedStart As Double
Private gStartDate As Double
Private gEndDate As Double
Private gNotional As Double
Private gWeightStep As Double
Private gWeightUnits As Long
Private gValidationSplit As Double
Private gNumDays As Long

Private gUseConstraints As Boolean
Private gMinimum30 As Double
Private gMinimum60 As Double
Private gMaximum6M As Double
Private gMaximumSingle As Double
Private gMaximumWAM As Double
Private gMinimumON As Double

Private gBalance() As Double
Private gDailyInterest() As Double
Private gOpeningPrincipal() As Double
Private gEndingValue() As Double
Private gTotalInterest() As Double
Private gAnnualizedReturnPct() As Double
Private gAverageRate() As Double
Private gCompletedTransactions() As Long
Private gPrincipalDaySum() As Double
Private gInterestDaySum() As Double

Private gDailyRows() As Variant
Private gDailyRowCount As Long
Private gTransactionRows As Collection
Private gScenarioRows As Collection
Private gTenorSummary(1 To TENOR_COUNT, 1 To 15) As Variant

Private gCommonStartIndices() As Long
Private gCommonStartCount As Long
Private gCommonReset() As Double

Private gMonthEndDates() As Double
Private gMonthBalanceIndex() As Double
Private gMonthlyReturns() As Double
Private gMonthCount As Long
Private gMonthlyCount As Long
Private gSplitMonthIndex As Long
Private gTenorMonthlyReturnPct() As Double
Private gTenorMonthlyVolBps() As Double

Private gConstrainedFrontier() As Double
Private gConstrainedFrontierCount As Long
Private gUnconstrainedFrontier() As Double
Private gUnconstrainedFrontierCount As Long
Private gTrainingFrontier() As Double
Private gTrainingFrontierCount As Long
Private gOutSampleData() As Variant

' ============================================================================
' SECTION 01 - EXECUTION CONTROL AND MODEL ORCHESTRATION
'
' Purpose:
'   Run the model in a controlled sequence, preserve user-maintained data,
'   report the exact processing stage, and restore Excel settings on exit.
' ============================================================================

Public Sub RunRatesAnalysis()
    Dim oldCalculation As XlCalculation
    Dim completionText As String

    On Error GoTo Fail

    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SetStage "preparing workbook"
    EnsureRatesWorkbook
    ResetRatesOutputs

    SetStage "loading curve and settings"
    LoadCurve
    LoadInputs
    LoadFrontierSettings
    ValidateInputs
    InitializeArrays

    SetStage "building rolling strategies"
    BuildStrategies

    SetStage "writing rolling data"
    WriteTransactions
    WriteDailyAccrual
    WriteRollingResults

    SetStage "building tenor scenarios"
    BuildTenorScenarios
    WriteTenorAnalysis

    SetStage "building aligned monthly returns"
    BuildMonthlyReturns

    SetStage "building static-sleeve frontiers"
    BuildPortfolioFrontiers
    WritePortfolioAnalysis

    SetStage "building out-of-sample validation"
    BuildOutOfSample
    WriteOutOfSample

    SetStage "building chart data"
    BuildChartData

    SetStage "building dashboard"
    BuildDashboard

    SetStage "running validation"
    WriteValidationResults

    completionText = "Rates analysis completed for " & _
                     Format$(gStartDate, "dd-mmm-yyyy") & " through " & _
                     Format$(gEndDate, "dd-mmm-yyyy") & "."

    If Len(gChartError) > 0 Then
        completionText = completionText & vbCrLf & _
                         "Calculation completed, but Excel returned a chart warning:" & _
                         vbCrLf & gChartError
    End If

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox completionText, vbInformation
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Rates analysis stopped during " & gStage & "." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub ValidateRatesAnalysis()
    If Not RatesSheetExists("Test_Results") Then
        MsgBox "Run RunRatesAnalysis first.", vbExclamation
    ElseIf ValidationPassed Then
        MsgBox "All calculation validation checks passed.", vbInformation
    Else
        MsgBox "At least one validation check failed. Review Test_Results.", _
               vbExclamation
    End If
End Sub

Private Sub SetStage(ByVal stageText As String)
    gStage = stageText
    Application.StatusBar = "Rates model: " & stageText & "..."
    DoEvents
End Sub

' ============================================================================
' SECTION 02 - INPUTS, CURVE LOADING, AND VALIDATION
'
' Purpose:
'   Read Inputs, Curve, and Frontier_Settings; enforce numeric rate inputs;
'   resolve the effective start date; and reject invalid dates, duplicated
'   curve observations, unsupported frontier steps, or invalid constraints.
' ============================================================================

Private Sub LoadInputs()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Inputs")

    If Not IsDate(ws.Range("B5").Value) Then
        Err.Raise vbObjectError + 101, , _
            "Inputs!B5 must contain the analysis start date."
    End If

    If Not IsDate(ws.Range("B6").Value) Then
        Err.Raise vbObjectError + 102, , _
            "Inputs!B6 must contain the analysis end date."
    End If

    If Not IsNumeric(ws.Range("B7").Value2) Then
        Err.Raise vbObjectError + 103, , _
            "Inputs!B7 must contain the initial notional."
    End If

    If Not IsNumeric(ws.Range("B8").Value2) Then
        Err.Raise vbObjectError + 104, , _
            "Inputs!B8 must contain the frontier step."
    End If

    If Not IsDate(ws.Range("B9").Value) Then
        Err.Raise vbObjectError + 105, , _
            "Inputs!B9 must contain the validation split date."
    End If

    gRequestedStart = CDbl(CDate(ws.Range("B5").Value))
    gEndDate = CDbl(CDate(ws.Range("B6").Value))
    gNotional = CDbl(ws.Range("B7").Value2)
    gWeightStep = CDbl(ws.Range("B8").Value2)
    gValidationSplit = CDbl(CDate(ws.Range("B9").Value))
End Sub

Private Sub LoadFrontierSettings()
    Dim ws As Worksheet
    Dim useText As String

    Set ws = ThisWorkbook.Worksheets("Frontier_Settings")

    useText = UCase$(Trim$(CStr(ws.Range("B5").Value2)))
    gUseConstraints = (useText = "YES" Or useText = "Y" Or useText = "TRUE")

    If Not IsNumeric(ws.Range("B6").Value2) Then
        Err.Raise vbObjectError + 106, , _
            "Frontier_Settings!B6 must be numeric."
    End If
    If Not IsNumeric(ws.Range("B7").Value2) Then
        Err.Raise vbObjectError + 107, , _
            "Frontier_Settings!B7 must be numeric."
    End If
    If Not IsNumeric(ws.Range("B8").Value2) Then
        Err.Raise vbObjectError + 108, , _
            "Frontier_Settings!B8 must be numeric."
    End If
    If Not IsNumeric(ws.Range("B9").Value2) Then
        Err.Raise vbObjectError + 109, , _
            "Frontier_Settings!B9 must be numeric."
    End If
    If Not IsNumeric(ws.Range("B10").Value2) Then
        Err.Raise vbObjectError + 110, , _
            "Frontier_Settings!B10 must be numeric."
    End If
    If Not IsNumeric(ws.Range("B11").Value2) Then
        Err.Raise vbObjectError + 111, , _
            "Frontier_Settings!B11 must be numeric."
    End If

    gMinimum30 = CDbl(ws.Range("B6").Value2) / 100#
    gMinimum60 = CDbl(ws.Range("B7").Value2) / 100#
    gMaximum6M = CDbl(ws.Range("B8").Value2) / 100#
    gMaximumSingle = CDbl(ws.Range("B9").Value2) / 100#
    gMaximumWAM = CDbl(ws.Range("B10").Value2)
    gMinimumON = CDbl(ws.Range("B11").Value2) / 100#
End Sub

Private Sub LoadCurve()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dateColumn As Long
    Dim tenorColumn(1 To TENOR_COUNT) As Long
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim curveIndex As Long
    Dim tenorIndex As Long
    Dim rawDate As Variant
    Dim rawRate As Variant

    Set ws = ThisWorkbook.Worksheets("Curve")
    headerRow = FindCurveHeaderRow(ws)

    If headerRow = 0 Then
        Err.Raise vbObjectError + 120, , _
            "Curve headers not found. Required: Date, ON, 1M, 2M, 3M, 6M."
    End If

    dateColumn = FindHeaderColumn(ws, headerRow, "DATE")
    tenorColumn(1) = FindHeaderColumn(ws, headerRow, "ON")
    tenorColumn(2) = FindHeaderColumn(ws, headerRow, "1M")
    tenorColumn(3) = FindHeaderColumn(ws, headerRow, "2M")
    tenorColumn(4) = FindHeaderColumn(ws, headerRow, "3M")
    tenorColumn(5) = FindHeaderColumn(ws, headerRow, "6M")

    If dateColumn = 0 Then
        Err.Raise vbObjectError + 121, , "Curve Date header is missing."
    End If

    For tenorIndex = 1 To TENOR_COUNT
        If tenorColumn(tenorIndex) = 0 Then
            Err.Raise vbObjectError + 122, , _
                "Curve " & TenorName(tenorIndex) & " header is missing."
        End If
    Next tenorIndex

    lastRow = ws.Cells(ws.Rows.Count, dateColumn).End(xlUp).Row
    If lastRow <= headerRow Then
        Err.Raise vbObjectError + 123, , "Curve contains no data rows."
    End If

    gCurveCount = 0
    For rowNumber = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, dateColumn).Value2))) > 0 Then
            gCurveCount = gCurveCount + 1
        End If
    Next rowNumber

    If gCurveCount < 2 Then
        Err.Raise vbObjectError + 124, , _
            "Curve requires at least two dated observations."
    End If

    ReDim gCurveDates(1 To gCurveCount)
    ReDim gRates(1 To gCurveCount, 1 To TENOR_COUNT)

    curveIndex = 0

    For rowNumber = headerRow + 1 To lastRow
        rawDate = ws.Cells(rowNumber, dateColumn).Value

        If Len(Trim$(CStr(rawDate))) = 0 Then GoTo NextCurveRow

        curveIndex = curveIndex + 1

        If IsDate(rawDate) Then
            gCurveDates(curveIndex) = CDbl(CDate(rawDate))
        ElseIf IsNumeric(rawDate) And CDbl(rawDate) > 0 Then
            gCurveDates(curveIndex) = CDbl(rawDate)
        Else
            Err.Raise vbObjectError + 125, , _
                "Invalid curve date on row " & rowNumber & "."
        End If

        For tenorIndex = 1 To TENOR_COUNT
            rawRate = ws.Cells(rowNumber, tenorColumn(tenorIndex)).Value2

            If Not IsNumeric(rawRate) Then
                Err.Raise vbObjectError + 126, , _
                    "Non-numeric " & TenorName(tenorIndex) & _
                    " rate on row " & rowNumber & "."
            End If

            gRates(curveIndex, tenorIndex) = CDbl(rawRate)
        Next tenorIndex

NextCurveRow:
    Next rowNumber

    QuickSortCurve 1, gCurveCount
End Sub

Private Sub ValidateInputs()
    Dim curveIndex As Long
    Dim tenorIndex As Long
    Dim unitsExact As Double
    Dim averageAbsoluteRate As Double
    Dim observationCount As Long

    If gRequestedStart > gEndDate Then
        Err.Raise vbObjectError + 130, , _
            "Analysis start date is after the end date."
    End If

    If gNotional <= 0 Then
        Err.Raise vbObjectError + 131, , _
            "Initial notional must be positive."
    End If

    For curveIndex = 2 To gCurveCount
        If gCurveDates(curveIndex) <= gCurveDates(curveIndex - 1) Then
            Err.Raise vbObjectError + 132, , _
                "Curve dates must be unique. Duplicate date: " & _
                Format$(gCurveDates(curveIndex), "dd-mmm-yyyy") & "."
        End If
    Next curveIndex

    If gRequestedStart > gCurveDates(gCurveCount) Then
        Err.Raise vbObjectError + 133, , _
            "Analysis start date is after the final curve date."
    End If

    If gEndDate > gCurveDates(gCurveCount) Then
        Err.Raise vbObjectError + 134, , _
            "Analysis end date exceeds the final curve date."
    End If

    gStartDate = FirstCurveDateOnOrAfter(gRequestedStart)

    If gStartDate = 0 Or gStartDate > gEndDate Then
        Err.Raise vbObjectError + 135, , _
            "No curve date exists in the requested analysis period."
    End If

    If gWeightStep < 5# Or gWeightStep > 25# Then
        Err.Raise vbObjectError + 136, , _
            "Frontier step must be between 5 and 25."
    End If

    unitsExact = 100# / gWeightStep
    gWeightUnits = CLng(Round(unitsExact, 0))

    If Abs(unitsExact - gWeightUnits) > 0.0000001 Then
        Err.Raise vbObjectError + 137, , _
            "Frontier step must divide 100 exactly. Use 5, 10, 20, or 25."
    End If

    ValidateConstraintPercent gMinimum30, "minimum 30-day liquidity"
    ValidateConstraintPercent gMinimum60, "minimum 60-day liquidity"
    ValidateConstraintPercent gMaximum6M, "maximum 6M allocation"
    ValidateConstraintPercent gMaximumSingle, _
                              "maximum single-tenor allocation"
    ValidateConstraintPercent gMinimumON, "minimum ON allocation"

    If gMinimum60 < gMinimum30 Then
        Err.Raise vbObjectError + 138, , _
            "Minimum 60-day liquidity cannot be below minimum 30-day liquidity."
    End If

    If gMaximumWAM < 0 Or gMaximumWAM > 6 Then
        Err.Raise vbObjectError + 139, , _
            "Maximum weighted-average maturity must be between 0 and 6 months."
    End If

    For curveIndex = 1 To gCurveCount
        For tenorIndex = 1 To TENOR_COUNT
            If gRates(curveIndex, tenorIndex) < -20# Or _
               gRates(curveIndex, tenorIndex) > 100# Then

                Err.Raise vbObjectError + 140, , _
                    "Rate outside the accepted numeric range on curve row " & _
                    curveIndex + 1 & "."
            End If

            averageAbsoluteRate = averageAbsoluteRate + _
                                  Abs(gRates(curveIndex, tenorIndex))
            observationCount = observationCount + 1
        Next tenorIndex
    Next curveIndex

    averageAbsoluteRate = averageAbsoluteRate / observationCount

    If averageAbsoluteRate < 0.25 Then
        Err.Raise vbObjectError + 141, , _
            "Rates appear to be Excel percentages. Enter 4.31 for 4.31 percent."
    End If

    gNumDays = CLng(gEndDate - gStartDate) + 1

    If gNumDays < 365 Then
        Err.Raise vbObjectError + 142, , _
            "Select at least one year for frontier and validation analysis."
    End If

    If gValidationSplit <= gStartDate Or _
       gValidationSplit >= gEndDate Then

        Err.Raise vbObjectError + 143, , _
            "Validation split date must fall inside the analysis period."
    End If
End Sub

Private Sub ValidateConstraintPercent(ByVal constraintValue As Double, _
                                      ByVal constraintName As String)
    If constraintValue < 0 Or constraintValue > 1 Then
        Err.Raise vbObjectError + 144, , _
            "Invalid " & constraintName & ". Enter a number from 0 through 100."
    End If
End Sub

Private Sub InitializeArrays()
    ReDim gBalance(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gDailyInterest(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gOpeningPrincipal(1 To TENOR_COUNT, 0 To gNumDays - 1)
    ReDim gEndingValue(1 To TENOR_COUNT)
    ReDim gTotalInterest(1 To TENOR_COUNT)
    ReDim gAnnualizedReturnPct(1 To TENOR_COUNT)
    ReDim gAverageRate(1 To TENOR_COUNT)
    ReDim gCompletedTransactions(1 To TENOR_COUNT)
    ReDim gPrincipalDaySum(1 To TENOR_COUNT)
    ReDim gInterestDaySum(1 To TENOR_COUNT)

    ReDim gDailyRows(1 To TENOR_COUNT * gNumDays, 1 To 20)
    gDailyRowCount = 0

    Set gTransactionRows = New Collection
    Set gScenarioRows = New Collection

    gChartError = vbNullString
    gChartCount = 0
End Sub

' ============================================================================
' SECTION 03 - ROLLING-STRATEGY RETURN ANALYSIS
'
' Purpose:
'   Build the realized ON, 1M, 2M, 3M, and 6M investment paths using ACT/360.
'   Interest is reinvested at each actual maturity, and each new term starts
'   from the prior actual maturity. This section produces economic balances,
'   total interest, annualized return, and incremental interest versus ON.
' ============================================================================

Private Sub BuildStrategies()
    Dim tenorIndex As Long

    For tenorIndex = 1 To TENOR_COUNT
        BuildOneStrategy tenorIndex
        DoEvents
    Next tenorIndex
End Sub

Private Sub BuildOneStrategy(ByVal tenorIndex As Long)
    Dim currentStart As Double
    Dim targetRoll As Double
    Dim actualRoll As Double
    Dim rateIndex As Long
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
        rateIndex = FindIndexOnOrBefore(currentStart)
        rateDate = gCurveDates(rateIndex)
        rateValue = gRates(rateIndex, tenorIndex)

        If tenorIndex = 1 Then
            targetRoll = currentStart + 1
            actualRoll = NextCurveDateAfter(currentStart)
        Else
            targetRoll = AddMonthsPreserveEndOfMonth( _
                currentStart, TenorMonths(tenorIndex))

            If targetRoll <= gCurveDates(gCurveCount) Then
                actualRoll = CurveDateOnOrBefore(targetRoll)
            Else
                actualRoll = 0
            End If
        End If

        If actualRoll > 0 And actualRoll <= currentStart Then
            Err.Raise vbObjectError + 200, , _
                "Invalid roll date for " & TenorName(tenorIndex) & _
                " starting " & Format$(currentStart, "dd-mmm-yyyy") & "."
        End If

        completed = (actualRoll > 0 And actualRoll <= gEndDate)

        If actualRoll > 0 Then
            transactionDays = CLng(actualRoll - currentStart)
        Else
            transactionDays = CLng(targetRoll - currentStart)
        End If

        dailyInterestValue = openingPrincipal * _
                             (rateValue / 100#) / DAY_COUNT
        fullPeriodInterest = dailyInterestValue * transactionDays
        closingPrincipal = openingPrincipal + fullPeriodInterest

        adjustmentFlag = vbNullString

        If rateDate <> currentStart Then
            adjustmentFlag = "Rate from prior curve date"
        End If

        If actualRoll > 0 And actualRoll <> targetRoll Then
            If Len(adjustmentFlag) > 0 Then
                adjustmentFlag = adjustmentFlag & "; "
            End If

            adjustmentFlag = adjustmentFlag & _
                             "Maturity adjusted to prior curve date"
        End If

        If completed Then
            statusText = "COMPLETED"
        Else
            statusText = "OPEN AT ANALYSIS END"
        End If

        AddTransactionRow tenorIndex, transactionID, currentStart, rateDate, _
                          rateValue, targetRoll, actualRoll, transactionDays, _
                          openingPrincipal, fullPeriodInterest, closingPrincipal, _
                          statusText, adjustmentFlag

        If actualRoll > 0 Then
            accrualEnd = MinimumDouble(gEndDate, actualRoll - 1)
        Else
            accrualEnd = gEndDate
        End If

        currentDay = currentStart

        Do While currentDay <= accrualEnd
            dayIndex = CLng(currentDay - gStartDate)
            daysAccrued = CLng(currentDay - currentStart) + 1

            gOpeningPrincipal(tenorIndex, dayIndex) = openingPrincipal
            gDailyInterest(tenorIndex, dayIndex) = dailyInterestValue
            gBalance(tenorIndex, dayIndex) = _
                openingPrincipal + dailyInterestValue * daysAccrued

            gPrincipalDaySum(tenorIndex) = _
                gPrincipalDaySum(tenorIndex) + openingPrincipal
            gInterestDaySum(tenorIndex) = _
                gInterestDaySum(tenorIndex) + dailyInterestValue

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
                        targetRoll, actualRoll, rateDate, rateValue, _
                        openingPrincipal, dailyInterestValue, _
                        dailyInterestValue * daysAccrued, fullPeriodInterest, _
                        InterestPaidValue(currentDay, currentStart, _
                                          priorInterestPaid), _
                        gBalance(tenorIndex, dayIndex), daysAccrued, _
                        transactionDays, DaysToRoll(currentDay, actualRoll, _
                                                    targetRoll), _
                        rollFlag, statusText, adjustmentFlag

            currentDay = currentDay + 1
        Loop

        If Not completed Then Exit Do

        gCompletedTransactions(tenorIndex) = _
            gCompletedTransactions(tenorIndex) + 1

        priorInterestPaid = fullPeriodInterest
        openingPrincipal = closingPrincipal
        currentStart = actualRoll
        transactionID = transactionID + 1
    Loop

    gEndingValue(tenorIndex) = _
        gBalance(tenorIndex, gNumDays - 1)
    gTotalInterest(tenorIndex) = _
        gEndingValue(tenorIndex) - gNotional

    gAnnualizedReturnPct(tenorIndex) = _
        ((gEndingValue(tenorIndex) / gNotional) ^ _
         (DAY_COUNT / gNumDays) - 1#) * 100#

    If gPrincipalDaySum(tenorIndex) > 0 Then
        gAverageRate(tenorIndex) = _
            gInterestDaySum(tenorIndex) * DAY_COUNT / _
            gPrincipalDaySum(tenorIndex) * 100#
    End If
End Sub

Private Function InterestPaidValue(ByVal currentDay As Double, _
                                   ByVal transactionStart As Double, _
                                   ByVal priorInterest As Double) As Double
    If currentDay = transactionStart Then
        InterestPaidValue = priorInterest
    Else
        InterestPaidValue = 0#
    End If
End Function

Private Sub AddTransactionRow(ByVal tenorIndex As Long, _
                              ByVal transactionID As Long, _
                              ByVal actualStart As Double, _
                              ByVal rateDate As Double, _
                              ByVal rateValue As Double, _
                              ByVal targetRoll As Double, _
                              ByVal actualRoll As Double, _
                              ByVal transactionDays As Long, _
                              ByVal openingPrincipal As Double, _
                              ByVal periodInterest As Double, _
                              ByVal closingPrincipal As Double, _
                              ByVal statusText As String, _
                              ByVal adjustmentFlag As String)
    Dim rowData As Variant

    rowData = Array( _
        TenorName(tenorIndex), _
        transactionID, _
        actualStart, _
        actualStart, _
        rateDate, _
        rateValue, _
        targetRoll, _
        BlankIfZero(actualRoll), _
        transactionDays, _
        openingPrincipal, _
        periodInterest, _
        closingPrincipal, _
        statusText, _
        adjustmentFlag)

    gTransactionRows.Add rowData
End Sub

Private Sub AddDailyRow(ByVal tenorIndex As Long, _
                        ByVal transactionID As Long, _
                        ByVal accrualDate As Double, _
                        ByVal transactionStart As Double, _
                        ByVal targetRoll As Double, _
                        ByVal actualRoll As Double, _
                        ByVal rateDate As Double, _
                        ByVal rateValue As Double, _
                        ByVal openingPrincipal As Double, _
                        ByVal dailyInterestValue As Double, _
                        ByVal cumulativeInterest As Double, _
                        ByVal fullPeriodInterest As Double, _
                        ByVal interestPaidToday As Double, _
                        ByVal economicBalance As Double, _
                        ByVal daysAccrued As Long, _
                        ByVal transactionDays As Long, _
                        ByVal daysToRollValue As Long, _
                        ByVal rollFlag As String, _
                        ByVal statusText As String, _
                        ByVal adjustmentFlag As String)
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

Private Function DaysToRoll(ByVal currentDate As Double, _
                            ByVal actualRoll As Double, _
                            ByVal targetRoll As Double) As Long
    If actualRoll > 0 Then
        DaysToRoll = MaximumLong(0, CLng(actualRoll - currentDate))
    Else
        DaysToRoll = MaximumLong(0, CLng(targetRoll - currentDate))
    End If
End Function

' ============================================================================
' SECTION 04 - TENOR-RATE AND REINVESTMENT-VOLATILITY ANALYSIS
'
' Purpose:
'   Treat every eligible curve date as a valid investment start; compare the
'   same-tenor quoted rate at start and actual maturity; retain all scenarios;
'   and calculate cross-tenor statistics on one common starting-date sample.
'   Reset volatility remains separate from earnings-return volatility.
' ============================================================================

Private Sub BuildTenorScenarios()
    Dim firstCurveIndex As Long
    Dim finalCurveIndex As Long
    Dim curveIndex As Long
    Dim tenorIndex As Long
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim maturityIndex As Long
    Dim startRate As Double
    Dim maturityRate As Double
    Dim actualDays As Long
    Dim horizonReturnPct As Double
    Dim interestValue As Double
    Dim resetChange As Double
    Dim commonFlag As Boolean
    Dim rowData As Variant
    Dim allCount(1 To TENOR_COUNT) As Long

    firstCurveIndex = FindIndexOnOrAfter(gStartDate)
    finalCurveIndex = FindIndexOnOrBefore(gEndDate)

    BuildCommonStartIndices firstCurveIndex, finalCurveIndex

    ReDim gCommonReset(1 To gCommonStartCount, 1 To TENOR_COUNT)

    For tenorIndex = 1 To TENOR_COUNT
        For curveIndex = firstCurveIndex To finalCurveIndex
            If tenorIndex = 1 Then
                targetMaturity = gCurveDates(curveIndex) + 1
                actualMaturity = NextCurveDateAfter(gCurveDates(curveIndex))

                If actualMaturity = 0 Or actualMaturity > gEndDate Then
                    Exit For
                End If
            Else
                targetMaturity = AddMonthsPreserveEndOfMonth( _
                    gCurveDates(curveIndex), TenorMonths(tenorIndex))

                If targetMaturity > gEndDate Then Exit For

                actualMaturity = CurveDateOnOrBefore(targetMaturity)

                If actualMaturity <= gCurveDates(curveIndex) Then
                    GoTo NextScenarioStart
                End If
            End If

            maturityIndex = FindIndexOnOrBefore(actualMaturity)
            startRate = gRates(curveIndex, tenorIndex)
            maturityRate = gRates(maturityIndex, tenorIndex)
            actualDays = CLng(actualMaturity - gCurveDates(curveIndex))
            horizonReturnPct = startRate * actualDays / DAY_COUNT
            interestValue = gNotional * horizonReturnPct / 100#
            resetChange = (maturityRate - startRate) * 100#
            commonFlag = _
                IsCommonStartIndex(curveIndex)

            rowData = Array( _
                TenorName(tenorIndex), _
                gCurveDates(curveIndex), _
                startRate, _
                targetMaturity, _
                actualMaturity, _
                maturityRate, _
                actualDays, _
                horizonReturnPct, _
                interestValue, _
                resetChange, _
                CommonFlagText(commonFlag))

            gScenarioRows.Add rowData
            allCount(tenorIndex) = allCount(tenorIndex) + 1

NextScenarioStart:
        Next curveIndex

        DoEvents
    Next tenorIndex

    BuildCommonResetMatrix
    CalculateTenorSummary allCount
End Sub

Private Sub BuildCommonStartIndices(ByVal firstCurveIndex As Long, _
                                    ByVal finalCurveIndex As Long)
    Dim curveIndex As Long
    Dim targetSixMonth As Double

    gCommonStartCount = 0

    For curveIndex = firstCurveIndex To finalCurveIndex
        targetSixMonth = AddMonthsPreserveEndOfMonth( _
            gCurveDates(curveIndex), 6)

        If targetSixMonth <= gEndDate Then
            gCommonStartCount = gCommonStartCount + 1
        Else
            Exit For
        End If
    Next curveIndex

    If gCommonStartCount < 30 Then
        Err.Raise vbObjectError + 300, , _
            "Fewer than 30 common daily starts are available for all tenors."
    End If

    ReDim gCommonStartIndices(1 To gCommonStartCount)

    gCommonStartCount = 0

    For curveIndex = firstCurveIndex To finalCurveIndex
        targetSixMonth = AddMonthsPreserveEndOfMonth( _
            gCurveDates(curveIndex), 6)

        If targetSixMonth <= gEndDate Then
            gCommonStartCount = gCommonStartCount + 1
            gCommonStartIndices(gCommonStartCount) = curveIndex
        Else
            Exit For
        End If
    Next curveIndex
End Sub

Private Sub BuildCommonResetMatrix()
    Dim commonRow As Long
    Dim tenorIndex As Long
    Dim curveIndex As Long
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim maturityIndex As Long

    For commonRow = 1 To gCommonStartCount
        curveIndex = gCommonStartIndices(commonRow)

        For tenorIndex = 1 To TENOR_COUNT
            If tenorIndex = 1 Then
                actualMaturity = NextCurveDateAfter(gCurveDates(curveIndex))
            Else
                targetMaturity = AddMonthsPreserveEndOfMonth( _
                    gCurveDates(curveIndex), TenorMonths(tenorIndex))
                actualMaturity = CurveDateOnOrBefore(targetMaturity)
            End If

            maturityIndex = FindIndexOnOrBefore(actualMaturity)

            gCommonReset(commonRow, tenorIndex) = _
                (gRates(maturityIndex, tenorIndex) - _
                 gRates(curveIndex, tenorIndex)) * 100#
        Next tenorIndex
    Next commonRow
End Sub

Private Sub CalculateTenorSummary(ByRef allCount() As Long)
    Dim tenorIndex As Long
    Dim commonRow As Long
    Dim curveIndex As Long
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim actualDays As Long
    Dim startRate As Double
    Dim horizonReturnPct As Double
    Dim interestValue As Double
    Dim resetValues() As Double
    Dim sumRate As Double
    Dim sumDays As Double
    Dim sumReturn As Double
    Dim sumInterest As Double
    Dim sumReset As Double

    For tenorIndex = 1 To TENOR_COUNT
        ReDim resetValues(1 To gCommonStartCount)

        sumRate = 0#
        sumDays = 0#
        sumReturn = 0#
        sumInterest = 0#
        sumReset = 0#

        For commonRow = 1 To gCommonStartCount
            curveIndex = gCommonStartIndices(commonRow)
            startRate = gRates(curveIndex, tenorIndex)

            If tenorIndex = 1 Then
                actualMaturity = NextCurveDateAfter(gCurveDates(curveIndex))
            Else
                targetMaturity = AddMonthsPreserveEndOfMonth( _
                    gCurveDates(curveIndex), TenorMonths(tenorIndex))
                actualMaturity = CurveDateOnOrBefore(targetMaturity)
            End If

            actualDays = CLng(actualMaturity - gCurveDates(curveIndex))
            horizonReturnPct = startRate * actualDays / DAY_COUNT
            interestValue = gNotional * horizonReturnPct / 100#

            resetValues(commonRow) = _
                gCommonReset(commonRow, tenorIndex)

            sumRate = sumRate + startRate
            sumDays = sumDays + actualDays
            sumReturn = sumReturn + horizonReturnPct
            sumInterest = sumInterest + interestValue
            sumReset = sumReset + resetValues(commonRow)
        Next commonRow

        SortDoubleArray resetValues

        gTenorSummary(tenorIndex, 1) = TenorName(tenorIndex)
        gTenorSummary(tenorIndex, 2) = allCount(tenorIndex)
        gTenorSummary(tenorIndex, 3) = gCommonStartCount
        gTenorSummary(tenorIndex, 4) = sumRate / gCommonStartCount
        gTenorSummary(tenorIndex, 5) = sumDays / gCommonStartCount
        gTenorSummary(tenorIndex, 6) = sumReturn / gCommonStartCount
        gTenorSummary(tenorIndex, 7) = sumInterest / gCommonStartCount
        gTenorSummary(tenorIndex, 8) = sumReset / gCommonStartCount
        gTenorSummary(tenorIndex, 9) = SampleStdDev(resetValues)
        gTenorSummary(tenorIndex, 10) = _
            PercentileSorted(resetValues, 0.05)
        gTenorSummary(tenorIndex, 11) = _
            PercentileSorted(resetValues, 0.5)
        gTenorSummary(tenorIndex, 12) = _
            PercentileSorted(resetValues, 0.95)
        gTenorSummary(tenorIndex, 13) = _
            resetValues(LBound(resetValues))
        gTenorSummary(tenorIndex, 14) = _
            resetValues(UBound(resetValues))
        gTenorSummary(tenorIndex, 15) = _
            PositiveShare(resetValues)
    Next tenorIndex
End Sub

Private Function IsCommonStartIndex(ByVal curveIndex As Long) As Boolean
    If gCommonStartCount = 0 Then Exit Function

    IsCommonStartIndex = _
        (curveIndex >= gCommonStartIndices(1) And _
         curveIndex <= gCommonStartIndices(gCommonStartCount))
End Function

Private Function CommonFlagText(ByVal isCommon As Boolean) As String
    If isCommon Then
        CommonFlagText = "COMMON"
    Else
        CommonFlagText = "ALL-STARTS ONLY"
    End If
End Function

' ============================================================================
' SECTION 05 - EARNINGS-RETURN AND EARNINGS-VOLATILITY ANALYSIS
'
' Purpose:
'   Convert rolling-strategy economic balances into aligned month-end returns.
'   These returns support annualized earnings return, sample volatility, and
'   covariance-consistent portfolio analysis. Quoted-rate volatility is not
'   substituted for earnings volatility in this section.
' ============================================================================

Private Sub BuildMonthlyReturns()
    Dim currentMonthEnd As Double
    Dim monthCount As Long
    Dim monthIndex As Long
    Dim tenorIndex As Long
    Dim dayIndex As Long
    Dim values() As Double
    Dim productValue As Double
    Dim outputData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 4) As Variant
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Monthly_Returns")
    SetOutputTitle ws, "Monthly Economic Returns"

    RatesWriteRow ws.Range("A3"), _
        Array("Month End", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A3:F3")

    currentMonthEnd = DateSerial( _
        Year(CDate(gStartDate)), _
        Month(CDate(gStartDate)) + 1, _
        0)

    Do While currentMonthEnd <= gEndDate
        monthCount = monthCount + 1
        currentMonthEnd = DateSerial( _
            Year(CDate(currentMonthEnd)), _
            Month(CDate(currentMonthEnd)) + 2, _
            0)
    Loop

    If monthCount < 18 Then
        Err.Raise vbObjectError + 400, , _
            "At least 18 month-end observations are required."
    End If

    gMonthCount = monthCount
    gMonthlyCount = gMonthCount - 1

    ReDim gMonthEndDates(1 To gMonthCount)
    ReDim gMonthBalanceIndex(1 To gMonthCount, 1 To TENOR_COUNT)
    ReDim gMonthlyReturns(1 To gMonthlyCount, 1 To TENOR_COUNT)
    ReDim gTenorMonthlyReturnPct(1 To TENOR_COUNT)
    ReDim gTenorMonthlyVolBps(1 To TENOR_COUNT)
    ReDim outputData(1 To gMonthlyCount, 1 To 6)

    currentMonthEnd = DateSerial( _
        Year(CDate(gStartDate)), _
        Month(CDate(gStartDate)) + 1, _
        0)

    For monthIndex = 1 To gMonthCount
        gMonthEndDates(monthIndex) = currentMonthEnd
        dayIndex = CLng(currentMonthEnd - gStartDate)

        For tenorIndex = 1 To TENOR_COUNT
            gMonthBalanceIndex(monthIndex, tenorIndex) = _
                gBalance(tenorIndex, dayIndex) / gNotional
        Next tenorIndex

        currentMonthEnd = DateSerial( _
            Year(CDate(currentMonthEnd)), _
            Month(CDate(currentMonthEnd)) + 2, _
            0)
    Next monthIndex

    For monthIndex = 1 To gMonthlyCount
        outputData(monthIndex, 1) = _
            gMonthEndDates(monthIndex + 1)

        For tenorIndex = 1 To TENOR_COUNT
            gMonthlyReturns(monthIndex, tenorIndex) = _
                gMonthBalanceIndex(monthIndex + 1, tenorIndex) / _
                gMonthBalanceIndex(monthIndex, tenorIndex) - 1#

            outputData(monthIndex, tenorIndex + 1) = _
                gMonthlyReturns(monthIndex, tenorIndex) * 100#
        Next tenorIndex
    Next monthIndex

    ws.Range("A4").Resize(gMonthlyCount, 6).Value = outputData
    ws.Range("A4:A" & gMonthlyCount + 3).NumberFormat = "mmm-yy"
    ws.Range("B4:F" & gMonthlyCount + 3).NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"

    RatesWriteRow ws.Range("H3"), _
        Array("Tenor", "Annualized Return", _
              "Annualized Volatility (bps)", "WAM (Months)")
    RatesStyleHeader ws.Range("H3:K3")

    For tenorIndex = 1 To TENOR_COUNT
        ReDim values(1 To gMonthlyCount)
        productValue = 1#

        For monthIndex = 1 To gMonthlyCount
            values(monthIndex) = _
                gMonthlyReturns(monthIndex, tenorIndex)

            productValue = productValue * _
                           (1# + values(monthIndex))
        Next monthIndex

        gTenorMonthlyReturnPct(tenorIndex) = _
            (productValue ^ (12# / gMonthlyCount) - 1#) * 100#

        gTenorMonthlyVolBps(tenorIndex) = _
            SampleStdDev(values) * Sqr(12#) * 10000#

        summaryData(tenorIndex, 1) = TenorName(tenorIndex)
        summaryData(tenorIndex, 2) = _
            gTenorMonthlyReturnPct(tenorIndex)
        summaryData(tenorIndex, 3) = _
            gTenorMonthlyVolBps(tenorIndex)
        summaryData(tenorIndex, 4) = _
            TenorMonths(tenorIndex)
    Next tenorIndex

    ws.Range("H4:K8").Value = summaryData
    ws.Range("I4:I8").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("J4:J8").NumberFormat = "0.00"
    ws.Range("K4:K8").NumberFormat = "0.0"

    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B:F").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 11
    ws.Columns("I:K").ColumnWidth = 20

    DetermineSplitMonthIndex
End Sub

Private Sub DetermineSplitMonthIndex()
    Dim monthIndex As Long

    gSplitMonthIndex = 0

    For monthIndex = 1 To gMonthCount
        If gMonthEndDates(monthIndex) <= gValidationSplit Then
            gSplitMonthIndex = monthIndex
        Else
            Exit For
        End If
    Next monthIndex

    If gSplitMonthIndex < 13 Then
        Err.Raise vbObjectError + 401, , _
            "Validation split leaves fewer than 12 training returns."
    End If

    If gMonthCount - gSplitMonthIndex < 6 Then
        Err.Raise vbObjectError + 402, , _
            "Validation split leaves fewer than 6 out-of-sample returns."
    End If
End Sub

' ============================================================================
' SECTION 06 - STATIC-SLEEVE EFFICIENT-FRONTIER ANALYSIS
'
' Purpose:
'   Allocate initial cash to fixed tenor sleeves that compound independently,
'   without implicit monthly rebalancing. Build constrained and unconstrained
'   frontiers using incremental annual return versus ON, earnings volatility,
'   downside risk, underperformance frequency, reset risk, WAM, and liquidity.
' ============================================================================

Private Sub BuildPortfolioFrontiers()
    Dim candidates() As Double
    Dim candidateCount As Long

    BuildCandidateSet 1, gMonthCount, candidates, candidateCount

    BuildFrontierFromCandidates candidates, candidateCount, False, _
                                gUnconstrainedFrontier, _
                                gUnconstrainedFrontierCount

    BuildFrontierFromCandidates candidates, candidateCount, True, _
                                gConstrainedFrontier, _
                                gConstrainedFrontierCount

    Erase candidates

    BuildCandidateSet 1, gSplitMonthIndex, candidates, candidateCount

    BuildFrontierFromCandidates candidates, candidateCount, True, _
                                gTrainingFrontier, _
                                gTrainingFrontierCount

    Erase candidates

    If gUnconstrainedFrontierCount = 0 Then
        Err.Raise vbObjectError + 500, , _
            "No unconstrained efficient-frontier portfolios were generated."
    End If

    If gConstrainedFrontierCount = 0 Then
        Err.Raise vbObjectError + 501, , _
            "No portfolios satisfy the current treasury constraints."
    End If

    If gTrainingFrontierCount = 0 Then
        Err.Raise vbObjectError + 502, , _
            "No estimation-period frontier was generated."
    End If
End Sub

Private Sub BuildCandidateSet(ByVal startMonthIndex As Long, _
                              ByVal endMonthIndex As Long, _
                              ByRef candidates() As Double, _
                              ByRef candidateCount As Long)
    Dim candidateCapacity As Long
    Dim w0 As Long
    Dim w1 As Long
    Dim w2 As Long
    Dim w3 As Long
    Dim w4 As Long
    Dim weights(1 To TENOR_COUNT) As Double
    Dim metrics(1 To 8) As Double
    Dim benchmarkReturnPct As Double
    Dim rowIndex As Long
    Dim tenorIndex As Long

    candidateCapacity = CountWeightCombinations(gWeightUnits + 4, 4)
    ReDim candidates(1 To candidateCapacity, 1 To CANDIDATE_COLUMNS)

    benchmarkReturnPct = WindowTenorAnnualReturnPct( _
        1, startMonthIndex, endMonthIndex)

    rowIndex = 0

    For w0 = 0 To gWeightUnits
        For w1 = 0 To gWeightUnits - w0
            For w2 = 0 To gWeightUnits - w0 - w1
                For w3 = 0 To gWeightUnits - w0 - w1 - w2
                    w4 = gWeightUnits - w0 - w1 - w2 - w3

                    weights(1) = w0 / gWeightUnits
                    weights(2) = w1 / gWeightUnits
                    weights(3) = w2 / gWeightUnits
                    weights(4) = w3 / gWeightUnits
                    weights(5) = w4 / gWeightUnits

                    CalculatePortfolioMetrics weights, _
                                              startMonthIndex, _
                                              endMonthIndex, _
                                              benchmarkReturnPct, _
                                              metrics

                    rowIndex = rowIndex + 1

                    For tenorIndex = 1 To TENOR_COUNT
                        candidates(rowIndex, tenorIndex) = _
                            weights(tenorIndex)
                    Next tenorIndex

                    candidates(rowIndex, 6) = metrics(1)
                    candidates(rowIndex, 7) = metrics(2)
                    candidates(rowIndex, 8) = metrics(3)
                    candidates(rowIndex, 9) = metrics(4)
                    candidates(rowIndex, 10) = metrics(5)
                    candidates(rowIndex, 11) = metrics(6)
                    candidates(rowIndex, 12) = metrics(7)
                    candidates(rowIndex, 13) = metrics(8)
                    candidates(rowIndex, 14) = PortfolioWAM(weights)
                    candidates(rowIndex, 15) = _
                        weights(1) + weights(2)
                    candidates(rowIndex, 16) = _
                        weights(1) + weights(2) + weights(3)
                    candidates(rowIndex, 17) = _
                        weights(1) + weights(2) + _
                        weights(3) + weights(4)

                    If PortfolioPassesConstraints(weights) Then
                        candidates(rowIndex, 18) = 1#
                    Else
                        candidates(rowIndex, 18) = 0#
                    End If

                    If metrics(3) > 0 Then
                        candidates(rowIndex, 19) = _
                            metrics(2) / metrics(3)
                    Else
                        candidates(rowIndex, 19) = 0#
                    End If
                Next w3
            Next w2
        Next w1

        If w0 Mod 3 = 0 Then DoEvents
    Next w0

    candidateCount = rowIndex
    QuickSortCandidates candidates, 1, candidateCount
End Sub

Private Sub CalculatePortfolioMetrics(ByRef weights() As Double, _
                                      ByVal startMonthIndex As Long, _
                                      ByVal endMonthIndex As Long, _
                                      ByVal benchmarkReturnPct As Double, _
                                      ByRef metrics() As Double)
    Dim returnCount As Long
    Dim portfolioReturns() As Double
    Dim excessReturns() As Double
    Dim baseIndex(1 To TENOR_COUNT) As Double
    Dim previousValue As Double
    Dim currentValue As Double
    Dim productValue As Double
    Dim downsideSum As Double
    Dim underperformCount As Long
    Dim monthIndex As Long
    Dim returnIndex As Long
    Dim tenorIndex As Long
    Dim commonRow As Long
    Dim expectedShortfallCount As Long
    Dim expectedShortfallSum As Double
    Dim portfolioReset As Double
    Dim resetSum As Double
    Dim resetSquareSum As Double
    Dim resetVariance As Double

    returnCount = endMonthIndex - startMonthIndex

    ReDim portfolioReturns(1 To returnCount)
    ReDim excessReturns(1 To returnCount)

    For tenorIndex = 1 To TENOR_COUNT
        baseIndex(tenorIndex) = _
            gMonthBalanceIndex(startMonthIndex, tenorIndex)
    Next tenorIndex

    previousValue = 1#
    productValue = 1#
    returnIndex = 0

    For monthIndex = startMonthIndex + 1 To endMonthIndex
        currentValue = 0#

        For tenorIndex = 1 To TENOR_COUNT
            currentValue = currentValue + _
                weights(tenorIndex) * _
                gMonthBalanceIndex(monthIndex, tenorIndex) / _
                baseIndex(tenorIndex)
        Next tenorIndex

        returnIndex = returnIndex + 1
        portfolioReturns(returnIndex) = _
            currentValue / previousValue - 1#

        excessReturns(returnIndex) = _
            portfolioReturns(returnIndex) - _
            (gMonthBalanceIndex(monthIndex, 1) / _
             gMonthBalanceIndex(monthIndex - 1, 1) - 1#)

        productValue = productValue * _
                       (1# + portfolioReturns(returnIndex))

        If excessReturns(returnIndex) < 0 Then
            downsideSum = downsideSum + _
                          excessReturns(returnIndex) ^ 2
            underperformCount = underperformCount + 1
        End If

        previousValue = currentValue
    Next monthIndex

    metrics(1) = _
        (productValue ^ (12# / returnCount) - 1#) * 100#
    metrics(2) = _
        (metrics(1) - benchmarkReturnPct) * 100#
    metrics(3) = _
        SampleStdDev(portfolioReturns) * Sqr(12#) * 10000#
    metrics(4) = _
        Sqr(downsideSum / returnCount) * Sqr(12#) * 10000#
    metrics(5) = underperformCount / returnCount

    SortDoubleArray excessReturns

    expectedShortfallCount = CLng( _
        Application.WorksheetFunction.RoundUp(returnCount * 0.05, 0))

    If expectedShortfallCount < 1 Then expectedShortfallCount = 1

    For returnIndex = 1 To expectedShortfallCount
        expectedShortfallSum = expectedShortfallSum + _
                               excessReturns(returnIndex)
    Next returnIndex

    metrics(6) = expectedShortfallSum / _
                 expectedShortfallCount * 10000#

    For commonRow = 1 To gCommonStartCount
        portfolioReset = 0#

        For tenorIndex = 1 To TENOR_COUNT
            portfolioReset = portfolioReset + _
                weights(tenorIndex) * _
                gCommonReset(commonRow, tenorIndex)
        Next tenorIndex

        resetSum = resetSum + portfolioReset
        resetSquareSum = resetSquareSum + _
                         portfolioReset * portfolioReset
    Next commonRow

    resetVariance = _
        (resetSquareSum - resetSum * resetSum / gCommonStartCount) / _
        (gCommonStartCount - 1)

    If resetVariance < 0 And resetVariance > -0.000000001 Then
        resetVariance = 0#
    End If

    metrics(7) = Sqr(resetVariance)
    metrics(8) = 0#

    For tenorIndex = 1 To TENOR_COUNT
        metrics(8) = metrics(8) + _
            weights(tenorIndex) * _
            CDbl(gTenorSummary(tenorIndex, 9))
    Next tenorIndex
End Sub

Private Function WindowTenorAnnualReturnPct( _
    ByVal tenorIndex As Long, _
    ByVal startMonthIndex As Long, _
    ByVal endMonthIndex As Long) As Double

    Dim monthIndex As Long
    Dim returnCount As Long
    Dim productValue As Double

    productValue = 1#
    returnCount = endMonthIndex - startMonthIndex

    For monthIndex = startMonthIndex + 1 To endMonthIndex
        productValue = productValue * _
            (gMonthBalanceIndex(monthIndex, tenorIndex) / _
             gMonthBalanceIndex(monthIndex - 1, tenorIndex))
    Next monthIndex

    WindowTenorAnnualReturnPct = _
        (productValue ^ (12# / returnCount) - 1#) * 100#
End Function

Private Function PortfolioPassesConstraints( _
    ByRef weights() As Double) As Boolean

    Dim tenorIndex As Long
    Dim maximumWeight As Double

    If Not gUseConstraints Then
        PortfolioPassesConstraints = True
        Exit Function
    End If

    For tenorIndex = 1 To TENOR_COUNT
        If weights(tenorIndex) > maximumWeight Then
            maximumWeight = weights(tenorIndex)
        End If
    Next tenorIndex

    If weights(1) < gMinimumON Then Exit Function
    If weights(5) > gMaximum6M Then Exit Function
    If maximumWeight > gMaximumSingle Then Exit Function

    If weights(1) + weights(2) < gMinimum30 Then
        Exit Function
    End If

    If weights(1) + weights(2) + weights(3) < gMinimum60 Then
        Exit Function
    End If

    If PortfolioWAM(weights) > gMaximumWAM Then
        Exit Function
    End If

    PortfolioPassesConstraints = True
End Function

Private Function PortfolioWAM(ByRef weights() As Double) As Double
    PortfolioWAM = weights(2) + _
                   2# * weights(3) + _
                   3# * weights(4) + _
                   6# * weights(5)
End Function

Private Sub BuildFrontierFromCandidates( _
    ByRef candidates() As Double, _
    ByVal candidateCount As Long, _
    ByVal constrainedOnly As Boolean, _
    ByRef frontier() As Double, _
    ByRef frontierCount As Long)

    Dim temporary() As Double
    Dim bestReturn As Double
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim includeRow As Boolean

    ReDim temporary(1 To candidateCount, 1 To CANDIDATE_COLUMNS)

    bestReturn = -1E+99
    frontierCount = 0

    For rowIndex = 1 To candidateCount
        includeRow = True

        If constrainedOnly Then
            includeRow = (candidates(rowIndex, 18) = 1#)
        End If

        If includeRow Then
            If candidates(rowIndex, 6) > bestReturn + 0.0000001 Then
                frontierCount = frontierCount + 1

                For columnIndex = 1 To CANDIDATE_COLUMNS
                    temporary(frontierCount, columnIndex) = _
                        candidates(rowIndex, columnIndex)
                Next columnIndex

                bestReturn = candidates(rowIndex, 6)
            End If
        End If
    Next rowIndex

    If frontierCount = 0 Then
        ReDim frontier(1 To 1, 1 To CANDIDATE_COLUMNS)
        Exit Sub
    End If

    ReDim frontier(1 To frontierCount, 1 To CANDIDATE_COLUMNS)

    For rowIndex = 1 To frontierCount
        For columnIndex = 1 To CANDIDATE_COLUMNS
            frontier(rowIndex, columnIndex) = _
                temporary(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
End Sub

' ============================================================================
' SECTION 07 - OUT-OF-SAMPLE VALIDATION
'
' Purpose:
'   Select representative portfolios from the estimation-period frontier and
'   apply the same weights to the later test period without re-optimizing.
'   This separates historical in-sample efficiency from realized validation.
' ============================================================================

Private Sub BuildOutOfSample()
    Dim selectedRows(1 To 3) As Long
    Dim selectedLabels As Variant
    Dim portfolioNumber As Long
    Dim frontierRow As Long
    Dim weights(1 To TENOR_COUNT) As Double
    Dim testMetrics(1 To 8) As Double
    Dim testBenchmark As Double
    Dim tenorIndex As Long

    selectedLabels = Array( _
        "Minimum Volatility", _
        "Balanced Frontier", _
        "Maximum Historical Return")

    selectedRows(1) = 1
    selectedRows(2) = _
        MaximumLong(1, (gTrainingFrontierCount + 1) \ 2)
    selectedRows(3) = gTrainingFrontierCount

    ReDim gOutSampleData(1 To 3, 1 To 22)

    testBenchmark = WindowTenorAnnualReturnPct( _
        1, gSplitMonthIndex, gMonthCount)

    For portfolioNumber = 1 To 3
        frontierRow = selectedRows(portfolioNumber)

        For tenorIndex = 1 To TENOR_COUNT
            weights(tenorIndex) = _
                gTrainingFrontier(frontierRow, tenorIndex)
        Next tenorIndex

        CalculatePortfolioMetrics weights, _
                                  gSplitMonthIndex, _
                                  gMonthCount, _
                                  testBenchmark, _
                                  testMetrics

        gOutSampleData(portfolioNumber, 1) = _
            selectedLabels(portfolioNumber - 1)
        gOutSampleData(portfolioNumber, 2) = _
            FrontierSegment(frontierRow, gTrainingFrontierCount)

        For tenorIndex = 1 To TENOR_COUNT
            gOutSampleData(portfolioNumber, tenorIndex + 2) = _
                weights(tenorIndex) * 100#
        Next tenorIndex

        gOutSampleData(portfolioNumber, 8) = _
            gTrainingFrontier(frontierRow, 14)
        gOutSampleData(portfolioNumber, 9) = _
            gTrainingFrontier(frontierRow, 15) * 100#
        gOutSampleData(portfolioNumber, 10) = _
            gTrainingFrontier(frontierRow, 6)
        gOutSampleData(portfolioNumber, 11) = _
            gTrainingFrontier(frontierRow, 7)
        gOutSampleData(portfolioNumber, 12) = _
            gTrainingFrontier(frontierRow, 8)
        gOutSampleData(portfolioNumber, 13) = _
            gTrainingFrontier(frontierRow, 9)
        gOutSampleData(portfolioNumber, 14) = _
            gTrainingFrontier(frontierRow, 10)
        gOutSampleData(portfolioNumber, 15) = testMetrics(1)
        gOutSampleData(portfolioNumber, 16) = testMetrics(2)
        gOutSampleData(portfolioNumber, 17) = testMetrics(3)
        gOutSampleData(portfolioNumber, 18) = testMetrics(4)
        gOutSampleData(portfolioNumber, 19) = testMetrics(5)
        gOutSampleData(portfolioNumber, 20) = testMetrics(6)
        gOutSampleData(portfolioNumber, 21) = _
            gNotional * testMetrics(2) / 10000#
        gOutSampleData(portfolioNumber, 22) = _
            AllocationText(weights)
    Next portfolioNumber
End Sub

' ============================================================================
' SECTION 08 - OUTPUT TABLES AND WORKSHEET PRESENTATION
'
' Purpose:
'   Write each analysis to its dedicated output sheet. User-maintained sheets
'   are never cleared here. Large data blocks are written in arrays to reduce
'   Excel calls and lower the risk of the model stopping mid-process.
' ============================================================================

Private Sub WriteTransactions()
    Dim ws As Worksheet
    Dim matrix As Variant
    Dim rowCount As Long

    Set ws = ThisWorkbook.Worksheets("Transactions")
    SetOutputTitle ws, "Transaction Schedule"

    RatesWriteRow ws.Range("A3"), _
        Array("Tenor", "Transaction ID", "Target Start Date", _
              "Actual Start Date", "Rate Observation Date", "Rate Used", _
              "Target Roll Date", "Actual Roll Date", "Transaction Days", _
              "Opening Notional ($)", "Period Interest ($)", _
              "Closing Notional ($)", "Status", "Adjustment Flag")

    RatesStyleHeader ws.Range("A3:N3")

    rowCount = gTransactionRows.Count

    If rowCount > 0 Then
        matrix = CollectionToMatrix(gTransactionRows, 14)
        ws.Range("A4").Resize(rowCount, 14).Value = matrix

        ws.Range("C4:E" & rowCount + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("G4:H" & rowCount + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("F4:F" & rowCount + 3).NumberFormat = "0.0000"
        ws.Range("J4:L" & rowCount + 3).NumberFormat = _
            "$#,##0;[Red]($#,##0);-"
    End If

    SetFixedWidths ws, Array(10, 12, 14, 14, 16, 12, 14, 14, _
                             13, 18, 18, 18, 22, 40)
End Sub

Private Sub WriteDailyAccrual()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Daily_Accrual")
    SetOutputTitle ws, "Daily Accrual Ledger"

    RatesWriteRow ws.Range("A3"), _
        Array("Accrual Date", "Tenor", "Transaction ID", _
              "Transaction Start Date", "Target Roll Date", _
              "Actual Roll Date", "Rate Observation Date", "Rate Used", _
              "Opening Notional ($)", "Daily Interest ($)", _
              "Cumulative Period Interest ($)", "Full Period Interest ($)", _
              "Interest Paid Today ($)", "Economic Balance ($)", _
              "Days Accrued", "Transaction Days", "Days to Roll", _
              "Roll Flag", "Status", "Adjustment Flag")

    RatesStyleHeader ws.Range("A3:T3")

    If gDailyRowCount > 0 Then
        ws.Range("A4").Resize(gDailyRowCount, 20).Value = gDailyRows
        ws.Range("A4:A" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("D4:G" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("H4:H" & gDailyRowCount + 3).NumberFormat = "0.0000"
        ws.Range("I4:N" & gDailyRowCount + 3).NumberFormat = _
            "$#,##0;[Red]($#,##0);-"
    End If

    SetFixedWidths ws, Array(14, 10, 12, 16, 14, 14, 16, 11, 18, 16, _
                             20, 18, 18, 18, 12, 14, 12, 18, 22, 40)
End Sub

Private Sub WriteRollingResults()
    Dim ws As Worksheet
    Dim growthData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 8) As Variant
    Dim dayIndex As Long
    Dim tenorIndex As Long

    Set ws = ThisWorkbook.Worksheets("Rolling_Results")
    SetOutputTitle ws, "Rolling Investment Results"

    RatesWriteRow ws.Range("A3"), _
        Array("Date", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A3:F3")

    ReDim growthData(1 To gNumDays, 1 To 6)

    For dayIndex = 0 To gNumDays - 1
        growthData(dayIndex + 1, 1) = gStartDate + dayIndex

        For tenorIndex = 1 To TENOR_COUNT
            growthData(dayIndex + 1, tenorIndex + 1) = _
                gBalance(tenorIndex, dayIndex)
        Next tenorIndex
    Next dayIndex

    ws.Range("A4").Resize(gNumDays, 6).Value = growthData
    ws.Range("A4:A" & gNumDays + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("B4:F" & gNumDays + 3).NumberFormat = _
        "$#,##0;[Red]($#,##0);-"

    RatesWriteRow ws.Range("H3"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", _
              "Total Return", "Annualized Return", _
              "Average Invested Rate", "Completed Transactions", _
              "Incremental Interest vs ON ($)")

    RatesStyleHeader ws.Range("H3:O3")

    For tenorIndex = 1 To TENOR_COUNT
        summaryData(tenorIndex, 1) = TenorName(tenorIndex)
        summaryData(tenorIndex, 2) = gEndingValue(tenorIndex)
        summaryData(tenorIndex, 3) = gTotalInterest(tenorIndex)
        summaryData(tenorIndex, 4) = _
            (gEndingValue(tenorIndex) / gNotional - 1#) * 100#
        summaryData(tenorIndex, 5) = _
            gAnnualizedReturnPct(tenorIndex)
        summaryData(tenorIndex, 6) = gAverageRate(tenorIndex)
        summaryData(tenorIndex, 7) = _
            gCompletedTransactions(tenorIndex)
        summaryData(tenorIndex, 8) = _
            gTotalInterest(tenorIndex) - gTotalInterest(1)
    Next tenorIndex

    ws.Range("H4:O8").Value = summaryData
    ws.Range("I4:J8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("K4:M8").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("N4:N8").NumberFormat = "0"
    ws.Range("O4:O8").NumberFormat = "$#,##0;[Red]($#,##0);-"

    SetFixedWidths ws, Array(14, 18, 18, 18, 18, 18, 3, 11, 18, 18, _
                             15, 15, 16, 18, 22)
End Sub

Private Sub WriteTenorAnalysis()
    Dim ws As Worksheet
    Dim detailMatrix As Variant
    Dim summaryMatrix(1 To TENOR_COUNT, 1 To 15) As Variant
    Dim tenorIndex As Long
    Dim columnIndex As Long

    Set ws = ThisWorkbook.Worksheets("Tenor_Analysis")
    SetOutputTitle ws, "Daily Tenor Scenario Analysis"

    RatesWriteRow ws.Range("A3"), _
        Array("Tenor", "Start Date", "Start Rate", "Target Maturity", _
              "Actual Maturity", "Maturity Rate", "Actual Days", _
              "Horizon Return", "Interest on Input Notional ($)", _
              "Reset Change (bps)", "Comparison Sample")

    RatesStyleHeader ws.Range("A3:K3")

    If gScenarioRows.Count > 0 Then
        detailMatrix = CollectionToMatrix(gScenarioRows, 11)
        ws.Range("A4").Resize(gScenarioRows.Count, 11).Value = detailMatrix

        ws.Range("B4:B" & gScenarioRows.Count + 3).NumberFormat = _
            "mm/dd/yyyy"
        ws.Range("D4:E" & gScenarioRows.Count + 3).NumberFormat = _
            "mm/dd/yyyy"
        ws.Range("C4:C" & gScenarioRows.Count + 3).NumberFormat = _
            "0.0000"
        ws.Range("F4:F" & gScenarioRows.Count + 3).NumberFormat = _
            "0.0000"
        ws.Range("H4:H" & gScenarioRows.Count + 3).NumberFormat = _
            "0.0000\%;[Red](0.0000\%);-"
        ws.Range("I4:I" & gScenarioRows.Count + 3).NumberFormat = _
            "$#,##0;[Red]($#,##0);-"
        ws.Range("J4:J" & gScenarioRows.Count + 3).NumberFormat = _
            "0.0;[Red](0.0);-"
    End If

    RatesWriteRow ws.Range("M3"), _
        Array("Tenor", "All Daily Starts", "Common Daily Starts", _
              "Average Quoted Rate", "Average Holding Days", _
              "Average Horizon Return", "Average Interest ($)", _
              "Average Reset (bps)", "Reset Volatility (bps)", _
              "5th Percentile (bps)", "Median (bps)", _
              "95th Percentile (bps)", "Worst Reset (bps)", _
              "Best Reset (bps)", "Positive Resets")

    RatesStyleHeader ws.Range("M3:AA3")

    For tenorIndex = 1 To TENOR_COUNT
        For columnIndex = 1 To 15
            summaryMatrix(tenorIndex, columnIndex) = _
                gTenorSummary(tenorIndex, columnIndex)
        Next columnIndex
    Next tenorIndex

    ws.Range("M4:AA8").Value = summaryMatrix
    ws.Range("N4:O8").NumberFormat = "0"
    ws.Range("P4:P8").NumberFormat = "0.0000"
    ws.Range("Q4:Q8").NumberFormat = "0.0"
    ws.Range("R4:R8").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("S4:S8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("T4:Z8").NumberFormat = "0.0;[Red](0.0);-"
    ws.Range("AA4:AA8").NumberFormat = "0.0%"

    SetFixedWidths ws, Array(10, 14, 12, 14, 14, 13, 11, 14, 20, 14, _
                             18, 3, 10, 14, 15, 16, 16, 16, 18, 16, _
                             16, 15, 15, 15, 15, 15, 15)
End Sub

Private Sub WritePortfolioAnalysis()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Portfolio_Analysis")
    SetOutputTitle ws, "Static-Sleeve Efficient Frontier"

    WriteFrontierTable ws, "A3", gConstrainedFrontier, _
                       gConstrainedFrontierCount, _
                       "Treasury-Constrained Frontier"

    WriteFrontierTable ws, "W3", gUnconstrainedFrontier, _
                       gUnconstrainedFrontierCount, _
                       "Unconstrained Historical Frontier"

    ws.Columns("A:AQ").ColumnWidth = 12
    ws.Columns("T:U").ColumnWidth = 40
    ws.Columns("AP:AQ").ColumnWidth = 40
End Sub

Private Sub WriteFrontierTable(ByVal ws As Worksheet, _
                               ByVal startAddress As String, _
                               ByRef frontier() As Double, _
                               ByVal frontierCount As Long, _
                               ByVal tableTitle As String)
    Dim startCell As Range
    Dim titleCell As Range
    Dim headerCell As Range
    Dim outputData() As Variant
    Dim frontierRow As Long
    Dim outputRow As Long
    Dim tenorIndex As Long
    Dim segmentText As String
    Dim descriptionText As String
    Dim weights(1 To TENOR_COUNT) As Double

    Set startCell = ws.Range(startAddress)
    Set titleCell = startCell
    Set headerCell = startCell.Offset(1, 0)

    titleCell.Value = tableTitle
    titleCell.Resize(1, FRONTIER_OUTPUT_COLUMNS).Interior.Color = COLOR_NAVY
    RatesStyleTitle titleCell.Resize(1, FRONTIER_OUTPUT_COLUMNS)

    RatesWriteRow headerCell, _
        Array("Rank", "Segment", "Annualized Return", _
              "Incremental Return vs ON (bps)", _
              "Earnings Volatility (bps)", "Downside Deviation (bps)", _
              "Months Underperforming ON", _
              "Worst 5% Average vs ON (Monthly bps)", _
              "Portfolio Reset Volatility (bps)", _
              "Conservative Reset Score (bps)", _
              "ON Weight", "1M Weight", "2M Weight", _
              "3M Weight", "6M Weight", "WAM (Months)", _
              "Available <=30D", "Available <=60D", _
              "Available <=90D", "Description", "Allocation Summary")

    RatesStyleHeader headerCell.Resize(1, FRONTIER_OUTPUT_COLUMNS)

    ReDim outputData(1 To frontierCount, 1 To FRONTIER_OUTPUT_COLUMNS)

    For frontierRow = 1 To frontierCount
        outputRow = frontierRow
        segmentText = FrontierSegment(frontierRow, frontierCount)
        descriptionText = FrontierDescription(segmentText)

        For tenorIndex = 1 To TENOR_COUNT
            weights(tenorIndex) = frontier(frontierRow, tenorIndex)
        Next tenorIndex

        outputData(outputRow, 1) = frontierRow
        outputData(outputRow, 2) = segmentText
        outputData(outputRow, 3) = frontier(frontierRow, 6)
        outputData(outputRow, 4) = frontier(frontierRow, 7)
        outputData(outputRow, 5) = frontier(frontierRow, 8)
        outputData(outputRow, 6) = frontier(frontierRow, 9)
        outputData(outputRow, 7) = frontier(frontierRow, 10)
        outputData(outputRow, 8) = frontier(frontierRow, 11)
        outputData(outputRow, 9) = frontier(frontierRow, 12)
        outputData(outputRow, 10) = frontier(frontierRow, 13)

        For tenorIndex = 1 To TENOR_COUNT
            outputData(outputRow, tenorIndex + 10) = _
                frontier(frontierRow, tenorIndex) * 100#
        Next tenorIndex

        outputData(outputRow, 16) = frontier(frontierRow, 14)
        outputData(outputRow, 17) = frontier(frontierRow, 15) * 100#
        outputData(outputRow, 18) = frontier(frontierRow, 16) * 100#
        outputData(outputRow, 19) = frontier(frontierRow, 17) * 100#
        outputData(outputRow, 20) = descriptionText
        outputData(outputRow, 21) = AllocationText(weights)
    Next frontierRow

    headerCell.Offset(1, 0).Resize( _
        frontierCount, FRONTIER_OUTPUT_COLUMNS).Value = outputData

    headerCell.Offset(1, 2).Resize(frontierCount, 1).NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    headerCell.Offset(1, 3).Resize(frontierCount, 7).NumberFormat = _
        "0.00;[Red](0.00);-"
    headerCell.Offset(1, 10).Resize(frontierCount, 5).NumberFormat = _
        "0.0\%"
    headerCell.Offset(1, 15).Resize(frontierCount, 1).NumberFormat = _
        "0.0"
    headerCell.Offset(1, 16).Resize(frontierCount, 3).NumberFormat = _
        "0.0\%"
    headerCell.Offset(1, 19).Resize(frontierCount, 2).WrapText = True
End Sub

Private Sub WriteOutOfSample()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Out_of_Sample")
    SetOutputTitle ws, "Out-of-Sample Frontier Validation"

    ws.Range("A3").Value = "Estimation Period"
    ws.Range("B3").Value = _
        Format$(gMonthEndDates(1), "dd-mmm-yyyy") & " to " & _
        Format$(gMonthEndDates(gSplitMonthIndex), "dd-mmm-yyyy")

    ws.Range("A4").Value = "Out-of-Sample Period"
    ws.Range("B4").Value = _
        Format$(gMonthEndDates(gSplitMonthIndex), "dd-mmm-yyyy") & _
        " to " & Format$(gMonthEndDates(gMonthCount), "dd-mmm-yyyy")

    ws.Range("A3:B4").Interior.Color = COLOR_PALE
    ws.Range("A3:A4").Font.Bold = True

    RatesWriteRow ws.Range("A7"), _
        Array("Portfolio", "Training Segment", "ON Weight", "1M Weight", _
              "2M Weight", "3M Weight", "6M Weight", "WAM (Months)", _
              "Available <=30D", "Training Annual Return", _
              "Training Incremental vs ON (bps)", _
              "Training Volatility (bps)", "Training Downside (bps)", _
              "Training Underperformance", "Test Annual Return", _
              "Test Incremental vs ON (bps)", "Test Volatility (bps)", _
              "Test Downside (bps)", "Test Underperformance", _
              "Test Worst 5% Average (Monthly bps)", _
              "Annual Dollar Pickup vs ON ($)", "Allocation Summary")

    RatesStyleHeader ws.Range("A7:V7")
    ws.Range("A8:V10").Value = gOutSampleData

    ws.Range("C8:G10").NumberFormat = "0.0\%"
    ws.Range("H8:H10").NumberFormat = "0.0"
    ws.Range("I8:I10").NumberFormat = "0.0\%"
    ws.Range("J8:J10").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("K8:M10").NumberFormat = "0.00;[Red](0.00);-"
    ws.Range("N8:N10").NumberFormat = "0.0%"
    ws.Range("O8:O10").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("P8:R10").NumberFormat = "0.00;[Red](0.00);-"
    ws.Range("S8:S10").NumberFormat = "0.0%"
    ws.Range("T8:T10").NumberFormat = "0.00;[Red](0.00);-"
    ws.Range("U8:U10").NumberFormat = _
        "$#,##0;[Red]($#,##0);-"
    ws.Range("V8:V10").WrapText = True

    SetFixedWidths ws, Array(23, 20, 11, 11, 11, 11, 11, 13, 15, 17, _
                             19, 17, 17, 18, 17, 18, 16, 16, 17, 20, _
                             20, 48)
End Sub

' ============================================================================
' SECTION 09 - CHART DATA AND DASHBOARD LAYOUT
'
' Purpose:
'   Build compact chart-source ranges and create a fixed six-chart dashboard.
'   Chart errors are isolated from calculations so a chart warning does not
'   invalidate completed return, volatility, tenor, or frontier results.
' ============================================================================

Private Sub BuildChartData()
    Dim ws As Worksheet
    Dim firstCurveIndex As Long
    Dim finalCurveIndex As Long
    Dim curveIndex As Long
    Dim tenorIndex As Long
    Dim chartRow As Long
    Dim previousMonth As Long
    Dim previousYear As Long
    Dim lastMonthIndex As Long
    Dim monthIndex As Long
    Dim finalDateIncluded As Boolean
    Dim selectedRow As Long

    Set ws = ThisWorkbook.Worksheets("Chart_Data")
    SetOutputTitle ws, "Chart Data"

    firstCurveIndex = FindIndexOnOrAfter(gStartDate)
    finalCurveIndex = FindIndexOnOrBefore(gEndDate)

    RatesWriteRow ws.Range("A3"), _
        Array("Month", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A3:F3")

    chartRow = 4
    previousMonth = -1
    previousYear = -1

    For curveIndex = firstCurveIndex To finalCurveIndex
        If previousMonth = -1 Then
            previousMonth = Month(CDate(gCurveDates(curveIndex)))
            previousYear = Year(CDate(gCurveDates(curveIndex)))
            lastMonthIndex = curveIndex

        ElseIf Month(CDate(gCurveDates(curveIndex))) <> previousMonth Or _
               Year(CDate(gCurveDates(curveIndex))) <> previousYear Then

            WriteCurveChartRow ws, chartRow, lastMonthIndex
            chartRow = chartRow + 1

            previousMonth = Month(CDate(gCurveDates(curveIndex)))
            previousYear = Year(CDate(gCurveDates(curveIndex)))
            lastMonthIndex = curveIndex
        Else
            lastMonthIndex = curveIndex
        End If
    Next curveIndex

    WriteCurveChartRow ws, chartRow, lastMonthIndex

    RatesWriteRow ws.Range("H3"), _
        Array("Tenor", "Incremental Interest vs ON ($000)")
    RatesStyleHeader ws.Range("H3:I3")

    For tenorIndex = 1 To TENOR_COUNT
        ws.Cells(tenorIndex + 3, 8).Value = TenorName(tenorIndex)
        ws.Cells(tenorIndex + 3, 9).Value = _
            (gTotalInterest(tenorIndex) - gTotalInterest(1)) / 1000#
    Next tenorIndex

    RatesWriteRow ws.Range("K3"), _
        Array("Tenor", "Average Rate Premium vs ON (bps)")
    RatesStyleHeader ws.Range("K3:L3")

    For tenorIndex = 1 To TENOR_COUNT
        ws.Cells(tenorIndex + 3, 11).Value = TenorName(tenorIndex)
        ws.Cells(tenorIndex + 3, 12).Value = _
            (CDbl(gTenorSummary(tenorIndex, 4)) - _
             CDbl(gTenorSummary(1, 4))) * 100#
    Next tenorIndex

    RatesWriteRow ws.Range("N3"), _
        Array("Constrained Volatility (bps)", _
              "Constrained Incremental Return (bps)")
    RatesStyleHeader ws.Range("N3:O3")

    For chartRow = 1 To gConstrainedFrontierCount
        ws.Cells(chartRow + 3, 14).Value = _
            gConstrainedFrontier(chartRow, 8)
        ws.Cells(chartRow + 3, 15).Value = _
            gConstrainedFrontier(chartRow, 7)
    Next chartRow

    RatesWriteRow ws.Range("Q3"), _
        Array("Unconstrained Volatility (bps)", _
              "Unconstrained Incremental Return (bps)")
    RatesStyleHeader ws.Range("Q3:R3")

    For chartRow = 1 To gUnconstrainedFrontierCount
        ws.Cells(chartRow + 3, 17).Value = _
            gUnconstrainedFrontier(chartRow, 8)
        ws.Cells(chartRow + 3, 18).Value = _
            gUnconstrainedFrontier(chartRow, 7)
    Next chartRow

    RatesWriteRow ws.Range("T3"), _
        Array("Portfolio Reset Volatility (bps)", _
              "Incremental Return (bps)")
    RatesStyleHeader ws.Range("T3:U3")

    For chartRow = 1 To gConstrainedFrontierCount
        ws.Cells(chartRow + 3, 20).Value = _
            gConstrainedFrontier(chartRow, 12)
        ws.Cells(chartRow + 3, 21).Value = _
            gConstrainedFrontier(chartRow, 7)
    Next chartRow

    RatesWriteRow ws.Range("W3"), _
        Array("Portfolio", "Training Incremental (bps)", _
              "Test Incremental (bps)")
    RatesStyleHeader ws.Range("W3:Y3")

    For chartRow = 1 To 3
        ws.Cells(chartRow + 3, 23).Value = _
            gOutSampleData(chartRow, 1)
        ws.Cells(chartRow + 3, 24).Value = _
            gOutSampleData(chartRow, 11)
        ws.Cells(chartRow + 3, 25).Value = _
            gOutSampleData(chartRow, 16)
    Next chartRow

    RatesWriteRow ws.Range("AA3"), _
        Array("Month", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("AA3:AE3")

    chartRow = 4

    For monthIndex = 1 To gMonthCount
        ws.Cells(chartRow, 27).Value = _
            Format$(gMonthEndDates(monthIndex), "mmm-yy")

        For tenorIndex = 2 To TENOR_COUNT
            ws.Cells(chartRow, tenorIndex + 26).Value = _
                (BalanceOnDate(tenorIndex, gMonthEndDates(monthIndex)) - _
                 BalanceOnDate(1, gMonthEndDates(monthIndex))) / 1000#
        Next tenorIndex

        chartRow = chartRow + 1

        If gMonthEndDates(monthIndex) = gEndDate Then
            finalDateIncluded = True
        End If
    Next monthIndex

    If Not finalDateIncluded Then
        ws.Cells(chartRow, 27).Value = _
            Format$(gEndDate, "mmm-yy")

        For tenorIndex = 2 To TENOR_COUNT
            ws.Cells(chartRow, tenorIndex + 26).Value = _
                (BalanceOnDate(tenorIndex, gEndDate) - _
                 BalanceOnDate(1, gEndDate)) / 1000#
        Next tenorIndex
    End If

    RatesWriteRow ws.Range("AG3"), _
        Array("Selected Portfolio", "Incremental Return (bps)", _
              "Volatility (bps)", "Reset Volatility (bps)")
    RatesStyleHeader ws.Range("AG3:AJ3")

    For chartRow = 1 To 3
        selectedRow = SelectedFrontierRow( _
            chartRow, gConstrainedFrontierCount)

        ws.Cells(chartRow + 3, 33).Value = _
            SelectedFrontierLabel(chartRow)
        ws.Cells(chartRow + 3, 34).Value = _
            gConstrainedFrontier(selectedRow, 7)
        ws.Cells(chartRow + 3, 35).Value = _
            gConstrainedFrontier(selectedRow, 8)
        ws.Cells(chartRow + 3, 36).Value = _
            gConstrainedFrontier(selectedRow, 12)
    Next chartRow

    ws.Columns("A:AJ").ColumnWidth = 13
End Sub

Private Sub WriteCurveChartRow(ByVal ws As Worksheet, _
                               ByVal chartRow As Long, _
                               ByVal curveIndex As Long)
    Dim tenorIndex As Long

    ws.Cells(chartRow, 1).Value = _
        Format$(gCurveDates(curveIndex), "mmm-yy")

    For tenorIndex = 1 To TENOR_COUNT
        ws.Cells(chartRow, tenorIndex + 1).Value = _
            gRates(curveIndex, tenorIndex)
    Next tenorIndex
End Sub

Private Sub BuildDashboard()
    Dim ws As Worksheet
    Dim bestTenor As Long
    Dim selectedRow As Long
    Dim summaryData(1 To TENOR_COUNT, 1 To 6) As Variant
    Dim selectedData(1 To 3, 1 To 8) As Variant
    Dim selectedNumber As Long
    Dim tenorIndex As Long
    Dim weights(1 To TENOR_COUNT) As Double

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    SetOutputTitle ws, _
        "Historical Cash Investment Analysis | Static-Sleeve Frontier"

    ws.Range("A2:Q2").Interior.Color = COLOR_NAVY
    ws.Range("A2").Value = _
        "Input-date rolling returns, common-start tenor risk, constraints, and validation"
    ws.Range("A2:Q2").Font.Color = RGB(255, 255, 255)
    ws.Range("A2:Q2").Font.Italic = True

    bestTenor = 1

    For tenorIndex = 2 To TENOR_COUNT
        If gEndingValue(tenorIndex) > gEndingValue(bestTenor) Then
            bestTenor = tenorIndex
        End If
    Next tenorIndex

    FormatDashboardCard ws, "A4:D6", "Analysis period", _
        Format$(gStartDate, "dd-mmm-yyyy") & " to " & _
        Format$(gEndDate, "dd-mmm-yyyy")

    FormatDashboardCard ws, "E4:H6", "Initial cash", _
        "$" & Format$(gNotional / 1000000#, "0.0") & "MM"

    FormatDashboardCard ws, "J4:M6", "Highest rolling ending value", _
        TenorName(bestTenor) & " | $" & _
        Format$(gEndingValue(bestTenor) / 1000000#, "0.000") & "MM"

    FormatDashboardCard ws, "N4:Q6", "Common tenor sample", _
        Format$(gCommonStartCount, "#,##0") & _
        " daily starts per tenor"

    BuildChartsSafe ws

    RatesWriteRow ws.Range("A54"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", _
              "Annualized Rolling Return", "Monthly Volatility (bps)", _
              "Common-Sample Reset Volatility (bps)")

    RatesStyleHeader ws.Range("A54:F54")

    For tenorIndex = 1 To TENOR_COUNT
        summaryData(tenorIndex, 1) = TenorName(tenorIndex)
        summaryData(tenorIndex, 2) = gEndingValue(tenorIndex)
        summaryData(tenorIndex, 3) = gTotalInterest(tenorIndex)
        summaryData(tenorIndex, 4) = _
            gAnnualizedReturnPct(tenorIndex)
        summaryData(tenorIndex, 5) = _
            gTenorMonthlyVolBps(tenorIndex)
        summaryData(tenorIndex, 6) = _
            gTenorSummary(tenorIndex, 9)
    Next tenorIndex

    ws.Range("A55:F59").Value = summaryData
    ws.Range("B55:C59").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("D55:D59").NumberFormat = _
        "0.0000\%;[Red](0.0000\%);-"
    ws.Range("E55:F59").NumberFormat = "0.00;[Red](0.00);-"

    RatesWriteRow ws.Range("H54"), _
        Array("Selected Constrained Portfolio", "Incremental Return (bps)", _
              "Earnings Volatility (bps)", "Downside Deviation (bps)", _
              "Reset Volatility (bps)", "WAM (Months)", _
              "Available <=30D", "Allocation")

    RatesStyleHeader ws.Range("H54:O54")

    For selectedNumber = 1 To 3
        selectedRow = SelectedFrontierRow( _
            selectedNumber, gConstrainedFrontierCount)

        For tenorIndex = 1 To TENOR_COUNT
            weights(tenorIndex) = _
                gConstrainedFrontier(selectedRow, tenorIndex)
        Next tenorIndex

        selectedData(selectedNumber, 1) = _
            SelectedFrontierLabel(selectedNumber)
        selectedData(selectedNumber, 2) = _
            gConstrainedFrontier(selectedRow, 7)
        selectedData(selectedNumber, 3) = _
            gConstrainedFrontier(selectedRow, 8)
        selectedData(selectedNumber, 4) = _
            gConstrainedFrontier(selectedRow, 9)
        selectedData(selectedNumber, 5) = _
            gConstrainedFrontier(selectedRow, 12)
        selectedData(selectedNumber, 6) = _
            gConstrainedFrontier(selectedRow, 14)
        selectedData(selectedNumber, 7) = _
            gConstrainedFrontier(selectedRow, 15) * 100#
        selectedData(selectedNumber, 8) = _
            AllocationText(weights)
    Next selectedNumber

    ws.Range("H55:O57").Value = selectedData
    ws.Range("I55:L57").NumberFormat = "0.00;[Red](0.00);-"
    ws.Range("M55:M57").NumberFormat = "0.0"
    ws.Range("N55:N57").NumberFormat = "0.0\%"
    ws.Range("O55:O57").WrapText = True

    ws.Range("A62:Q66").Interior.Color = COLOR_PALE
    ws.Range("A62").Value = "Model interpretation"
    ws.Range("A62").Font.Bold = True
    ws.Range("A63").Value = _
        "• Frontier portfolios are static sleeves. Monthly rebalancing is not assumed."
    ws.Range("A64").Value = _
        "• Constrained portfolios satisfy the current Frontier_Settings values."
    ws.Range("A65").Value = _
        "• Earnings volatility and common-sample reset volatility remain separate."
    ws.Range("A66").Value = _
        "• Out-of-sample results apply estimation-period allocations to later data."

    ws.Columns("A:Q").ColumnWidth = 11
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("O").ColumnWidth = 48
    ws.Rows("63:66").RowHeight = 21
End Sub

Private Sub BuildChartsSafe(ByVal dashboardSheet As Worksheet)
    On Error GoTo ChartFail

    DeleteAllCharts dashboardSheet
    gChartCount = 0

    AddHistoricalRatesChart dashboardSheet
    AddIncrementalInterestChart dashboardSheet
    AddAveragePremiumChart dashboardSheet
    AddFrontierChart dashboardSheet
    AddResetReturnChart dashboardSheet
    AddOutSampleChart dashboardSheet

    Exit Sub

ChartFail:
    gChartError = Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub AddHistoricalRatesChart(ByVal dashboardSheet As Worksheet)
    Dim dataSheet As Worksheet
    Dim finalRow As Long
    Dim chartObject As ChartObject

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")
    finalRow = dataSheet.Cells(dataSheet.Rows.Count, 1).End(xlUp).Row

    Set chartObject = CreateRangeChart( _
        dashboardSheet, dataSheet.Range("A3:F" & finalRow), _
        xlLine, "Historical Deposit Rates | Monthly Observations", _
        "A8", "I21", True)

    On Error Resume Next
    chartObject.Chart.Axes(xlValue).TickLabels.NumberFormat = "0.0""%"""
    chartObject.Chart.Axes(xlCategory).TickLabelSpacing = 6
    On Error GoTo 0

    gChartCount = gChartCount + 1
End Sub

Private Sub AddIncrementalInterestChart( _
    ByVal dashboardSheet As Worksheet)

    Dim dataSheet As Worksheet

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")

    CreateRangeChart dashboardSheet, dataSheet.Range("H3:I8"), _
                     xlBarClustered, _
                     "Final Rolling Interest vs ON ($000)", _
                     "J8", "Q21", False

    gChartCount = gChartCount + 1
End Sub

Private Sub AddAveragePremiumChart(ByVal dashboardSheet As Worksheet)
    Dim dataSheet As Worksheet

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")

    CreateRangeChart dashboardSheet, dataSheet.Range("K3:L8"), _
                     xlColumnClustered, _
                     "Average Quoted-Rate Premium vs ON | Common Starts", _
                     "A23", "I36", False

    gChartCount = gChartCount + 1
End Sub

Private Sub AddFrontierChart(ByVal dashboardSheet As Worksheet)
    Dim dataSheet As Worksheet
    Dim chartObject As ChartObject
    Dim chartValue As Chart
    Dim constrainedRow As Long
    Dim unconstrainedRow As Long
    Dim seriesValue As Series

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")
    constrainedRow = 3 + gConstrainedFrontierCount
    unconstrainedRow = 3 + gUnconstrainedFrontierCount

    Set chartObject = CreateEmptyChart( _
        dashboardSheet, xlXYScatterLinesNoMarkers, _
        "Static-Sleeve Efficient Frontier | Incremental Return vs ON", _
        "J23", "Q36")

    Set chartValue = chartObject.Chart

    Set seriesValue = chartValue.SeriesCollection.NewSeries
    seriesValue.Name = "Treasury Constrained"
    seriesValue.XValues = dataSheet.Range("N4:N" & constrainedRow)
    seriesValue.Values = dataSheet.Range("O4:O" & constrainedRow)

    Set seriesValue = chartValue.SeriesCollection.NewSeries
    seriesValue.Name = "Unconstrained"
    seriesValue.XValues = dataSheet.Range("Q4:Q" & unconstrainedRow)
    seriesValue.Values = dataSheet.Range("R4:R" & unconstrainedRow)

    chartValue.HasLegend = True
    chartValue.Legend.Position = xlLegendPositionBottom

    ApplyAxisTitles chartValue, _
                    "Annualized Earnings Volatility (bps)", _
                    "Incremental Annual Return vs ON (bps)"

    On Error Resume Next
    chartValue.Axes(xlCategory).TickLabels.NumberFormat = "0.00"
    chartValue.Axes(xlValue).TickLabels.NumberFormat = "0.00"
    On Error GoTo 0

    gChartCount = gChartCount + 1
End Sub

Private Sub AddResetReturnChart(ByVal dashboardSheet As Worksheet)
    Dim dataSheet As Worksheet
    Dim chartObject As ChartObject
    Dim chartValue As Chart
    Dim finalRow As Long
    Dim seriesValue As Series

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")
    finalRow = 3 + gConstrainedFrontierCount

    Set chartObject = CreateEmptyChart( _
        dashboardSheet, xlXYScatter, _
        "Constrained Frontier | Return vs Reinvestment Risk", _
        "A38", "I51")

    Set chartValue = chartObject.Chart

    Set seriesValue = chartValue.SeriesCollection.NewSeries
    seriesValue.Name = "Constrained Portfolios"
    seriesValue.XValues = dataSheet.Range("T4:T" & finalRow)
    seriesValue.Values = dataSheet.Range("U4:U" & finalRow)

    chartValue.HasLegend = False

    ApplyAxisTitles chartValue, _
                    "Portfolio Reset Volatility (bps)", _
                    "Incremental Annual Return vs ON (bps)"

    On Error Resume Next
    chartValue.Axes(xlCategory).TickLabels.NumberFormat = "0.00"
    chartValue.Axes(xlValue).TickLabels.NumberFormat = "0.00"
    On Error GoTo 0

    gChartCount = gChartCount + 1
End Sub

Private Sub AddOutSampleChart(ByVal dashboardSheet As Worksheet)
    Dim dataSheet As Worksheet

    Set dataSheet = ThisWorkbook.Worksheets("Chart_Data")

    CreateRangeChart dashboardSheet, dataSheet.Range("W3:Y6"), _
                     xlColumnClustered, _
                     "Estimation vs Out-of-Sample Incremental Return", _
                     "J38", "Q51", True

    gChartCount = gChartCount + 1
End Sub

Private Function CreateRangeChart( _
    ByVal dashboardSheet As Worksheet, _
    ByVal sourceRange As Range, _
    ByVal chartType As XlChartType, _
    ByVal chartTitle As String, _
    ByVal topLeftCell As String, _
    ByVal bottomRightCell As String, _
    ByVal showLegend As Boolean) As ChartObject

    Dim chartObject As ChartObject

    Set chartObject = CreateEmptyChart( _
        dashboardSheet, chartType, chartTitle, _
        topLeftCell, bottomRightCell)

    chartObject.Chart.SetSourceData _
        Source:=sourceRange, PlotBy:=xlColumns

    chartObject.Chart.HasLegend = showLegend

    If showLegend Then
        chartObject.Chart.Legend.Position = xlLegendPositionBottom
    End If

    Set CreateRangeChart = chartObject
End Function

Private Function CreateEmptyChart( _
    ByVal dashboardSheet As Worksheet, _
    ByVal chartType As XlChartType, _
    ByVal chartTitle As String, _
    ByVal topLeftCell As String, _
    ByVal bottomRightCell As String) As ChartObject

    Dim leftValue As Double
    Dim topValue As Double
    Dim widthValue As Double
    Dim heightValue As Double
    Dim chartObject As ChartObject

    leftValue = dashboardSheet.Range(topLeftCell).Left
    topValue = dashboardSheet.Range(topLeftCell).Top
    widthValue = dashboardSheet.Range(bottomRightCell).Left + _
                 dashboardSheet.Range(bottomRightCell).Width - leftValue
    heightValue = dashboardSheet.Range(bottomRightCell).Top + _
                  dashboardSheet.Range(bottomRightCell).Height - topValue

    Set chartObject = dashboardSheet.ChartObjects.Add( _
        leftValue, topValue, widthValue, heightValue)

    chartObject.Chart.ChartType = chartType
    chartObject.Chart.HasTitle = True
    chartObject.Chart.ChartTitle.Text = chartTitle

    Set CreateEmptyChart = chartObject
End Function

Private Sub ApplyAxisTitles(ByVal chartValue As Chart, _
                            ByVal horizontalTitle As String, _
                            ByVal verticalTitle As String)
    On Error Resume Next

    chartValue.Axes(xlCategory).HasTitle = True
    chartValue.Axes(xlCategory).AxisTitle.Text = horizontalTitle
    chartValue.Axes(xlValue).HasTitle = True
    chartValue.Axes(xlValue).AxisTitle.Text = verticalTitle

    On Error GoTo 0
End Sub

Private Sub FormatDashboardCard(ByVal ws As Worksheet, _
                                ByVal cardAddress As String, _
                                ByVal labelText As String, _
                                ByVal valueText As String)
    Dim firstCell As Range

    Set firstCell = ws.Range(cardAddress).Cells(1, 1)

    ws.Range(cardAddress).Interior.Color = COLOR_PALE
    ws.Range(cardAddress).Borders.LineStyle = xlContinuous
    ws.Range(cardAddress).Borders.Color = RGB(220, 225, 230)

    firstCell.Value = labelText
    firstCell.Font.Bold = True
    firstCell.Font.Color = RGB(110, 120, 130)

    firstCell.Offset(1, 0).Value = valueText
    firstCell.Offset(1, 0).Font.Bold = True
    firstCell.Offset(1, 0).Font.Size = 12
End Sub

' ============================================================================
' SECTION 10 - MODEL VALIDATION AND RECONCILIATION
'
' Purpose:
'   Reconcile daily rows, common samples, frontier generation, pure-tenor
'   static-sleeve returns, training/test observations, and dashboard charts.
' ============================================================================

Private Sub WriteValidationResults()
    Dim ws As Worksheet
    Dim rows As Collection
    Dim matrix As Variant
    Dim expectedDailyRows As Long
    Dim tenorIndex As Long
    Dim weights(1 To TENOR_COUNT) As Double
    Dim metrics(1 To 8) As Double
    Dim differenceValue As Double
    Dim benchmarkReturn As Double

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    SetOutputTitle ws, "Model Validation Results"

    RatesWriteRow ws.Range("A3"), _
        Array("Test", "Actual", "Expected / Tolerance", "Status")
    RatesStyleHeader ws.Range("A3:D3")

    Set rows = New Collection

    AddValidationRow rows, "Effective start date", gStartDate, _
                     gStartDate, "PASS"
    AddValidationRow rows, "Analysis end date", gEndDate, _
                     gEndDate, "PASS"

    expectedDailyRows = TENOR_COUNT * gNumDays

    AddValidationRow rows, "Daily ledger rows", gDailyRowCount, _
                     expectedDailyRows, _
                     PassFail(gDailyRowCount = expectedDailyRows)

    AddValidationRow rows, "Common daily starts per tenor", _
                     gCommonStartCount, "> 30", _
                     PassFail(gCommonStartCount > 30)

    AddValidationRow rows, "Constrained frontier points", _
                     gConstrainedFrontierCount, "> 0", _
                     PassFail(gConstrainedFrontierCount > 0)

    AddValidationRow rows, "Unconstrained frontier points", _
                     gUnconstrainedFrontierCount, "> 0", _
                     PassFail(gUnconstrainedFrontierCount > 0)

    AddValidationRow rows, "Training frontier points", _
                     gTrainingFrontierCount, "> 0", _
                     PassFail(gTrainingFrontierCount > 0)

    AddValidationRow rows, "Dashboard charts", gChartCount, 6, _
                     ChartStatus()

    benchmarkReturn = WindowTenorAnnualReturnPct( _
        1, 1, gMonthCount)

    For tenorIndex = 1 To TENOR_COUNT
        Erase weights
        weights(tenorIndex) = 1#

        CalculatePortfolioMetrics weights, 1, gMonthCount, _
                                  benchmarkReturn, metrics

        differenceValue = Abs( _
            metrics(1) - gTenorMonthlyReturnPct(tenorIndex))

        AddValidationRow rows, _
            TenorName(tenorIndex) & _
            " static-sleeve return reconciliation", _
            differenceValue, 0.000001, _
            PassFail(differenceValue <= 0.000001)
    Next tenorIndex

    AddValidationRow rows, "Validation training returns", _
                     gSplitMonthIndex - 1, ">= 12", _
                     PassFail(gSplitMonthIndex - 1 >= 12)

    AddValidationRow rows, "Validation test returns", _
                     gMonthCount - gSplitMonthIndex, ">= 6", _
                     PassFail(gMonthCount - gSplitMonthIndex >= 6)

    matrix = CollectionToMatrix(rows, 4)
    ws.Range("A4").Resize(rows.Count, 4).Value = matrix

    ws.Range("B5:B6").NumberFormat = "mm/dd/yyyy"
    ws.Columns("A").ColumnWidth = 46
    ws.Columns("B:C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 12
End Sub

Private Sub AddValidationRow(ByVal rows As Collection, _
                             ByVal testName As String, _
                             ByVal actualValue As Variant, _
                             ByVal expectedValue As Variant, _
                             ByVal statusText As String)
    rows.Add Array(testName, actualValue, expectedValue, statusText)
End Sub

Private Function ChartStatus() As String
    If gChartCount = 6 Then
        ChartStatus = "PASS"
    ElseIf Len(gChartError) > 0 Then
        ChartStatus = "WARN"
    Else
        ChartStatus = "FAIL"
    End If
End Function

Private Function PassFail(ByVal conditionValue As Boolean) As String
    If conditionValue Then
        PassFail = "PASS"
    Else
        PassFail = "FAIL"
    End If
End Function

Private Function ValidationPassed() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim statusText As String

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ValidationPassed = True

    For rowNumber = 4 To lastRow
        statusText = UCase$(Trim$(CStr(ws.Cells(rowNumber, 4).Value2)))

        If statusText = "FAIL" Then
            ValidationPassed = False
            Exit Function
        End If
    Next rowNumber
End Function

' ============================================================================
' SECTION 11 - SHARED CALCULATION AND WORKSHEET HELPERS
'
' Purpose:
'   Centralize date conventions, binary searches, statistics, sorting,
'   formatting, chart creation, and reusable utility functions.
' ============================================================================

Private Sub SetOutputTitle(ByVal ws As Worksheet, _
                           ByVal titleText As String)
    ws.Range("A1:AZ2").ClearFormats
    ws.Range("A1:AZ1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = titleText
    RatesStyleTitle ws.Range("A1:AZ1")
    ws.Rows(1).RowHeight = 28
End Sub

Private Sub SetFixedWidths(ByVal ws As Worksheet, _
                           ByVal widths As Variant)
    Dim columnIndex As Long

    For columnIndex = LBound(widths) To UBound(widths)
        ws.Columns(columnIndex + 1).ColumnWidth = _
            CDbl(widths(columnIndex))
    Next columnIndex
End Sub

Private Function SelectedFrontierRow(ByVal selectedNumber As Long, _
                                     ByVal frontierCount As Long) As Long
    Select Case selectedNumber
        Case 1
            SelectedFrontierRow = 1
        Case 2
            SelectedFrontierRow = _
                MaximumLong(1, (frontierCount + 1) \ 2)
        Case Else
            SelectedFrontierRow = frontierCount
    End Select
End Function

Private Function SelectedFrontierLabel( _
    ByVal selectedNumber As Long) As String

    Select Case selectedNumber
        Case 1
            SelectedFrontierLabel = "Minimum Volatility"
        Case 2
            SelectedFrontierLabel = "Balanced Frontier"
        Case Else
            SelectedFrontierLabel = "Maximum Historical Return"
    End Select
End Function

Private Function FrontierSegment(ByVal rankValue As Long, _
                                 ByVal totalCount As Long) As String
    Dim positionValue As Double

    If totalCount <= 1 Or rankValue = 1 Then
        FrontierSegment = "Minimum Volatility"
        Exit Function
    End If

    positionValue = (rankValue - 1) / (totalCount - 1)

    If positionValue <= 0.25 Then
        FrontierSegment = "Defensive"
    ElseIf positionValue <= 0.5 Then
        FrontierSegment = "Conservative"
    ElseIf positionValue <= 0.75 Then
        FrontierSegment = "Balanced"
    ElseIf positionValue < 1 Then
        FrontierSegment = "Return Oriented"
    Else
        FrontierSegment = "Maximum Historical Return"
    End If
End Function

Private Function FrontierDescription( _
    ByVal segmentText As String) As String

    Select Case segmentText
        Case "Minimum Volatility"
            FrontierDescription = _
                "Lowest historical static-sleeve earnings volatility."
        Case "Defensive"
            FrontierDescription = _
                "Small volatility increase for incremental historical return."
        Case "Conservative"
            FrontierDescription = _
                "Lower-half frontier portfolio with moderate maturity extension."
        Case "Balanced"
            FrontierDescription = _
                "Middle-frontier balance of return, volatility, reset risk, and liquidity."
        Case "Return Oriented"
            FrontierDescription = _
                "Higher historical return with greater earnings or reset risk."
        Case Else
            FrontierDescription = _
                "Highest in-sample historical return among efficient tested portfolios."
    End Select
End Function

Private Function AllocationText(ByRef weights() As Double) As String
    AllocationText = _
        Format$(weights(1), "0%") & " ON / " & _
        Format$(weights(2), "0%") & " 1M / " & _
        Format$(weights(3), "0%") & " 2M / " & _
        Format$(weights(4), "0%") & " 3M / " & _
        Format$(weights(5), "0%") & " 6M"
End Function

Private Function BalanceOnDate(ByVal tenorIndex As Long, _
                               ByVal balanceDate As Double) As Double
    Dim dayIndex As Long

    dayIndex = CLng(balanceDate - gStartDate)

    If dayIndex < 0 Or dayIndex > gNumDays - 1 Then
        Err.Raise vbObjectError + 700, , _
            "Requested balance date is outside the analysis period."
    End If

    BalanceOnDate = gBalance(tenorIndex, dayIndex)
End Function

Private Function FindCurveHeaderRow(ByVal ws As Worksheet) As Long
    Dim rowNumber As Long
    Dim foundDate As Boolean
    Dim foundON As Boolean

    For rowNumber = 1 To 20
        foundDate = (FindHeaderColumn(ws, rowNumber, "DATE") > 0)
        foundON = (FindHeaderColumn(ws, rowNumber, "ON") > 0)

        If foundDate And foundON Then
            FindCurveHeaderRow = rowNumber
            Exit Function
        End If
    Next rowNumber
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, _
                                  ByVal headerRow As Long, _
                                  ByVal headerText As String) As Long
    Dim columnNumber As Long
    Dim cellText As String

    For columnNumber = 1 To 30
        cellText = UCase$(Trim$(CStr( _
            ws.Cells(headerRow, columnNumber).Value2)))

        If cellText = UCase$(headerText) Then
            FindHeaderColumn = columnNumber
            Exit Function
        End If
    Next columnNumber
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

    If resultValue = 0 Then
        Err.Raise vbObjectError + 701, , _
            "No curve date exists on or before " & _
            Format$(targetDate, "dd-mmm-yyyy") & "."
    End If

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

Private Function FirstCurveDateOnOrAfter( _
    ByVal targetDate As Double) As Double

    Dim curveIndex As Long

    curveIndex = FindIndexOnOrAfter(targetDate)

    If curveIndex > 0 Then
        FirstCurveDateOnOrAfter = gCurveDates(curveIndex)
    End If
End Function

Private Function CurveDateOnOrBefore( _
    ByVal targetDate As Double) As Double

    CurveDateOnOrBefore = _
        gCurveDates(FindIndexOnOrBefore(targetDate))
End Function

Private Function NextCurveDateAfter( _
    ByVal currentDate As Double) As Double

    Dim curveIndex As Long

    curveIndex = FindIndexOnOrAfter(currentDate + 0.0000001)

    If curveIndex > 0 Then
        NextCurveDateAfter = gCurveDates(curveIndex)
    End If
End Function

Private Function AddMonthsPreserveEndOfMonth( _
    ByVal startDate As Double, _
    ByVal monthCount As Long) As Double

    Dim sourceDate As Date
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim sourceLastDay As Long
    Dim targetLastDay As Long
    Dim targetDay As Long

    sourceDate = CDate(startDate)
    targetYear = Year(DateAdd("m", monthCount, sourceDate))
    targetMonth = Month(DateAdd("m", monthCount, sourceDate))
    sourceLastDay = Day(DateSerial(Year(sourceDate), _
                                   Month(sourceDate) + 1, 0))
    targetLastDay = Day(DateSerial(targetYear, targetMonth + 1, 0))

    If Day(sourceDate) = sourceLastDay Then
        targetDay = targetLastDay
    Else
        targetDay = MinimumLong(Day(sourceDate), targetLastDay)
    End If

    AddMonthsPreserveEndOfMonth = _
        CDbl(DateSerial(targetYear, targetMonth, targetDay))
End Function

Private Function TenorName(ByVal tenorIndex As Long) As String
    Select Case tenorIndex
        Case 1
            TenorName = "ON"
        Case 2
            TenorName = "1M"
        Case 3
            TenorName = "2M"
        Case 4
            TenorName = "3M"
        Case 5
            TenorName = "6M"
        Case Else
            Err.Raise vbObjectError + 702, , "Invalid tenor index."
    End Select
End Function

Private Function TenorMonths(ByVal tenorIndex As Long) As Long
    Select Case tenorIndex
        Case 1
            TenorMonths = 0
        Case 2
            TenorMonths = 1
        Case 3
            TenorMonths = 2
        Case 4
            TenorMonths = 3
        Case 5
            TenorMonths = 6
        Case Else
            Err.Raise vbObjectError + 703, , "Invalid tenor index."
    End Select
End Function

Private Function BlankIfZero(ByVal numericValue As Double) As Variant
    If numericValue = 0 Then
        BlankIfZero = vbNullString
    Else
        BlankIfZero = numericValue
    End If
End Function

Private Function CollectionToMatrix(ByVal rows As Collection, _
                                    ByVal columnCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowData As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    ReDim matrix(1 To rows.Count, 1 To columnCount)

    For rowIndex = 1 To rows.Count
        rowData = rows(rowIndex)

        For columnIndex = 1 To columnCount
            matrix(rowIndex, columnIndex) = _
                rowData(columnIndex - 1)
        Next columnIndex
    Next rowIndex

    CollectionToMatrix = matrix
End Function

Private Function SampleStdDev(ByRef values() As Double) As Double
    Dim valueCount As Long
    Dim index As Long
    Dim averageValue As Double
    Dim varianceSum As Double

    valueCount = UBound(values) - LBound(values) + 1

    If valueCount < 2 Then Exit Function

    For index = LBound(values) To UBound(values)
        averageValue = averageValue + values(index)
    Next index

    averageValue = averageValue / valueCount

    For index = LBound(values) To UBound(values)
        varianceSum = varianceSum + _
            (values(index) - averageValue) ^ 2
    Next index

    SampleStdDev = Sqr(varianceSum / (valueCount - 1))
End Function

Private Function PercentileSorted(ByRef sortedValues() As Double, _
                                  ByVal percentileValue As Double) As Double
    Dim positionValue As Double
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim lowerBoundValue As Long

    lowerBoundValue = LBound(sortedValues)
    positionValue = (UBound(sortedValues) - lowerBoundValue) * _
                    percentileValue
    lowerIndex = lowerBoundValue + Int(positionValue)
    upperIndex = lowerBoundValue + _
                 Application.WorksheetFunction.RoundUp(positionValue, 0)

    If lowerIndex = upperIndex Then
        PercentileSorted = sortedValues(lowerIndex)
    Else
        PercentileSorted = sortedValues(lowerIndex) + _
            (positionValue - Int(positionValue)) * _
            (sortedValues(upperIndex) - sortedValues(lowerIndex))
    End If
End Function

Private Function PositiveShare(ByRef sortedValues() As Double) As Double
    Dim index As Long
    Dim positiveCount As Long
    Dim valueCount As Long

    valueCount = UBound(sortedValues) - LBound(sortedValues) + 1

    For index = LBound(sortedValues) To UBound(sortedValues)
        If sortedValues(index) > 0 Then
            positiveCount = positiveCount + 1
        End If
    Next index

    PositiveShare = positiveCount / valueCount
End Function

Private Sub SortDoubleArray(ByRef values() As Double)
    QuickSortDouble values, LBound(values), UBound(values)
End Sub

Private Sub QuickSortDouble(ByRef values() As Double, _
                            ByVal firstIndex As Long, _
                            ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivotValue As Double
    Dim temporaryValue As Double

    lowIndex = firstIndex
    highIndex = lastIndex
    pivotValue = values((firstIndex + lastIndex) \ 2)

    Do While lowIndex <= highIndex
        Do While values(lowIndex) < pivotValue
            lowIndex = lowIndex + 1
        Loop

        Do While values(highIndex) > pivotValue
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then
            temporaryValue = values(lowIndex)
            values(lowIndex) = values(highIndex)
            values(highIndex) = temporaryValue
            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then
        QuickSortDouble values, firstIndex, highIndex
    End If

    If lowIndex < lastIndex Then
        QuickSortDouble values, lowIndex, lastIndex
    End If
End Sub

Private Sub QuickSortCurve(ByVal firstIndex As Long, _
                           ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivotValue As Double
    Dim temporaryDate As Double
    Dim temporaryRate As Double
    Dim tenorIndex As Long

    lowIndex = firstIndex
    highIndex = lastIndex
    pivotValue = gCurveDates((firstIndex + lastIndex) \ 2)

    Do While lowIndex <= highIndex
        Do While gCurveDates(lowIndex) < pivotValue
            lowIndex = lowIndex + 1
        Loop

        Do While gCurveDates(highIndex) > pivotValue
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then
            temporaryDate = gCurveDates(lowIndex)
            gCurveDates(lowIndex) = gCurveDates(highIndex)
            gCurveDates(highIndex) = temporaryDate

            For tenorIndex = 1 To TENOR_COUNT
                temporaryRate = gRates(lowIndex, tenorIndex)
                gRates(lowIndex, tenorIndex) = _
                    gRates(highIndex, tenorIndex)
                gRates(highIndex, tenorIndex) = temporaryRate
            Next tenorIndex

            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then
        QuickSortCurve firstIndex, highIndex
    End If

    If lowIndex < lastIndex Then
        QuickSortCurve lowIndex, lastIndex
    End If
End Sub

Private Sub QuickSortCandidates(ByRef candidates() As Double, _
                                ByVal firstIndex As Long, _
                                ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivotVolatility As Double
    Dim pivotReturn As Double

    lowIndex = firstIndex
    highIndex = lastIndex
    pivotVolatility = candidates((firstIndex + lastIndex) \ 2, 8)
    pivotReturn = candidates((firstIndex + lastIndex) \ 2, 6)

    Do While lowIndex <= highIndex
        Do While CandidateComesBefore( _
            candidates, lowIndex, pivotVolatility, pivotReturn)
            lowIndex = lowIndex + 1
        Loop

        Do While CandidateComesAfter( _
            candidates, highIndex, pivotVolatility, pivotReturn)
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then
            SwapCandidateRows candidates, lowIndex, highIndex
            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then
        QuickSortCandidates candidates, firstIndex, highIndex
    End If

    If lowIndex < lastIndex Then
        QuickSortCandidates candidates, lowIndex, lastIndex
    End If
End Sub

Private Function CandidateComesBefore( _
    ByRef candidates() As Double, _
    ByVal rowIndex As Long, _
    ByVal pivotVolatility As Double, _
    ByVal pivotReturn As Double) As Boolean

    If candidates(rowIndex, 8) < pivotVolatility - 0.000000001 Then
        CandidateComesBefore = True
    ElseIf Abs(candidates(rowIndex, 8) - pivotVolatility) <= _
           0.000000001 Then

        CandidateComesBefore = _
            (candidates(rowIndex, 6) > pivotReturn)
    End If
End Function

Private Function CandidateComesAfter( _
    ByRef candidates() As Double, _
    ByVal rowIndex As Long, _
    ByVal pivotVolatility As Double, _
    ByVal pivotReturn As Double) As Boolean

    If candidates(rowIndex, 8) > pivotVolatility + 0.000000001 Then
        CandidateComesAfter = True
    ElseIf Abs(candidates(rowIndex, 8) - pivotVolatility) <= _
           0.000000001 Then

        CandidateComesAfter = _
            (candidates(rowIndex, 6) < pivotReturn)
    End If
End Function

Private Sub SwapCandidateRows(ByRef candidates() As Double, _
                              ByVal firstRow As Long, _
                              ByVal secondRow As Long)
    Dim columnIndex As Long
    Dim temporaryValue As Double

    If firstRow = secondRow Then Exit Sub

    For columnIndex = 1 To CANDIDATE_COLUMNS
        temporaryValue = candidates(firstRow, columnIndex)
        candidates(firstRow, columnIndex) = _
            candidates(secondRow, columnIndex)
        candidates(secondRow, columnIndex) = temporaryValue
    Next columnIndex
End Sub

Private Function CountWeightCombinations(ByVal nValue As Long, _
                                  ByVal kValue As Long) As Long
    Dim index As Long
    Dim resultValue As Double

    If kValue > nValue - kValue Then
        kValue = nValue - kValue
    End If

    resultValue = 1#

    For index = 1 To kValue
        resultValue = resultValue * _
                      (nValue - kValue + index) / index
    Next index

    CountWeightCombinations = CLng(Round(resultValue, 0))
End Function

Private Function MinimumDouble(ByVal firstValue As Double, _
                               ByVal secondValue As Double) As Double
    If firstValue < secondValue Then
        MinimumDouble = firstValue
    Else
        MinimumDouble = secondValue
    End If
End Function

Private Function MinimumLong(ByVal firstValue As Long, _
                             ByVal secondValue As Long) As Long
    If firstValue < secondValue Then
        MinimumLong = firstValue
    Else
        MinimumLong = secondValue
    End If
End Function

Private Function MaximumLong(ByVal firstValue As Long, _
                             ByVal secondValue As Long) As Long
    If firstValue > secondValue Then
        MaximumLong = firstValue
    Else
        MaximumLong = secondValue
    End If
End Function

Private Sub DeleteAllCharts(ByVal ws As Worksheet)
    Dim chartObject As ChartObject

    On Error Resume Next

    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject

    On Error GoTo 0
End Sub
