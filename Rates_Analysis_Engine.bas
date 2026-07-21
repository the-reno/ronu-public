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
' Corrected model design:
'   - Processes only Inputs!B5:B6.
'   - Rates are ordinary numbers: 4.31 means 4.31 percent.
'   - ACT/360 simple interest, reinvested at each actual maturity.
'   - Every new term deposit is dated from the prior actual maturity.
'   - Term maturity is the latest curve date on or before the target date.
'   - Tenor detail uses every eligible curve date as a separate start.
'   - Cross-tenor statistics use the same common start dates for all tenors.
'   - Efficient-frontier risk uses aligned monthly rolling-strategy returns.
'   - No cells are merged.
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

Private gStage As String
Private gCurveDates() As Double
Private gRates() As Double
Private gCurveCount As Long

Private gRequestedStart As Double
Private gStartDate As Double
Private gEndDate As Double
Private gNotional As Double
Private gWeightStep As Double
Private gWeightUnits As Long
Private gNumDays As Long

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

Private gMonthEndDates() As Double
Private gMonthlyReturns() As Double
Private gMonthlyCount As Long
Private gTenorMonthlyReturnPct() As Double
Private gTenorMonthlyVolBps() As Double

Private gPortfolio() As Double
Private gPortfolioCount As Long
Private gFrontier() As Double
Private gFrontierCount As Long

Public Sub RunRatesAnalysis()
    Dim oldCalculation As XlCalculation

    On Error GoTo Fail

    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SetStage "preparing workbook"
    EnsureRatesWorkbook
    ClearOutputs

    SetStage "loading curve"
    LoadCurve

    SetStage "loading and validating inputs"
    LoadInputs
    ValidateInputs
    InitializeArrays

    SetStage "building rolling strategies"
    BuildStrategies

    SetStage "writing transaction and daily ledgers"
    WriteTransactions
    WriteDailyAccrual
    WriteRollingResults

    SetStage "building daily tenor scenarios"
    BuildTenorScenarios
    WriteTenorAnalysis

    SetStage "building aligned monthly returns"
    BuildMonthlyReturns

    SetStage "building efficient frontier"
    BuildPortfolioFrontier

    SetStage "building chart data"
    BuildChartData

    SetStage "building dashboard charts"
    BuildDashboard

    SetStage "running validation"
    WriteValidationResults
    If Not ValidationPassed Then
        Err.Raise vbObjectError + 900, , _
            "One or more validation checks failed. Review Test_Results."
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
    MsgBox "Rates analysis stopped during " & gStage & "." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub ValidateRatesAnalysis()
    If Not SheetExists("Test_Results") Then
        MsgBox "Run RunRatesAnalysis first.", vbExclamation
    ElseIf ValidationPassed Then
        MsgBox "All current validation checks passed.", vbInformation
    Else
        MsgBox "At least one validation check failed. Review Test_Results.", vbExclamation
    End If
End Sub

Private Sub SetStage(ByVal stageText As String)
    gStage = stageText
    Application.StatusBar = "Rates model: " & stageText & "..."
    DoEvents
End Sub

' ============================================================================
' Inputs and curve
' ============================================================================

Private Sub LoadInputs()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Inputs")

    If Not IsDate(ws.Range("B5").Value) Then _
        Err.Raise vbObjectError + 101, , "Inputs!B5 must contain a start date."
    If Not IsDate(ws.Range("B6").Value) Then _
        Err.Raise vbObjectError + 102, , "Inputs!B6 must contain an end date."
    If Not IsNumeric(ws.Range("B7").Value) Then _
        Err.Raise vbObjectError + 103, , "Inputs!B7 must contain a notional."
    If Not IsNumeric(ws.Range("B8").Value) Then _
        Err.Raise vbObjectError + 104, , "Inputs!B8 must contain a frontier step."

    gRequestedStart = CDbl(CDate(ws.Range("B5").Value))
    gEndDate = CDbl(CDate(ws.Range("B6").Value))
    gNotional = CDbl(ws.Range("B7").Value)
    gWeightStep = CDbl(ws.Range("B8").Value)
End Sub

Private Sub LoadCurve()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dateColumn As Long
    Dim tenorColumn(1 To TENOR_COUNT) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim t As Long
    Dim rawDate As Variant
    Dim rawRate As Variant

    Set ws = ThisWorkbook.Worksheets("Curve")
    headerRow = FindCurveHeaderRow(ws)
    If headerRow = 0 Then
        Err.Raise vbObjectError + 110, , _
            "Curve headers not found. Required: Date, ON, 1M, 2M, 3M, 6M."
    End If

    dateColumn = FindHeaderColumn(ws, headerRow, "DATE")
    tenorColumn(1) = FindHeaderColumn(ws, headerRow, "ON")
    tenorColumn(2) = FindHeaderColumn(ws, headerRow, "1M")
    tenorColumn(3) = FindHeaderColumn(ws, headerRow, "2M")
    tenorColumn(4) = FindHeaderColumn(ws, headerRow, "3M")
    tenorColumn(5) = FindHeaderColumn(ws, headerRow, "6M")

    If dateColumn = 0 Then Err.Raise vbObjectError + 111, , "Date header is missing."
    For t = 1 To TENOR_COUNT
        If tenorColumn(t) = 0 Then
            Err.Raise vbObjectError + 112, , TenorName(t) & " header is missing."
        End If
    Next t

    lastRow = ws.Cells(ws.Rows.Count, dateColumn).End(xlUp).Row
    If lastRow <= headerRow Then _
        Err.Raise vbObjectError + 113, , "Curve contains no data rows."

    gCurveCount = 0
    For r = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, dateColumn).Value2))) > 0 Then _
            gCurveCount = gCurveCount + 1
    Next r

    If gCurveCount < 2 Then _
        Err.Raise vbObjectError + 114, , "Curve requires at least two dated rows."

    ReDim gCurveDates(1 To gCurveCount)
    ReDim gRates(1 To gCurveCount, 1 To TENOR_COUNT)

    i = 0
    For r = headerRow + 1 To lastRow
        rawDate = ws.Cells(r, dateColumn).Value
        If Len(Trim$(CStr(rawDate))) = 0 Then GoTo NextCurveRow

        If IsDate(rawDate) Then
            i = i + 1
            gCurveDates(i) = CDbl(CDate(rawDate))
        ElseIf IsNumeric(rawDate) And CDbl(rawDate) > 0 Then
            i = i + 1
            gCurveDates(i) = CDbl(rawDate)
        Else
            Err.Raise vbObjectError + 115, , "Invalid curve date on row " & r & "."
        End If

        For t = 1 To TENOR_COUNT
            rawRate = ws.Cells(r, tenorColumn(t)).Value2
            If Not IsNumeric(rawRate) Then
                Err.Raise vbObjectError + 116, , _
                    "Non-numeric " & TenorName(t) & " rate on row " & r & "."
            End If
            gRates(i, t) = CDbl(rawRate)
        Next t
NextCurveRow:
    Next r
End Sub

Private Sub ValidateInputs()
    Dim i As Long
    Dim t As Long
    Dim unitsExact As Double
    Dim rateAbsoluteSum As Double
    Dim rateObservationCount As Long

    If gRequestedStart > gEndDate Then _
        Err.Raise vbObjectError + 120, , "Start date is after end date."
    If gNotional <= 0 Then _
        Err.Raise vbObjectError + 121, , "Initial notional must be positive."

    For i = 2 To gCurveCount
        If gCurveDates(i) <= gCurveDates(i - 1) Then
            Err.Raise vbObjectError + 122, , _
                "Curve dates must be unique and strictly increasing. Sort the Curve sheet first."
        End If
    Next i

    If gRequestedStart > gCurveDates(gCurveCount) Then _
        Err.Raise vbObjectError + 123, , "Start date is after the final curve date."
    If gEndDate > gCurveDates(gCurveCount) Then _
        Err.Raise vbObjectError + 124, , "End date exceeds the final curve date."

    gStartDate = FirstCurveDateOnOrAfter(gRequestedStart)
    If gStartDate = 0 Or gStartDate > gEndDate Then _
        Err.Raise vbObjectError + 125, , "No curve date exists in the requested period."

    If gWeightStep < 5# Or gWeightStep > 50# Then
        Err.Raise vbObjectError + 126, , _
            "Frontier step must be between 5 and 50. Use 10 for ten-percent increments."
    End If

    unitsExact = 100# / gWeightStep
    gWeightUnits = CLng(Round(unitsExact, 0))
    If Abs(unitsExact - gWeightUnits) > 0.0000001 Then
        Err.Raise vbObjectError + 127, , _
            "Frontier step must divide 100 exactly."
    End If

    For i = 1 To gCurveCount
        For t = 1 To TENOR_COUNT
            If gRates(i, t) < -20# Or gRates(i, t) > 100# Then
                Err.Raise vbObjectError + 128, , _
                    "Rate outside the accepted numeric range on curve row " & i + 1 & "."
            End If
            rateAbsoluteSum = rateAbsoluteSum + Abs(gRates(i, t))
            rateObservationCount = rateObservationCount + 1
        Next t
    Next i

    If rateAbsoluteSum / rateObservationCount < 0.25 Then
        Err.Raise vbObjectError + 129, , _
            "Rates appear to be Excel percentages. Enter 4.31 for 4.31 percent."
    End If

    gNumDays = CLng(gEndDate - gStartDate) + 1
    If gNumDays < 90 Then
        Err.Raise vbObjectError + 130, , _
            "Select at least 90 calendar days for tenor and frontier analysis."
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
End Sub

' ============================================================================
' Rolling strategies
' ============================================================================

Private Sub BuildStrategies()
    Dim t As Long
    For t = 1 To TENOR_COUNT
        BuildOneStrategy t
        DoEvents
    Next t
End Sub

Private Sub BuildOneStrategy(ByVal tenorIndex As Long)
    Dim currentStart As Double
    Dim targetRoll As Double
    Dim actualRoll As Double
    Dim rateValue As Double
    Dim openingPrincipal As Double
    Dim closingPrincipal As Double
    Dim periodInterest As Double
    Dim dailyInterestValue As Double
    Dim transactionDays As Long
    Dim transactionID As Long
    Dim accrualEnd As Double
    Dim currentDay As Double
    Dim dayIndex As Long
    Dim daysAccrued As Long
    Dim completed As Boolean
    Dim statusText As String
    Dim adjustmentText As String
    Dim rollText As String
    Dim priorInterest As Double

    currentStart = gStartDate
    openingPrincipal = gNotional
    transactionID = 1

    Do While currentStart <= gEndDate
        rateValue = RateOnDate(currentStart, tenorIndex)

        If tenorIndex = 1 Then
            targetRoll = currentStart + 1
            actualRoll = NextCurveDateAfter(currentStart)
        Else
            targetRoll = AddMonthsPreserveEndOfMonth(currentStart, TenorMonths(tenorIndex))
            If targetRoll <= gCurveDates(gCurveCount) Then
                actualRoll = CurveDateOnOrBefore(targetRoll)
            Else
                actualRoll = 0
            End If
        End If

        If actualRoll > 0 And actualRoll <= currentStart Then
            Err.Raise vbObjectError + 200, , _
                "Invalid actual maturity for " & TenorName(tenorIndex) & "."
        End If

        completed = (actualRoll > 0 And actualRoll <= gEndDate)
        If actualRoll > 0 Then
            transactionDays = CLng(actualRoll - currentStart)
        Else
            transactionDays = CLng(targetRoll - currentStart)
        End If

        dailyInterestValue = openingPrincipal * (rateValue / 100#) / DAY_COUNT
        periodInterest = dailyInterestValue * transactionDays
        closingPrincipal = openingPrincipal + periodInterest

        adjustmentText = vbNullString
        If actualRoll > 0 And actualRoll <> targetRoll Then _
            adjustmentText = "Maturity moved to prior curve date"

        If completed Then
            statusText = "COMPLETED"
        Else
            statusText = "OPEN AT ANALYSIS END"
        End If

        AddTransaction tenorIndex, transactionID, currentStart, rateValue, targetRoll, _
                       actualRoll, transactionDays, openingPrincipal, periodInterest, _
                       closingPrincipal, statusText, adjustmentText

        If actualRoll > 0 Then
            accrualEnd = MinDouble(gEndDate, actualRoll - 1)
        Else
            accrualEnd = gEndDate
        End If

        currentDay = currentStart
        Do While currentDay <= accrualEnd
            dayIndex = CLng(currentDay - gStartDate)
            daysAccrued = CLng(currentDay - currentStart) + 1

            gOpeningPrincipal(tenorIndex, dayIndex) = openingPrincipal
            gDailyInterest(tenorIndex, dayIndex) = dailyInterestValue
            gBalance(tenorIndex, dayIndex) = openingPrincipal + _
                                               dailyInterestValue * daysAccrued
            gPrincipalDaySum(tenorIndex) = gPrincipalDaySum(tenorIndex) + openingPrincipal
            gInterestDaySum(tenorIndex) = gInterestDaySum(tenorIndex) + dailyInterestValue

            If currentDay = currentStart Then
                If transactionID = 1 Then
                    rollText = "START"
                Else
                    rollText = "ROLL / NEW DEAL"
                End If
            Else
                rollText = vbNullString
            End If

            AddDailyRow tenorIndex, transactionID, currentDay, currentStart, _
                        targetRoll, actualRoll, rateValue, openingPrincipal, _
                        dailyInterestValue, dailyInterestValue * daysAccrued, _
                        periodInterest, priorInterest, gBalance(tenorIndex, dayIndex), _
                        daysAccrued, transactionDays, rollText, statusText, adjustmentText

            currentDay = currentDay + 1
        Loop

        If Not completed Then Exit Do

        gCompletedTransactions(tenorIndex) = gCompletedTransactions(tenorIndex) + 1
        priorInterest = periodInterest
        openingPrincipal = closingPrincipal
        currentStart = actualRoll
        transactionID = transactionID + 1
    Loop

    gEndingValue(tenorIndex) = gBalance(tenorIndex, gNumDays - 1)
    gTotalInterest(tenorIndex) = gEndingValue(tenorIndex) - gNotional
    gAnnualizedReturnPct(tenorIndex) = _
        ((gEndingValue(tenorIndex) / gNotional) ^ (DAY_COUNT / gNumDays) - 1#) * 100#

    If gPrincipalDaySum(tenorIndex) > 0 Then
        gAverageRate(tenorIndex) = _
            gInterestDaySum(tenorIndex) * DAY_COUNT / gPrincipalDaySum(tenorIndex) * 100#
    End If
End Sub

Private Sub AddTransaction(ByVal tenorIndex As Long, _
                           ByVal transactionID As Long, _
                           ByVal startDate As Double, _
                           ByVal rateValue As Double, _
                           ByVal targetRoll As Double, _
                           ByVal actualRoll As Double, _
                           ByVal transactionDays As Long, _
                           ByVal openingPrincipal As Double, _
                           ByVal periodInterest As Double, _
                           ByVal closingPrincipal As Double, _
                           ByVal statusText As String, _
                           ByVal adjustmentText As String)
    Dim rowData As Variant

    rowData = Array(TenorName(tenorIndex), transactionID, startDate, startDate, _
                    startDate, rateValue, targetRoll, BlankIfZero(actualRoll), _
                    transactionDays, openingPrincipal, periodInterest, closingPrincipal, _
                    statusText, adjustmentText)
    gTransactionRows.Add rowData
End Sub

Private Sub AddDailyRow(ByVal tenorIndex As Long, _
                        ByVal transactionID As Long, _
                        ByVal accrualDate As Double, _
                        ByVal startDate As Double, _
                        ByVal targetRoll As Double, _
                        ByVal actualRoll As Double, _
                        ByVal rateValue As Double, _
                        ByVal openingPrincipal As Double, _
                        ByVal dailyInterestValue As Double, _
                        ByVal cumulativeInterest As Double, _
                        ByVal periodInterest As Double, _
                        ByVal priorInterest As Double, _
                        ByVal economicBalance As Double, _
                        ByVal daysAccrued As Long, _
                        ByVal transactionDays As Long, _
                        ByVal rollText As String, _
                        ByVal statusText As String, _
                        ByVal adjustmentText As String)
    Dim daysToRoll As Long

    gDailyRowCount = gDailyRowCount + 1
    If actualRoll > 0 Then
        daysToRoll = MaxLong(0, CLng(actualRoll - accrualDate))
    Else
        daysToRoll = MaxLong(0, CLng(targetRoll - accrualDate))
    End If

    gDailyRows(gDailyRowCount, 1) = accrualDate
    gDailyRows(gDailyRowCount, 2) = TenorName(tenorIndex)
    gDailyRows(gDailyRowCount, 3) = transactionID
    gDailyRows(gDailyRowCount, 4) = startDate
    gDailyRows(gDailyRowCount, 5) = targetRoll
    gDailyRows(gDailyRowCount, 6) = BlankIfZero(actualRoll)
    gDailyRows(gDailyRowCount, 7) = startDate
    gDailyRows(gDailyRowCount, 8) = rateValue
    gDailyRows(gDailyRowCount, 9) = openingPrincipal
    gDailyRows(gDailyRowCount, 10) = dailyInterestValue
    gDailyRows(gDailyRowCount, 11) = cumulativeInterest
    gDailyRows(gDailyRowCount, 12) = periodInterest
    If accrualDate = startDate Then
        gDailyRows(gDailyRowCount, 13) = priorInterest
    Else
        gDailyRows(gDailyRowCount, 13) = 0#
    End If
    gDailyRows(gDailyRowCount, 14) = economicBalance
    gDailyRows(gDailyRowCount, 15) = daysAccrued
    gDailyRows(gDailyRowCount, 16) = transactionDays
    gDailyRows(gDailyRowCount, 17) = daysToRoll
    gDailyRows(gDailyRowCount, 18) = rollText
    gDailyRows(gDailyRowCount, 19) = statusText
    gDailyRows(gDailyRowCount, 20) = adjustmentText
End Sub

' ============================================================================
' Daily tenor scenarios
' ============================================================================

Private Sub BuildTenorScenarios()
    Dim firstIndex As Long
    Dim finalIndex As Long
    Dim startIndex As Long
    Dim maturityIndex As Long
    Dim tenorIndex As Long
    Dim startDate As Double
    Dim targetDate As Double
    Dim actualDate As Double
    Dim startRate As Double
    Dim maturityRate As Double
    Dim actualDays As Long
    Dim horizonReturnPct As Double
    Dim horizonInterest As Double
    Dim resetChangeBps As Double
    Dim directionText As String
    Dim commonSample As String

    firstIndex = FindIndexOnOrAfter(gStartDate)
    finalIndex = FindIndexOnOrBefore(gEndDate)

    For tenorIndex = 1 To TENOR_COUNT
        For startIndex = firstIndex To finalIndex
            startDate = gCurveDates(startIndex)

            If tenorIndex = 1 Then
                actualDate = NextCurveDateAfter(startDate)
                targetDate = startDate + 1
                If actualDate = 0 Or actualDate > gEndDate Then Exit For
            Else
                targetDate = AddMonthsPreserveEndOfMonth(startDate, TenorMonths(tenorIndex))
                If targetDate > gEndDate Then Exit For
                maturityIndex = FindIndexOnOrBefore(targetDate)
                actualDate = gCurveDates(maturityIndex)
                If actualDate <= startDate Then GoTo NextScenario
            End If

            maturityIndex = FindIndexOnOrBefore(actualDate)
            startRate = gRates(startIndex, tenorIndex)
            maturityRate = gRates(maturityIndex, tenorIndex)
            actualDays = CLng(actualDate - startDate)
            horizonReturnPct = startRate * actualDays / DAY_COUNT
            horizonInterest = gNotional * horizonReturnPct / 100#
            resetChangeBps = (maturityRate - startRate) * 100#

            If resetChangeBps > 0 Then
                directionText = "Higher"
            ElseIf resetChangeBps < 0 Then
                directionText = "Lower"
            Else
                directionText = "Unchanged"
            End If

            If AddMonthsPreserveEndOfMonth(startDate, 6) <= gEndDate Then
                commonSample = "YES"
            Else
                commonSample = "NO"
            End If

            gScenarioRows.Add Array(TenorName(tenorIndex), startDate, startRate, _
                                    targetDate, actualDate, maturityRate, actualDays, _
                                    horizonReturnPct, horizonInterest, resetChangeBps, _
                                    directionText, commonSample)
NextScenario:
        Next startIndex
    Next tenorIndex

    CalculateTenorSummary
End Sub

Private Sub CalculateTenorSummary()
    Dim tenorIndex As Long
    Dim rowIndex As Long
    Dim rowData As Variant
    Dim allCount As Long
    Dim commonCount As Long
    Dim resetValues() As Double
    Dim sumReset As Double
    Dim sumRate As Double
    Dim sumDays As Double
    Dim sumReturn As Double
    Dim sumInterest As Double

    For tenorIndex = 1 To TENOR_COUNT
        allCount = 0
        commonCount = 0

        For rowIndex = 1 To gScenarioRows.Count
            rowData = gScenarioRows(rowIndex)
            If CStr(rowData(0)) = TenorName(tenorIndex) Then
                allCount = allCount + 1
                If CStr(rowData(11)) = "YES" Then commonCount = commonCount + 1
            End If
        Next rowIndex

        If commonCount < 2 Then
            Err.Raise vbObjectError + 300, , _
                "Insufficient common-start scenarios for " & TenorName(tenorIndex) & "."
        End If

        ReDim resetValues(1 To commonCount)
        commonCount = 0

        For rowIndex = 1 To gScenarioRows.Count
            rowData = gScenarioRows(rowIndex)
            If CStr(rowData(0)) = TenorName(tenorIndex) And _
               CStr(rowData(11)) = "YES" Then
                commonCount = commonCount + 1
                sumRate = sumRate + CDbl(rowData(2))
                sumDays = sumDays + CDbl(rowData(6))
                sumReturn = sumReturn + CDbl(rowData(7))
                sumInterest = sumInterest + CDbl(rowData(8))
                resetValues(commonCount) = CDbl(rowData(9))
                sumReset = sumReset + resetValues(commonCount)
            End If
        Next rowIndex

        SortDoubleArray resetValues
        gTenorSummary(tenorIndex, 1) = TenorName(tenorIndex)
        gTenorSummary(tenorIndex, 2) = allCount
        gTenorSummary(tenorIndex, 3) = commonCount
        gTenorSummary(tenorIndex, 4) = sumRate / commonCount
        gTenorSummary(tenorIndex, 5) = sumDays / commonCount
        gTenorSummary(tenorIndex, 6) = sumReturn / commonCount
        gTenorSummary(tenorIndex, 7) = sumInterest / commonCount
        gTenorSummary(tenorIndex, 8) = sumReset / commonCount
        gTenorSummary(tenorIndex, 9) = SampleStdDev(resetValues)
        gTenorSummary(tenorIndex, 10) = PercentileSorted(resetValues, 0.05)
        gTenorSummary(tenorIndex, 11) = PercentileSorted(resetValues, 0.5)
        gTenorSummary(tenorIndex, 12) = PercentileSorted(resetValues, 0.95)
        gTenorSummary(tenorIndex, 13) = resetValues(LBound(resetValues))
        gTenorSummary(tenorIndex, 14) = resetValues(UBound(resetValues))
        gTenorSummary(tenorIndex, 15) = PositiveShare(resetValues)

        sumRate = 0#
        sumDays = 0#
        sumReturn = 0#
        sumInterest = 0#
        sumReset = 0#
    Next tenorIndex
End Sub

' ============================================================================
' Monthly returns and frontier
' ============================================================================

Private Sub BuildMonthlyReturns()
    Dim ws As Worksheet
    Dim currentMonthEnd As Double
    Dim monthEndCount As Long
    Dim monthEnds() As Double
    Dim i As Long
    Dim t As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim outputData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 4) As Variant
    Dim values() As Double
    Dim compoundProduct As Double

    Set ws = ThisWorkbook.Worksheets("Monthly_Returns")
    ClearSheetBody ws

    currentMonthEnd = DateSerial(Year(CDate(gStartDate)), Month(CDate(gStartDate)) + 1, 0)
    Do While currentMonthEnd <= gEndDate
        monthEndCount = monthEndCount + 1
        If monthEndCount = 1 Then
            ReDim monthEnds(1 To 1)
        Else
            ReDim Preserve monthEnds(1 To monthEndCount)
        End If
        monthEnds(monthEndCount) = currentMonthEnd
        currentMonthEnd = DateSerial(Year(CDate(currentMonthEnd)), _
                                     Month(CDate(currentMonthEnd)) + 2, 0)
    Loop

    If monthEndCount < 3 Then
        Err.Raise vbObjectError + 400, , _
            "At least three complete month-end observations are required."
    End If

    gMonthlyCount = monthEndCount - 1
    ReDim gMonthEndDates(1 To monthEndCount)
    ReDim gMonthlyReturns(1 To gMonthlyCount, 1 To TENOR_COUNT)
    ReDim gTenorMonthlyReturnPct(1 To TENOR_COUNT)
    ReDim gTenorMonthlyVolBps(1 To TENOR_COUNT)
    ReDim outputData(1 To gMonthlyCount, 1 To 6)

    For i = 1 To monthEndCount
        gMonthEndDates(i) = monthEnds(i)
    Next i

    For i = 1 To gMonthlyCount
        outputData(i, 1) = monthEnds(i + 1)
        startIndex = CLng(monthEnds(i) - gStartDate)
        endIndex = CLng(monthEnds(i + 1) - gStartDate)

        For t = 1 To TENOR_COUNT
            gMonthlyReturns(i, t) = _
                gBalance(t, endIndex) / gBalance(t, startIndex) - 1#
            outputData(i, t + 1) = gMonthlyReturns(i, t) * 100#
        Next t
    Next i

    RatesWriteRow ws.Range("A3"), Array("Month End", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A3:F3")
    ws.Range("A4").Resize(gMonthlyCount, 6).Value2 = outputData
    ws.Range("A4:A" & gMonthlyCount + 3).NumberFormat = "mmm-yy"
    ws.Range("B4:F" & gMonthlyCount + 3).NumberFormat = "0.0000\%"

    RatesWriteRow ws.Range("H3"), _
        Array("Tenor", "Annualized Return", "Annualized Volatility (bps)", "WAM (Months)")
    RatesStyleHeader ws.Range("H3:K3")

    For t = 1 To TENOR_COUNT
        ReDim values(1 To gMonthlyCount)
        compoundProduct = 1#
        For i = 1 To gMonthlyCount
            values(i) = gMonthlyReturns(i, t)
            compoundProduct = compoundProduct * (1# + values(i))
        Next i

        gTenorMonthlyReturnPct(t) = _
            (compoundProduct ^ (12# / gMonthlyCount) - 1#) * 100#
        gTenorMonthlyVolBps(t) = SampleStdDev(values) * Sqr(12#) * 10000#

        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gTenorMonthlyReturnPct(t)
        summaryData(t, 3) = gTenorMonthlyVolBps(t)
        summaryData(t, 4) = TenorMonths(t)
    Next t

    ws.Range("H4:K8").Value2 = summaryData
    ws.Range("I4:I8").NumberFormat = "0.0000\%"
    ws.Range("J4:J8").NumberFormat = "0.00"
    ws.Range("K4:K8").NumberFormat = "0.0"
    SetStandardWidths ws, 11
End Sub

Private Sub BuildPortfolioFrontier()
    Dim ws As Worksheet
    Dim combinationCount As Long
    Dim w0 As Long
    Dim w1 As Long
    Dim w2 As Long
    Dim w3 As Long
    Dim w4 As Long
    Dim weight(1 To TENOR_COUNT) As Double
    Dim monthlyPortfolio() As Double
    Dim rowIndex As Long
    Dim i As Long
    Dim t As Long
    Dim compoundProduct As Double
    Dim annualReturn As Double
    Dim annualVol As Double
    Dim ratioValue As Double
    Dim selectedRows As Collection
    Dim selectedOutput() As Variant
    Dim frontierOutput() As Variant
    Dim rank As Long
    Dim p As Long
    Dim portfolioRow As Long
    Dim segmentText As String
    Dim selectedItem As Variant

    Set ws = ThisWorkbook.Worksheets("Portfolio_Analysis")
    ClearSheetBody ws

    combinationCount = CombinationCount(gWeightUnits + 4, 4)
    ReDim gPortfolio(1 To combinationCount, 1 To 13)
    ReDim monthlyPortfolio(1 To gMonthlyCount)

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

                    For i = 1 To gMonthlyCount
                        monthlyPortfolio(i) = 0#
                        For t = 1 To TENOR_COUNT
                            monthlyPortfolio(i) = monthlyPortfolio(i) + _
                                                  weight(t) * gMonthlyReturns(i, t)
                        Next t
                    Next i

                    compoundProduct = 1#
                    For i = 1 To gMonthlyCount
                        compoundProduct = compoundProduct * (1# + monthlyPortfolio(i))
                    Next i

                    annualReturn = _
                        (compoundProduct ^ (12# / gMonthlyCount) - 1#) * 100#
                    annualVol = SampleStdDev(monthlyPortfolio) * Sqr(12#) * 10000#
                    If annualVol > 0 Then
                        ratioValue = annualReturn / annualVol
                    Else
                        ratioValue = 0#
                    End If

                    rowIndex = rowIndex + 1
                    For t = 1 To TENOR_COUNT
                        gPortfolio(rowIndex, t) = weight(t)
                    Next t
                    gPortfolio(rowIndex, 6) = annualReturn
                    gPortfolio(rowIndex, 7) = annualVol
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

    RatesWriteRow ws.Range("A3"), _
        Array("Portfolio", "ON Weight", "1M Weight", "2M Weight", "3M Weight", _
              "6M Weight", "Annualized Return", "Annualized Volatility (bps)", _
              "WAM (Months)", "Description")
    RatesStyleHeader ws.Range("A3:J3")

    Set selectedRows = New Collection
    selectedRows.Add Array("100% ON", FindPortfolioRow(1#, 0#, 0#, 0#, 0#), _
                           "Immediate-liquidity reference portfolio.")
    selectedRows.Add Array("100% 1M", FindPortfolioRow(0#, 1#, 0#, 0#, 0#), _
                           "One-month reference portfolio.")
    selectedRows.Add Array("100% 2M", FindPortfolioRow(0#, 0#, 1#, 0#, 0#), _
                           "Two-month reference portfolio.")
    selectedRows.Add Array("100% 3M", FindPortfolioRow(0#, 0#, 0#, 1#, 0#), _
                           "Three-month reference portfolio.")
    selectedRows.Add Array("100% 6M", FindPortfolioRow(0#, 0#, 0#, 0#, 1#), _
                           "Six-month reference portfolio.")
    selectedRows.Add Array("Minimum Volatility", MinimumVolatilityRow(), _
                           "Lowest historical monthly earnings volatility.")
    selectedRows.Add Array("Maximum Historical Return", MaximumReturnRow(), _
                           "Highest historical annualized return among tested allocations.")

    ReDim selectedOutput(1 To selectedRows.Count, 1 To 10)
    For p = 1 To selectedRows.Count
        selectedItem = selectedRows(p)
        portfolioRow = CLng(selectedItem(1))
        selectedOutput(p, 1) = CStr(selectedItem(0))
        For t = 1 To TENOR_COUNT
            selectedOutput(p, t + 1) = gPortfolio(portfolioRow, t) * 100#
        Next t
        selectedOutput(p, 7) = gPortfolio(portfolioRow, 6)
        selectedOutput(p, 8) = gPortfolio(portfolioRow, 7)
        selectedOutput(p, 9) = gPortfolio(portfolioRow, 9)
        selectedOutput(p, 10) = CStr(selectedItem(2))
    Next p

    ws.Range("A4").Resize(selectedRows.Count, 10).Value2 = selectedOutput
    ws.Range("B4:F" & selectedRows.Count + 3).NumberFormat = "0.0\%"
    ws.Range("G4:G" & selectedRows.Count + 3).NumberFormat = "0.0000\%"
    ws.Range("H4:H" & selectedRows.Count + 3).NumberFormat = "0.00"
    ws.Range("I4:I" & selectedRows.Count + 3).NumberFormat = "0.0"
    ws.Range("J4:J" & selectedRows.Count + 3).WrapText = True

    RatesWriteRow ws.Range("A13"), _
        Array("Annualized Volatility (bps)", "Annualized Return", "ON Weight", _
              "1M Weight", "2M Weight", "3M Weight", "6M Weight", _
              "WAM (Months)", "Available <=30D", "Available <=60D", _
              "Available <=90D", "Frontier Rank", "Frontier Segment", _
              "Description", "Allocation Summary")
    RatesStyleHeader ws.Range("A13:O13")

    ReDim frontierOutput(1 To gFrontierCount, 1 To 15)
    For rank = 1 To gFrontierCount
        segmentText = FrontierSegment(rank, gFrontierCount)
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
        frontierOutput(rank, 14) = FrontierDescription(segmentText)
        frontierOutput(rank, 15) = AllocationDescription(rank)
    Next rank

    ws.Range("A14").Resize(gFrontierCount, 15).Value2 = frontierOutput
    ws.Range("A14:A" & gFrontierCount + 13).NumberFormat = "0.00"
    ws.Range("B14:B" & gFrontierCount + 13).NumberFormat = "0.0000\%"
    ws.Range("C14:G" & gFrontierCount + 13).NumberFormat = "0.0\%"
    ws.Range("H14:H" & gFrontierCount + 13).NumberFormat = "0.0"
    ws.Range("I14:K" & gFrontierCount + 13).NumberFormat = "0.0\%"
    ws.Range("N14:O" & gFrontierCount + 13).WrapText = True
    SetStandardWidths ws, 15
    ws.Columns("J").ColumnWidth = 42
    ws.Columns("N:O").ColumnWidth = 48
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

' ============================================================================
' Output sheets
' ============================================================================

Private Sub WriteTransactions()
    Dim ws As Worksheet
    Dim data As Variant
    Dim countValue As Long

    Set ws = ThisWorkbook.Worksheets("Transactions")
    ClearSheetBody ws
    RatesWriteRow ws.Range("A3"), _
        Array("Tenor", "Transaction ID", "Target Start Date", "Actual Start Date", _
              "Rate Observation Date", "Rate Used", "Target Roll Date", _
              "Actual Roll Date", "Transaction Days", "Opening Notional ($)", _
              "Period Interest ($)", "Closing Notional ($)", "Status", "Adjustment Flag")
    RatesStyleHeader ws.Range("A3:N3")

    countValue = gTransactionRows.Count
    If countValue = 0 Then Exit Sub
    data = CollectionToMatrix(gTransactionRows, 14)
    ws.Range("A4").Resize(countValue, 14).Value2 = data
    ws.Range("C4:E" & countValue + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("G4:H" & countValue + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("F4:F" & countValue + 3).NumberFormat = "0.0000"
    ws.Range("J4:L" & countValue + 3).NumberFormat = "$#,##0;[Red]($#,##0);-"
    SetStandardWidths ws, 14
    ws.Columns("N").ColumnWidth = 38
End Sub

Private Sub WriteDailyAccrual()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Daily_Accrual")
    ClearSheetBody ws
    RatesWriteRow ws.Range("A3"), _
        Array("Accrual Date", "Tenor", "Transaction ID", "Transaction Start Date", _
              "Target Roll Date", "Actual Roll Date", "Rate Observation Date", _
              "Rate Used", "Opening Notional ($)", "Daily Interest ($)", _
              "Cumulative Period Interest ($)", "Full Period Interest ($)", _
              "Interest Paid Today ($)", "Economic Balance ($)", "Days Accrued", _
              "Transaction Days", "Days to Roll", "Roll Flag", "Status", "Adjustment Flag")
    RatesStyleHeader ws.Range("A3:T3")

    If gDailyRowCount = 0 Then Exit Sub
    ws.Range("A4").Resize(gDailyRowCount, 20).Value2 = gDailyRows
    ws.Range("A4:A" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("D4:G" & gDailyRowCount + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("H4:H" & gDailyRowCount + 3).NumberFormat = "0.0000"
    ws.Range("I4:N" & gDailyRowCount + 3).NumberFormat = "$#,##0;[Red]($#,##0);-"
    SetStandardWidths ws, 20
    ws.Columns("T").ColumnWidth = 38
End Sub

Private Sub WriteRollingResults()
    Dim ws As Worksheet
    Dim growthData() As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 8) As Variant
    Dim d As Long
    Dim t As Long

    Set ws = ThisWorkbook.Worksheets("Rolling_Results")
    ClearSheetBody ws
    RatesWriteRow ws.Range("A3"), Array("Date", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A3:F3")

    ReDim growthData(1 To gNumDays, 1 To 6)
    For d = 0 To gNumDays - 1
        growthData(d + 1, 1) = gStartDate + d
        For t = 1 To TENOR_COUNT
            growthData(d + 1, t + 1) = gBalance(t, d)
        Next t
    Next d
    ws.Range("A4").Resize(gNumDays, 6).Value2 = growthData
    ws.Range("A4:A" & gNumDays + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("B4:F" & gNumDays + 3).NumberFormat = "$#,##0;[Red]($#,##0);-"

    RatesWriteRow ws.Range("H3"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", "Total Return", _
              "Annualized Return", "Average Invested Rate", _
              "Completed Transactions", "Incremental Interest vs ON ($)")
    RatesStyleHeader ws.Range("H3:O3")

    For t = 1 To TENOR_COUNT
        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gEndingValue(t)
        summaryData(t, 3) = gTotalInterest(t)
        summaryData(t, 4) = (gEndingValue(t) / gNotional - 1#) * 100#
        summaryData(t, 5) = gAnnualizedReturnPct(t)
        summaryData(t, 6) = gAverageRate(t)
        summaryData(t, 7) = gCompletedTransactions(t)
        summaryData(t, 8) = gTotalInterest(t) - gTotalInterest(1)
    Next t

    ws.Range("H4:O8").Value2 = summaryData
    ws.Range("I4:J8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("K4:M8").NumberFormat = "0.0000\%"
    ws.Range("N4:N8").NumberFormat = "0"
    ws.Range("O4:O8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    SetStandardWidths ws, 15
End Sub

Private Sub WriteTenorAnalysis()
    Dim ws As Worksheet
    Dim detailData As Variant
    Dim summaryData(1 To TENOR_COUNT, 1 To 15) As Variant
    Dim t As Long
    Dim c As Long

    Set ws = ThisWorkbook.Worksheets("Tenor_Analysis")
    ClearSheetBody ws
    RatesWriteRow ws.Range("A3"), _
        Array("Tenor", "Start Date", "Start Rate", "Target Maturity", _
              "Actual Maturity", "Maturity Rate", "Actual Days", _
              "Horizon Return", "Horizon Interest ($)", "Reset Change (bps)", _
              "Direction", "Common Comparison Sample")
    RatesStyleHeader ws.Range("A3:L3")

    If gScenarioRows.Count > 0 Then
        detailData = CollectionToMatrix(gScenarioRows, 12)
        ws.Range("A4").Resize(gScenarioRows.Count, 12).Value2 = detailData
        ws.Range("B4:B" & gScenarioRows.Count + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("D4:E" & gScenarioRows.Count + 3).NumberFormat = "mm/dd/yyyy"
        ws.Range("C4:C" & gScenarioRows.Count + 3).NumberFormat = "0.0000"
        ws.Range("F4:F" & gScenarioRows.Count + 3).NumberFormat = "0.0000"
        ws.Range("H4:H" & gScenarioRows.Count + 3).NumberFormat = "0.0000\%"
        ws.Range("I4:I" & gScenarioRows.Count + 3).NumberFormat = _
            "$#,##0;[Red]($#,##0);-"
        ws.Range("J4:J" & gScenarioRows.Count + 3).NumberFormat = _
            "0.0;[Red](0.0);-"
    End If

    RatesWriteRow ws.Range("N3"), _
        Array("Tenor", "All Eligible Scenarios", "Common Start Scenarios", _
              "Average Quoted Rate", "Average Horizon Days", _
              "Average Horizon Return", "Average Horizon Interest ($)", _
              "Average Reset (bps)", "Reset Volatility (bps)", _
              "5th Percentile (bps)", "Median Reset (bps)", _
              "95th Percentile (bps)", "Worst Reset (bps)", _
              "Best Reset (bps)", "Positive Resets")
    RatesStyleHeader ws.Range("N3:AB3")

    For t = 1 To TENOR_COUNT
        For c = 1 To 15
            summaryData(t, c) = gTenorSummary(t, c)
        Next c
    Next t
    ws.Range("N4:AB8").Value2 = summaryData
    ws.Range("P4:Q8").NumberFormat = "0"
    ws.Range("R4:R8").NumberFormat = "0.0000"
    ws.Range("S4:S8").NumberFormat = "0.0"
    ws.Range("T4:T8").NumberFormat = "0.0000\%"
    ws.Range("U4:U8").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("V4:AA8").NumberFormat = "0.0;[Red](0.0);-"
    ws.Range("AB4:AB8").NumberFormat = "0.0%"
    SetStandardWidths ws, 28
    ws.Columns("U").ColumnWidth = 20
End Sub

' ============================================================================
' Chart data and dashboard
' ============================================================================

Private Sub BuildChartData()
    Dim ws As Worksheet
    Dim firstCurveIndex As Long
    Dim finalCurveIndex As Long
    Dim currentIndex As Long
    Dim lastMonthIndex As Long
    Dim previousMonth As Long
    Dim previousYear As Long
    Dim outputRow As Long
    Dim t As Long
    Dim i As Long
    Dim incremental(1 To TENOR_COUNT, 1 To 2) As Variant
    Dim averageRate(1 To TENOR_COUNT, 1 To 2) As Variant
    Dim resetVol(1 To TENOR_COUNT, 1 To 2) As Variant
    Dim resetDistribution(1 To TENOR_COUNT, 1 To 4) As Variant

    Set ws = ThisWorkbook.Worksheets("Chart_Data")
    ws.Cells.Clear

    RatesWriteRow ws.Range("A1"), Array("Date", "ON", "1M", "2M", "3M", "6M")
    RatesStyleHeader ws.Range("A1:F1")

    firstCurveIndex = FindIndexOnOrAfter(gStartDate)
    finalCurveIndex = FindIndexOnOrBefore(gEndDate)
    previousMonth = -1
    outputRow = 2

    For currentIndex = firstCurveIndex To finalCurveIndex
        If previousMonth = -1 Then
            previousMonth = Month(CDate(gCurveDates(currentIndex)))
            previousYear = Year(CDate(gCurveDates(currentIndex)))
            lastMonthIndex = currentIndex
        ElseIf Month(CDate(gCurveDates(currentIndex))) <> previousMonth Or _
               Year(CDate(gCurveDates(currentIndex))) <> previousYear Then
            ws.Cells(outputRow, 1).Value2 = Format$(gCurveDates(lastMonthIndex), "mmm-yy")
            For t = 1 To TENOR_COUNT
                ws.Cells(outputRow, t + 1).Value2 = gRates(lastMonthIndex, t)
            Next t
            outputRow = outputRow + 1
            previousMonth = Month(CDate(gCurveDates(currentIndex)))
            previousYear = Year(CDate(gCurveDates(currentIndex)))
            lastMonthIndex = currentIndex
        Else
            lastMonthIndex = currentIndex
        End If
    Next currentIndex

    ws.Cells(outputRow, 1).Value2 = Format$(gCurveDates(lastMonthIndex), "mmm-yy")
    For t = 1 To TENOR_COUNT
        ws.Cells(outputRow, t + 1).Value2 = gRates(lastMonthIndex, t)
    Next t
    ws.Range("A2:A" & outputRow).NumberFormat = "@"
    ws.Range("B2:F" & outputRow).NumberFormat = "0.0000"

    RatesWriteRow ws.Range("H1"), Array("Tenor", "Incremental Interest vs ON ($000)")
    RatesStyleHeader ws.Range("H1:I1")
    RatesWriteRow ws.Range("K1"), Array("Tenor", "Average Rate Premium vs ON (bps)")
    RatesStyleHeader ws.Range("K1:L1")
    RatesWriteRow ws.Range("N1"), Array("Tenor", "Reset Volatility (bps)")
    RatesStyleHeader ws.Range("N1:O1")
    RatesWriteRow ws.Range("Q1"), _
        Array("Tenor", "5th Percentile", "Median", "95th Percentile")
    RatesStyleHeader ws.Range("Q1:T1")

    For t = 1 To TENOR_COUNT
        incremental(t, 1) = TenorName(t)
        incremental(t, 2) = (gTotalInterest(t) - gTotalInterest(1)) / 1000#
        averageRate(t, 1) = TenorName(t)
        averageRate(t, 2) = (CDbl(gTenorSummary(t, 4)) - CDbl(gTenorSummary(1, 4))) * 100#
        resetVol(t, 1) = TenorName(t)
        resetVol(t, 2) = gTenorSummary(t, 9)
        resetDistribution(t, 1) = TenorName(t)
        resetDistribution(t, 2) = gTenorSummary(t, 10)
        resetDistribution(t, 3) = gTenorSummary(t, 11)
        resetDistribution(t, 4) = gTenorSummary(t, 12)
    Next t

    ws.Range("H2:I6").Value2 = incremental
    ws.Range("K2:L6").Value2 = averageRate
    ws.Range("N2:O6").Value2 = resetVol
    ws.Range("Q2:T6").Value2 = resetDistribution

    RatesWriteRow ws.Range("V1"), _
        Array("Annualized Volatility (bps)", "Annualized Return")
    RatesStyleHeader ws.Range("V1:W1")
    For i = 1 To gFrontierCount
        ws.Cells(i + 1, 22).Value2 = gFrontier(i, 7)
        ws.Cells(i + 1, 23).Value2 = gFrontier(i, 6)
    Next i

    RatesWriteRow ws.Range("Y1"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", _
              "Annualized Return", "Monthly Volatility (bps)", _
              "Reset Volatility (bps)")
    RatesStyleHeader ws.Range("Y1:AD1")
    For t = 1 To TENOR_COUNT
        ws.Cells(t + 1, 25).Value2 = TenorName(t)
        ws.Cells(t + 1, 26).Value2 = gEndingValue(t)
        ws.Cells(t + 1, 27).Value2 = gTotalInterest(t)
        ws.Cells(t + 1, 28).Value2 = gAnnualizedReturnPct(t)
        ws.Cells(t + 1, 29).Value2 = gTenorMonthlyVolBps(t)
        ws.Cells(t + 1, 30).Value2 = gTenorSummary(t, 9)
    Next t

    SetStandardWidths ws, 30
End Sub

Private Sub BuildDashboard()
    Dim ws As Worksheet
    Dim chartData As Worksheet
    Dim lastRateRow As Long
    Dim bestTenor As Long
    Dim t As Long
    Dim summaryData(1 To TENOR_COUNT, 1 To 6) As Variant

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set chartData = ThisWorkbook.Worksheets("Chart_Data")
    DeleteAllCharts ws
    ws.Cells.Clear

    ws.Range("A1:Q1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Historical Cash Investment Analysis | Corrected Tenor Model"
    ws.Range("A1:Q1").Font.Bold = True
    ws.Range("A1:Q1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:Q1").Font.Size = 17
    ws.Range("A2:Q2").Interior.Color = COLOR_NAVY
    ws.Range("A2").Value = _
        "Input-date rolling returns, all-date tenor scenarios and efficient frontier"
    ws.Range("A2:Q2").Font.Color = RGB(255, 255, 255)
    ws.Range("A2:Q2").Font.Italic = True

    bestTenor = 1
    For t = 2 To TENOR_COUNT
        If gEndingValue(t) > gEndingValue(bestTenor) Then bestTenor = t
    Next t

    FormatCard ws.Range("A4:D6"), "Analysis period", _
        Format$(gStartDate, "dd-mmm-yyyy") & " to " & Format$(gEndDate, "dd-mmm-yyyy")
    FormatCard ws.Range("E4:H6"), "Initial cash", _
        "$" & Format$(gNotional / 1000000#, "0.0") & "MM"
    FormatCard ws.Range("J4:M6"), "Highest rolling ending value", _
        TenorName(bestTenor) & " | $" & _
        Format$(gEndingValue(bestTenor) / 1000000#, "0.000") & "MM"
    FormatCard ws.Range("N4:Q6"), "Term analysis", _
        Format$(gScenarioRows.Count, "#,##0") & " total; " & _
        Format$(gTenorSummary(1, 3), "#,##0") & " common starts per tenor"

    lastRateRow = chartData.Cells(chartData.Rows.Count, 1).End(xlUp).Row
    CreateLineChart ws, chartData.Range("A1:F" & lastRateRow), _
        "Historical deposit rates | monthly observations", "A8", "I21"
    CreateBarChart ws, chartData.Range("H1:I6"), _
        "Incremental rolling interest versus ON ($000)", "J8", "Q21", True
    CreateColumnChart ws, chartData.Range("K1:L6"), _
        "Average quoted-rate premium vs ON | all daily starts", "A23", "I36"
    CreateColumnChart ws, chartData.Range("N1:O6"), _
        "Reset volatility | full tenor horizon", "J23", "Q36"
    CreateColumnChart ws, chartData.Range("Q1:T6"), _
        "Reset distribution | 5th, median and 95th percentile", "A38", "I51"
    CreateFrontierChart ws, chartData, "J38", "Q51"

    RatesWriteRow ws.Range("A54"), _
        Array("Tenor", "Ending Value ($)", "Total Interest ($)", _
              "Annualized Rolling Return", "Monthly Earnings Volatility (bps)", _
              "Reset Volatility (bps)")
    RatesStyleHeader ws.Range("A54:F54")

    For t = 1 To TENOR_COUNT
        summaryData(t, 1) = TenorName(t)
        summaryData(t, 2) = gEndingValue(t)
        summaryData(t, 3) = gTotalInterest(t)
        summaryData(t, 4) = gAnnualizedReturnPct(t)
        summaryData(t, 5) = gTenorMonthlyVolBps(t)
        summaryData(t, 6) = gTenorSummary(t, 9)
    Next t
    ws.Range("A55:F59").Value2 = summaryData
    ws.Range("B55:C59").NumberFormat = "$#,##0;[Red]($#,##0);-"
    ws.Range("D55:D59").NumberFormat = "0.0000\%"
    ws.Range("E55:F59").NumberFormat = "0.00"

    ws.Range("H54:Q54").Interior.Color = COLOR_NAVY
    ws.Range("H54").Value = "Interpretation"
    ws.Range("H54:Q54").Font.Bold = True
    ws.Range("H54:Q54").Font.Color = RGB(255, 255, 255)
    ws.Range("H55:Q60").Interior.Color = COLOR_PALE
    ws.Range("H55").Value = _
        "Rolling results follow one reinvestment path from the effective start date. " & _
        "Tenor analysis is different: every eligible curve date is a separate start, " & _
        "and the same tenor is compared at actual maturity. Detail retains every eligible " & _
        "start; comparison statistics use the same common start dates for every tenor. " & _
        "Reset volatility measures full-horizon reinvestment-rate uncertainty. The frontier " & _
        "uses aligned monthly rolling-strategy returns and does not substitute reset volatility."
    ws.Range("H55:Q60").WrapText = True

    SetDashboardWidths ws
End Sub

' ============================================================================
' Validation
' ============================================================================

Private Sub WriteValidationResults()
    Dim ws As Worksheet
    Dim rows As Collection
    Dim matrix As Variant
    Dim t As Long
    Dim differenceValue As Double

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    ClearSheetBody ws
    RatesWriteRow ws.Range("A3"), Array("Test", "Actual", "Expected / Tolerance", "Status")
    RatesStyleHeader ws.Range("A3:D3")
    Set rows = New Collection

    rows.Add TestRow("Effective start respects input", gStartDate, _
                     ">= requested start", IIf(gStartDate >= gRequestedStart, "PASS", "FAIL"))
    rows.Add TestRow("Analysis end", gEndDate, "Input end date", "PASS")
    rows.Add TestRow("Daily ledger rows", gDailyRowCount, TENOR_COUNT * gNumDays, _
                     IIf(gDailyRowCount = TENOR_COUNT * gNumDays, "PASS", "FAIL"))
    rows.Add TestRow("Daily tenor scenarios", gScenarioRows.Count, "> 0", _
                     IIf(gScenarioRows.Count > 0, "PASS", "FAIL"))
    rows.Add TestRow("Efficient frontier points", gFrontierCount, "> 0", _
                     IIf(gFrontierCount > 0, "PASS", "FAIL"))
    rows.Add TestRow("Dashboard charts", _
                     ThisWorkbook.Worksheets("Dashboard").ChartObjects.Count, 6, _
                     IIf(ThisWorkbook.Worksheets("Dashboard").ChartObjects.Count = 6, _
                         "PASS", "FAIL"))

    For t = 1 To TENOR_COUNT
        differenceValue = SumDailyInterest(t) - gTotalInterest(t)
        rows.Add TestRow(TenorName(t) & " interest reconciliation", _
                         Abs(differenceValue), 0.05, _
                         IIf(Abs(differenceValue) <= 0.05, "PASS", "FAIL"))
        rows.Add TestRow(TenorName(t) & " scenario count", _
                         CLng(gTenorSummary(t, 2)), "> 1", _
                         IIf(CLng(gTenorSummary(t, 2)) > 1, "PASS", "FAIL"))
        rows.Add TestRow(TenorName(t) & " common-start count", _
                         CLng(gTenorSummary(t, 3)), CLng(gTenorSummary(1, 3)), _
                         IIf(CLng(gTenorSummary(t, 3)) = CLng(gTenorSummary(1, 3)), _
                             "PASS", "FAIL"))
    Next t

    matrix = CollectionToMatrix(rows, 4)
    ws.Range("A4").Resize(rows.Count, 4).Value2 = matrix
    ws.Range("B4:B5").NumberFormat = "mm/dd/yyyy"
    ws.Columns("A").ColumnWidth = 42
    ws.Columns("B:D").ColumnWidth = 20
End Sub

Private Function ValidationPassed() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("Test_Results")
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    ValidationPassed = True
    For r = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 4).Value2))) <> "PASS" Then
            ValidationPassed = False
            Exit Function
        End If
    Next r
End Function

Private Function TestRow(ByVal testName As String, _
                         ByVal actualValue As Variant, _
                         ByVal expectedValue As Variant, _
                         ByVal statusText As String) As Variant
    TestRow = Array(testName, actualValue, expectedValue, statusText)
End Function

' ============================================================================
' Chart helpers
' ============================================================================

Private Sub CreateLineChart(ByVal dashboard As Worksheet, _
                            ByVal sourceRange As Range, _
                            ByVal titleText As String, _
                            ByVal topLeft As String, _
                            ByVal bottomRight As String)
    Dim chartObject As ChartObject

    Set chartObject = AddChartBox(dashboard, topLeft, bottomRight)
    With chartObject.Chart
        .ChartType = xlLine
        .SetSourceData Source:=sourceRange
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        On Error Resume Next
        .Axes(xlCategory).TickLabelSpacing = 6
        On Error GoTo 0
    End With
    SafeFormatAxes chartObject.Chart, "0.0", "@"
End Sub

Private Sub CreateColumnChart(ByVal dashboard As Worksheet, _
                              ByVal sourceRange As Range, _
                              ByVal titleText As String, _
                              ByVal topLeft As String, _
                              ByVal bottomRight As String)
    Dim chartObject As ChartObject

    Set chartObject = AddChartBox(dashboard, topLeft, bottomRight)
    With chartObject.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=sourceRange
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = (sourceRange.Columns.Count > 2)
        If .HasLegend Then .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Private Sub CreateBarChart(ByVal dashboard As Worksheet, _
                           ByVal sourceRange As Range, _
                           ByVal titleText As String, _
                           ByVal topLeft As String, _
                           ByVal bottomRight As String, _
                           ByVal horizontalBars As Boolean)
    Dim chartObject As ChartObject

    Set chartObject = AddChartBox(dashboard, topLeft, bottomRight)
    With chartObject.Chart
        If horizontalBars Then
            .ChartType = xlBarClustered
        Else
            .ChartType = xlColumnClustered
        End If
        .SetSourceData Source:=sourceRange
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = False
    End With
End Sub

Private Sub CreateFrontierChart(ByVal dashboard As Worksheet, _
                                ByVal chartData As Worksheet, _
                                ByVal topLeft As String, _
                                ByVal bottomRight As String)
    Dim chartObject As ChartObject
    Dim lastRow As Long
    Dim seriesObject As Series

    lastRow = chartData.Cells(chartData.Rows.Count, 22).End(xlUp).Row
    Set chartObject = AddChartBox(dashboard, topLeft, bottomRight)

    With chartObject.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = True
        .ChartTitle.Text = "Historical efficient frontier"
        .HasLegend = False
        Set seriesObject = .SeriesCollection.NewSeries
        seriesObject.Name = "Efficient frontier"
        seriesObject.XValues = chartData.Range("V2:V" & lastRow)
        seriesObject.Values = chartData.Range("W2:W" & lastRow)
    End With
    SafeFormatAxes chartObject.Chart, "0.00", "0.0000"
End Sub

Private Function AddChartBox(ByVal ws As Worksheet, _
                             ByVal topLeft As String, _
                             ByVal bottomRight As String) As ChartObject
    Dim leftValue As Double
    Dim topValue As Double
    Dim widthValue As Double
    Dim heightValue As Double

    leftValue = ws.Range(topLeft).Left
    topValue = ws.Range(topLeft).Top
    widthValue = ws.Range(bottomRight).Left + ws.Range(bottomRight).Width - leftValue
    heightValue = ws.Range(bottomRight).Top + ws.Range(bottomRight).Height - topValue

    Set AddChartBox = ws.ChartObjects.Add(leftValue, topValue, widthValue, heightValue)
End Function

Private Sub SafeFormatAxes(ByVal chartObject As Chart, _
                           ByVal valueFormat As String, _
                           ByVal categoryFormat As String)
    On Error Resume Next
    chartObject.Axes(xlValue).TickLabels.NumberFormat = valueFormat
    chartObject.Axes(xlCategory).TickLabels.NumberFormat = categoryFormat
    On Error GoTo 0
End Sub

Private Sub FormatCard(ByVal target As Range, _
                       ByVal labelText As String, _
                       ByVal valueText As String)
    target.Interior.Color = COLOR_PALE
    target.Borders.LineStyle = xlContinuous
    target.Borders.Color = RGB(220, 225, 230)
    target.Cells(1, 1).Value = labelText
    target.Cells(1, 1).Font.Bold = True
    target.Cells(1, 1).Font.Color = RGB(110, 120, 130)
    target.Cells(2, 1).Value = valueText
    target.Cells(2, 1).Font.Bold = True
    target.Cells(2, 1).Font.Size = 12
End Sub

Private Sub DeleteAllCharts(ByVal ws As Worksheet)
    Dim chartObject As ChartObject
    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
End Sub

' ============================================================================
' General helpers
' ============================================================================

Private Function FindCurveHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 10
        If FindHeaderColumn(ws, r, "DATE") > 0 And _
           FindHeaderColumn(ws, r, "ON") > 0 And _
           FindHeaderColumn(ws, r, "1M") > 0 And _
           FindHeaderColumn(ws, r, "2M") > 0 And _
           FindHeaderColumn(ws, r, "3M") > 0 And _
           FindHeaderColumn(ws, r, "6M") > 0 Then
            FindCurveHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, _
                                  ByVal headerRow As Long, _
                                  ByVal headerText As String) As Long
    Dim finalColumn As Long
    Dim c As Long

    finalColumn = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To finalColumn
        If UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value2))) = UCase$(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function FirstCurveDateOnOrAfter(ByVal targetDate As Double) As Double
    Dim indexValue As Long
    indexValue = FindIndexOnOrAfter(targetDate)
    If indexValue > 0 Then FirstCurveDateOnOrAfter = gCurveDates(indexValue)
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

Private Function CurveDateOnOrBefore(ByVal targetDate As Double) As Double
    Dim indexValue As Long
    indexValue = FindIndexOnOrBefore(targetDate)
    If indexValue = 0 Then
        Err.Raise vbObjectError + 600, , "No curve date on or before " & Format$(targetDate, "dd-mmm-yyyy")
    End If
    CurveDateOnOrBefore = gCurveDates(indexValue)
End Function

Private Function NextCurveDateAfter(ByVal targetDate As Double) As Double
    Dim indexValue As Long
    indexValue = FindIndexOnOrAfter(targetDate + 0.0000001)
    If indexValue > 0 Then NextCurveDateAfter = gCurveDates(indexValue)
End Function

Private Function RateOnDate(ByVal curveDate As Double, ByVal tenorIndex As Long) As Double
    Dim indexValue As Long
    indexValue = FindIndexOnOrBefore(curveDate)
    If indexValue = 0 Or gCurveDates(indexValue) <> curveDate Then
        Err.Raise vbObjectError + 601, , _
            "No exact curve observation on " & Format$(curveDate, "dd-mmm-yyyy") & "."
    End If
    RateOnDate = gRates(indexValue, tenorIndex)
End Function

Private Function AddMonthsPreserveEndOfMonth(ByVal startDate As Double, _
                                             ByVal monthCount As Long) As Double
    Dim sourceDate As Date
    Dim candidateDate As Date
    Dim sourceLastDay As Long
    Dim targetLastDay As Long
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim totalMonths As Long
    Dim targetDay As Long

    sourceDate = CDate(startDate)
    totalMonths = Year(sourceDate) * 12 + Month(sourceDate) - 1 + monthCount
    targetYear = totalMonths \ 12
    targetMonth = totalMonths Mod 12 + 1
    sourceLastDay = Day(DateSerial(Year(sourceDate), Month(sourceDate) + 1, 0))
    targetLastDay = Day(DateSerial(targetYear, targetMonth + 1, 0))

    If Day(sourceDate) = sourceLastDay Then
        targetDay = targetLastDay
    Else
        targetDay = Day(sourceDate)
        If targetDay > targetLastDay Then targetDay = targetLastDay
    End If

    candidateDate = DateSerial(targetYear, targetMonth, targetDay)
    AddMonthsPreserveEndOfMonth = CDbl(candidateDate)
End Function

Private Function TenorName(ByVal tenorIndex As Long) As String
    Select Case tenorIndex
        Case 1: TenorName = "ON"
        Case 2: TenorName = "1M"
        Case 3: TenorName = "2M"
        Case 4: TenorName = "3M"
        Case 5: TenorName = "6M"
        Case Else: Err.Raise vbObjectError + 602, , "Invalid tenor index."
    End Select
End Function

Private Function TenorMonths(ByVal tenorIndex As Long) As Long
    Select Case tenorIndex
        Case 1: TenorMonths = 0
        Case 2: TenorMonths = 1
        Case 3: TenorMonths = 2
        Case 4: TenorMonths = 3
        Case 5: TenorMonths = 6
        Case Else: Err.Raise vbObjectError + 603, , "Invalid tenor index."
    End Select
End Function

Private Function CollectionToMatrix(ByVal rows As Collection, _
                                    ByVal columnCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowData As Variant
    Dim r As Long
    Dim c As Long

    ReDim matrix(1 To rows.Count, 1 To columnCount)
    For r = 1 To rows.Count
        rowData = rows(r)
        For c = 1 To columnCount
            matrix(r, c) = rowData(c - 1)
        Next c
    Next r
    CollectionToMatrix = matrix
End Function

Private Sub SortDoubleArray(ByRef values() As Double)
    QuickSortDouble values, LBound(values), UBound(values)
End Sub

Private Sub QuickSortDouble(ByRef values() As Double, _
                            ByVal lowValue As Long, _
                            ByVal highValue As Long)
    Dim i As Long
    Dim j As Long
    Dim pivotValue As Double
    Dim temporaryValue As Double

    i = lowValue
    j = highValue
    pivotValue = values((lowValue + highValue) \ 2)

    Do While i <= j
        Do While values(i) < pivotValue
            i = i + 1
        Loop
        Do While values(j) > pivotValue
            j = j - 1
        Loop
        If i <= j Then
            temporaryValue = values(i)
            values(i) = values(j)
            values(j) = temporaryValue
            i = i + 1
            j = j - 1
        End If
    Loop

    If lowValue < j Then QuickSortDouble values, lowValue, j
    If i < highValue Then QuickSortDouble values, i, highValue
End Sub

Private Function SampleStdDev(ByRef values() As Double) As Double
    Dim i As Long
    Dim countValue As Long
    Dim averageValue As Double
    Dim sumSquared As Double

    countValue = UBound(values) - LBound(values) + 1
    If countValue < 2 Then Exit Function

    For i = LBound(values) To UBound(values)
        averageValue = averageValue + values(i)
    Next i
    averageValue = averageValue / countValue

    For i = LBound(values) To UBound(values)
        sumSquared = sumSquared + (values(i) - averageValue) ^ 2
    Next i
    SampleStdDev = Sqr(sumSquared / (countValue - 1))
End Function

Private Function PercentileSorted(ByRef values() As Double, _
                                  ByVal percentileValue As Double) As Double
    Dim positionValue As Double
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim fractionValue As Double

    positionValue = (UBound(values) - LBound(values)) * percentileValue + LBound(values)
    lowerIndex = Int(positionValue)
    upperIndex = lowerIndex
    If positionValue > lowerIndex Then upperIndex = lowerIndex + 1

    If lowerIndex = upperIndex Then
        PercentileSorted = values(lowerIndex)
    Else
        fractionValue = positionValue - lowerIndex
        PercentileSorted = values(lowerIndex) + _
                           fractionValue * (values(upperIndex) - values(lowerIndex))
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

Private Sub QuickSortPortfolio(ByVal lowValue As Long, ByVal highValue As Long)
    Dim i As Long
    Dim j As Long
    Dim pivotVol As Double
    Dim pivotReturn As Double

    i = lowValue
    j = highValue
    pivotVol = gPortfolio((lowValue + highValue) \ 2, 7)
    pivotReturn = gPortfolio((lowValue + highValue) \ 2, 6)

    Do While i <= j
        Do While PortfolioComesBefore(i, pivotVol, pivotReturn)
            i = i + 1
            If i > highValue Then Exit Do
        Loop
        Do While PortfolioComesAfter(j, pivotVol, pivotReturn)
            j = j - 1
            If j < lowValue Then Exit Do
        Loop
        If i <= j Then
            SwapPortfolioRows i, j
            i = i + 1
            j = j - 1
        End If
    Loop

    If lowValue < j Then QuickSortPortfolio lowValue, j
    If i < highValue Then QuickSortPortfolio i, highValue
End Sub

Private Function PortfolioComesBefore(ByVal rowIndex As Long, _
                                      ByVal pivotVol As Double, _
                                      ByVal pivotReturn As Double) As Boolean
    If gPortfolio(rowIndex, 7) < pivotVol - 0.0000000001 Then
        PortfolioComesBefore = True
    ElseIf Abs(gPortfolio(rowIndex, 7) - pivotVol) <= 0.0000000001 Then
        PortfolioComesBefore = (gPortfolio(rowIndex, 6) > pivotReturn)
    End If
End Function

Private Function PortfolioComesAfter(ByVal rowIndex As Long, _
                                     ByVal pivotVol As Double, _
                                     ByVal pivotReturn As Double) As Boolean
    If gPortfolio(rowIndex, 7) > pivotVol + 0.0000000001 Then
        PortfolioComesAfter = True
    ElseIf Abs(gPortfolio(rowIndex, 7) - pivotVol) <= 0.0000000001 Then
        PortfolioComesAfter = (gPortfolio(rowIndex, 6) < pivotReturn)
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

Private Function FindPortfolioRow(ByVal wON As Double, ByVal w1M As Double, _
                                  ByVal w2M As Double, ByVal w3M As Double, _
                                  ByVal w6M As Double) As Long
    Dim p As Long
    For p = 1 To gPortfolioCount
        If Abs(gPortfolio(p, 1) - wON) < 0.0000001 And _
           Abs(gPortfolio(p, 2) - w1M) < 0.0000001 And _
           Abs(gPortfolio(p, 3) - w2M) < 0.0000001 And _
           Abs(gPortfolio(p, 4) - w3M) < 0.0000001 And _
           Abs(gPortfolio(p, 5) - w6M) < 0.0000001 Then
            FindPortfolioRow = p
            Exit Function
        End If
    Next p
    Err.Raise vbObjectError + 604, , "Requested portfolio was not generated."
End Function

Private Function MinimumVolatilityRow() As Long
    Dim p As Long
    Dim bestRow As Long
    bestRow = 1
    For p = 2 To gPortfolioCount
        If gPortfolio(p, 7) < gPortfolio(bestRow, 7) Then bestRow = p
    Next p
    MinimumVolatilityRow = bestRow
End Function

Private Function MaximumReturnRow() As Long
    Dim p As Long
    Dim bestRow As Long
    bestRow = 1
    For p = 2 To gPortfolioCount
        If gPortfolio(p, 6) > gPortfolio(bestRow, 6) Then bestRow = p
    Next p
    MaximumReturnRow = bestRow
End Function

Private Function FrontierSegment(ByVal rank As Long, ByVal totalCount As Long) As String
    Dim positionValue As Double

    If rank = 1 Then
        FrontierSegment = "Minimum Volatility"
        Exit Function
    End If
    If totalCount <= 1 Then
        FrontierSegment = "Maximum Historical Return"
        Exit Function
    End If

    positionValue = (rank - 1) / (totalCount - 1)
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
            FrontierDescription = "Lowest historical monthly earnings volatility on the frontier."
        Case "Defensive"
            FrontierDescription = "Small volatility increase for incremental historical return."
        Case "Conservative"
            FrontierDescription = "Moderate return improvement in the lower half of frontier risk."
        Case "Balanced"
            FrontierDescription = "Balances historical return, monthly volatility and liquidity."
        Case "Return Oriented"
            FrontierDescription = "Higher historical return with greater volatility or maturity exposure."
        Case Else
            FrontierDescription = "Highest historical return among efficient tested allocations."
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
        Format$(gFrontier(frontierRow, 10), "0%") & " available within 30 days."
End Function

Private Function SumDailyInterest(ByVal tenorIndex As Long) As Double
    Dim d As Long
    For d = 0 To gNumDays - 1
        SumDailyInterest = SumDailyInterest + gDailyInterest(tenorIndex, d)
    Next d
End Function

Private Function BlankIfZero(ByVal valueNumber As Double) As Variant
    If valueNumber = 0 Then
        BlankIfZero = vbNullString
    Else
        BlankIfZero = valueNumber
    End If
End Function

Private Function MinDouble(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue < secondValue Then
        MinDouble = firstValue
    Else
        MinDouble = secondValue
    End If
End Function

Private Function MaxLong(ByVal firstValue As Long, ByVal secondValue As Long) As Long
    If firstValue > secondValue Then
        MaxLong = firstValue
    Else
        MaxLong = secondValue
    End If
End Function

Private Sub ClearOutputs()
    Dim names As Variant
    Dim item As Variant
    Dim ws As Worksheet

    names = Array("Transactions", "Daily_Accrual", "Rolling_Results", _
                  "Tenor_Analysis", "Monthly_Returns", "Portfolio_Analysis", _
                  "Chart_Data", "Dashboard", "Test_Results")

    For Each item In names
        Set ws = ThisWorkbook.Worksheets(CStr(item))
        DeleteAllCharts ws
        ClearUsedBody ws
    Next item
End Sub

Private Sub ClearSheetBody(ByVal ws As Worksheet)
    ClearUsedBody ws
End Sub

Private Sub ClearUsedBody(ByVal ws As Worksheet)
    Dim finalCell As Range
    Dim lastRow As Long
    Dim lastColumn As Long

    On Error Resume Next
    Set finalCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                                  LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If finalCell Is Nothing Then Exit Sub
    lastRow = finalCell.Row

    On Error Resume Next
    Set finalCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                                  LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If finalCell Is Nothing Then Exit Sub
    lastColumn = finalCell.Column

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastColumn)).Clear
    End If
End Sub

Private Sub SetStandardWidths(ByVal ws As Worksheet, ByVal columnCount As Long)
    Dim c As Long
    For c = 1 To columnCount
        ws.Columns(c).ColumnWidth = 14
    Next c
    ws.Columns(1).ColumnWidth = 16
End Sub

Private Sub SetDashboardWidths(ByVal ws As Worksheet)
    Dim c As Long
    For c = 1 To 17
        ws.Columns(c).ColumnWidth = 11
    Next c
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 3
    ws.Rows("55:60").RowHeight = 24
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
