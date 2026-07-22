Attribute VB_Name = "RatesRolloverAuditDatabase"
Option Explicit

' ============================================================================
' ROLLOVER RETURN AUDIT DATABASE
'
' Purpose:
'   Build transparent transaction-level databases for auditing tenor returns,
'   rollover mechanics, annualized-equivalent returns, and return volatility.
'
' Public macro:
'   BuildRolloverAuditDatabase
'
' Required input sheets:
'   Inputs
'     B5 = analysis start date
'     B6 = analysis end date
'     B7 = initial notional
'
'   Curve
'     Required headers: Date | ON | 1M | 2M | 3M | 6M
'
' Output sheets created or replaced:
'   Rollover_Path_DB
'       Sequential rollover path used by an actual investment strategy.
'
'   Rollover_All_Starts
'       One completed tenor-return observation for every eligible curve-date
'       start. These overlapping observations provide a larger historical
'       database for tenor return-distribution analysis.
'
'   Rollover_Audit_Summary
'       Summary statistics for both databases, including the volatility of
'       holding-period returns and annualized-equivalent returns.
'
' Methodology:
'   - Rates are ordinary numbers. 4.31 means 4.31 percent.
'   - ACT/360 simple interest.
'   - Interest is reinvested on the sequential path.
'   - Each sequential term begins on the prior actual maturity.
'   - Target maturity is start date plus tenor.
'   - Actual maturity is the latest curve date on or before target maturity.
'   - ON matures on the next available curve date.
'   - Only completed observations inside Inputs!B5:B6 are retained.
'   - Return volatility is calculated separately by tenor.
'   - No frontier, optimization, or dashboard logic is included.
' ============================================================================

Private Const TENOR_COUNT As Long = 5
Private Const DAY_COUNT As Double = 360#
Private Const COLOR_NAVY As Long = 3809035
Private Const COLOR_PALE As Long = 16185078

Private gCurveDates() As Double
Private gRates() As Double
Private gCurveCount As Long

Private gRequestedStart As Double
Private gStartDate As Double
Private gEndDate As Double
Private gInitialNotional As Double

Public Sub BuildRolloverAuditDatabase()
    Dim oldCalculation As XlCalculation
    Dim stageName As String

    On Error GoTo Fail

    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stageName = "loading curve"
    LoadAuditCurve

    stageName = "loading inputs"
    LoadAuditInputs
    ValidateAuditInputs

    stageName = "preparing output sheets"
    PrepareAuditSheet "Rollover_Path_DB"
    PrepareAuditSheet "Rollover_All_Starts"
    PrepareAuditSheet "Rollover_Audit_Summary"

    stageName = "building sequential rollover database"
    BuildSequentialRolloverDatabase

    stageName = "building all-start observations"
    BuildAllStartsDatabase

    stageName = "building audit summary"
    BuildAuditSummary

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Rollover audit database completed." & vbCrLf & _
           "Review Rollover_Path_DB, Rollover_All_Starts, and " & _
           "Rollover_Audit_Summary.", vbInformation
    Exit Sub

Fail:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Rollover audit database stopped during " & stageName & "." & _
           vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================================
' SECTION 01 - INPUTS AND CURVE
' ============================================================================

Private Sub LoadAuditInputs()
    Dim ws As Worksheet

    If Not AuditSheetExists("Inputs") Then
        Err.Raise vbObjectError + 100, , "Inputs sheet was not found."
    End If

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

    gRequestedStart = CDbl(CDate(ws.Range("B5").Value))
    gEndDate = CDbl(CDate(ws.Range("B6").Value))
    gInitialNotional = CDbl(ws.Range("B7").Value2)
End Sub

Private Sub LoadAuditCurve()
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

    If Not AuditSheetExists("Curve") Then
        Err.Raise vbObjectError + 110, , "Curve sheet was not found."
    End If

    Set ws = ThisWorkbook.Worksheets("Curve")
    headerRow = FindAuditHeaderRow(ws)

    If headerRow = 0 Then
        Err.Raise vbObjectError + 111, , _
            "Curve headers not found. Required: Date, ON, 1M, 2M, 3M, 6M."
    End If

    dateColumn = FindAuditHeaderColumn(ws, headerRow, "DATE")
    tenorColumn(1) = FindAuditHeaderColumn(ws, headerRow, "ON")
    tenorColumn(2) = FindAuditHeaderColumn(ws, headerRow, "1M")
    tenorColumn(3) = FindAuditHeaderColumn(ws, headerRow, "2M")
    tenorColumn(4) = FindAuditHeaderColumn(ws, headerRow, "3M")
    tenorColumn(5) = FindAuditHeaderColumn(ws, headerRow, "6M")

    lastRow = ws.Cells(ws.Rows.Count, dateColumn).End(xlUp).Row

    gCurveCount = 0

    For rowNumber = headerRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, dateColumn).Value2))) > 0 Then
            gCurveCount = gCurveCount + 1
        End If
    Next rowNumber

    If gCurveCount < 2 Then
        Err.Raise vbObjectError + 112, , _
            "Curve requires at least two dated rows."
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
            Err.Raise vbObjectError + 113, , _
                "Invalid curve date on row " & rowNumber & "."
        End If

        For tenorIndex = 1 To TENOR_COUNT
            rawRate = ws.Cells(rowNumber, tenorColumn(tenorIndex)).Value2

            If Not IsNumeric(rawRate) Then
                Err.Raise vbObjectError + 114, , _
                    "Non-numeric " & AuditTenorName(tenorIndex) & _
                    " rate on row " & rowNumber & "."
            End If

            gRates(curveIndex, tenorIndex) = CDbl(rawRate)
        Next tenorIndex

NextCurveRow:
    Next rowNumber

    QuickSortAuditCurve 1, gCurveCount
End Sub

Private Sub ValidateAuditInputs()
    Dim curveIndex As Long
    Dim tenorIndex As Long
    Dim averageAbsoluteRate As Double
    Dim observationCount As Long

    If gRequestedStart > gEndDate Then
        Err.Raise vbObjectError + 120, , _
            "Analysis start date is after the end date."
    End If

    If gInitialNotional <= 0 Then
        Err.Raise vbObjectError + 121, , _
            "Initial notional must be positive."
    End If

    For curveIndex = 2 To gCurveCount
        If gCurveDates(curveIndex) <= gCurveDates(curveIndex - 1) Then
            Err.Raise vbObjectError + 122, , _
                "Curve dates must be unique. Duplicate date: " & _
                Format$(gCurveDates(curveIndex), "dd-mmm-yyyy") & "."
        End If
    Next curveIndex

    If gEndDate > gCurveDates(gCurveCount) Then
        Err.Raise vbObjectError + 123, , _
            "Analysis end date exceeds the final curve date."
    End If

    gStartDate = FirstAuditCurveDateOnOrAfter(gRequestedStart)

    If gStartDate = 0 Or gStartDate > gEndDate Then
        Err.Raise vbObjectError + 124, , _
            "No curve date exists in the requested analysis period."
    End If

    For curveIndex = 1 To gCurveCount
        For tenorIndex = 1 To TENOR_COUNT
            averageAbsoluteRate = averageAbsoluteRate + _
                                  Abs(gRates(curveIndex, tenorIndex))
            observationCount = observationCount + 1
        Next tenorIndex
    Next curveIndex

    averageAbsoluteRate = averageAbsoluteRate / observationCount

    If averageAbsoluteRate < 0.25 Then
        Err.Raise vbObjectError + 125, , _
            "Rates appear to be Excel percentages. Enter 4.31 for 4.31 percent."
    End If
End Sub

' ============================================================================
' SECTION 02 - SEQUENTIAL ROLLOVER DATABASE
' ============================================================================

Private Sub BuildSequentialRolloverDatabase()
    Dim ws As Worksheet
    Dim outputRows As Collection
    Dim tenorIndex As Long
    Dim transactionID As Long
    Dim currentStart As Double
    Dim openingPrincipal As Double
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim startRateIndex As Long
    Dim maturityRateIndex As Long
    Dim nextRateIndex As Long
    Dim startRateDate As Double
    Dim maturityRateDate As Double
    Dim nextRateDate As Double
    Dim startRate As Double
    Dim maturitySameTenorRate As Double
    Dim nextRolloverRate As Double
    Dim holdingDays As Long
    Dim interestAmount As Double
    Dim closingPrincipal As Double
    Dim holdingReturnPct As Double
    Dim annualizedEquivalentPct As Double
    Dim sameTenorRateChangeBps As Double
    Dim nextRolloverRateChangeBps As Double
    Dim maturityAdjustmentDays As Long
    Dim rowData As Variant
    Dim outputMatrix As Variant

    Set ws = ThisWorkbook.Worksheets("Rollover_Path_DB")
    Set outputRows = New Collection

    WriteAuditDatabaseHeader ws, _
        "Sequential Rollover Path Database", _
        "Actual strategy path: each new transaction starts at the prior actual maturity."

    For tenorIndex = 1 To TENOR_COUNT
        currentStart = gStartDate
        openingPrincipal = gInitialNotional
        transactionID = 1

        Do While currentStart <= gEndDate
            startRateIndex = FindAuditIndexOnOrBefore(currentStart)
            startRateDate = gCurveDates(startRateIndex)
            startRate = gRates(startRateIndex, tenorIndex)

            ResolveAuditMaturity tenorIndex, currentStart, _
                                 targetMaturity, actualMaturity

            If actualMaturity = 0 Or actualMaturity > gEndDate Then Exit Do
            If actualMaturity <= currentStart Then Exit Do

            holdingDays = CLng(actualMaturity - currentStart)
            interestAmount = openingPrincipal * _
                             (startRate / 100#) * _
                             holdingDays / DAY_COUNT
            closingPrincipal = openingPrincipal + interestAmount
            holdingReturnPct = interestAmount / openingPrincipal * 100#

            annualizedEquivalentPct = _
                ((closingPrincipal / openingPrincipal) ^ _
                 (DAY_COUNT / holdingDays) - 1#) * 100#

            maturityRateIndex = FindAuditIndexOnOrBefore(actualMaturity)
            maturityRateDate = gCurveDates(maturityRateIndex)
            maturitySameTenorRate = _
                gRates(maturityRateIndex, tenorIndex)

            nextRateIndex = FindAuditIndexOnOrBefore(actualMaturity)
            nextRateDate = gCurveDates(nextRateIndex)
            nextRolloverRate = gRates(nextRateIndex, tenorIndex)

            sameTenorRateChangeBps = _
                (maturitySameTenorRate - startRate) * 100#
            nextRolloverRateChangeBps = _
                (nextRolloverRate - startRate) * 100#
            maturityAdjustmentDays = _
                CLng(actualMaturity - targetMaturity)

            rowData = Array( _
                "SEQUENTIAL", _
                AuditTenorName(tenorIndex), _
                transactionID, _
                currentStart, _
                startRateDate, _
                startRate, _
                targetMaturity, _
                actualMaturity, _
                maturityAdjustmentDays, _
                holdingDays, _
                openingPrincipal, _
                interestAmount, _
                closingPrincipal, _
                holdingReturnPct, _
                annualizedEquivalentPct, _
                maturityRateDate, _
                maturitySameTenorRate, _
                sameTenorRateChangeBps, _
                nextRateDate, _
                nextRolloverRate, _
                nextRolloverRateChangeBps)

            outputRows.Add rowData

            openingPrincipal = closingPrincipal
            currentStart = actualMaturity
            transactionID = transactionID + 1
        Loop
    Next tenorIndex

    If outputRows.Count > 0 Then
        outputMatrix = AuditCollectionToMatrix(outputRows, 21)
        ws.Range("A5").Resize(outputRows.Count, 21).Value = outputMatrix
        FormatAuditDatabase ws, outputRows.Count
    End If
End Sub

' ============================================================================
' SECTION 03 - ALL-START RETURN DATABASE
' ============================================================================

Private Sub BuildAllStartsDatabase()
    Dim ws As Worksheet
    Dim outputRows As Collection
    Dim firstCurveIndex As Long
    Dim finalCurveIndex As Long
    Dim startCurveIndex As Long
    Dim tenorIndex As Long
    Dim targetMaturity As Double
    Dim actualMaturity As Double
    Dim maturityRateIndex As Long
    Dim startRate As Double
    Dim maturityRate As Double
    Dim holdingDays As Long
    Dim interestAmount As Double
    Dim closingValue As Double
    Dim holdingReturnPct As Double
    Dim annualizedEquivalentPct As Double
    Dim rateChangeBps As Double
    Dim maturityAdjustmentDays As Long
    Dim rowData As Variant
    Dim outputMatrix As Variant

    Set ws = ThisWorkbook.Worksheets("Rollover_All_Starts")
    Set outputRows = New Collection

    WriteAuditDatabaseHeader ws, _
        "All Eligible Start-Date Return Database", _
        "Overlapping observations: one completed return for every eligible curve-date start."

    firstCurveIndex = FindAuditIndexOnOrAfter(gStartDate)
    finalCurveIndex = FindAuditIndexOnOrBefore(gEndDate)

    For tenorIndex = 1 To TENOR_COUNT
        For startCurveIndex = firstCurveIndex To finalCurveIndex
            ResolveAuditMaturity tenorIndex, _
                                 gCurveDates(startCurveIndex), _
                                 targetMaturity, actualMaturity

            If actualMaturity = 0 Or actualMaturity > gEndDate Then Exit For
            If actualMaturity <= gCurveDates(startCurveIndex) Then _
                GoTo NextStartObservation

            startRate = gRates(startCurveIndex, tenorIndex)
            maturityRateIndex = FindAuditIndexOnOrBefore(actualMaturity)
            maturityRate = gRates(maturityRateIndex, tenorIndex)
            holdingDays = CLng(actualMaturity - gCurveDates(startCurveIndex))

            interestAmount = gInitialNotional * _
                             (startRate / 100#) * _
                             holdingDays / DAY_COUNT
            closingValue = gInitialNotional + interestAmount
            holdingReturnPct = interestAmount / gInitialNotional * 100#

            annualizedEquivalentPct = _
                ((closingValue / gInitialNotional) ^ _
                 (DAY_COUNT / holdingDays) - 1#) * 100#

            rateChangeBps = (maturityRate - startRate) * 100#
            maturityAdjustmentDays = _
                CLng(actualMaturity - targetMaturity)

            rowData = Array( _
                "ALL_STARTS", _
                AuditTenorName(tenorIndex), _
                startCurveIndex - firstCurveIndex + 1, _
                gCurveDates(startCurveIndex), _
                gCurveDates(startCurveIndex), _
                startRate, _
                targetMaturity, _
                actualMaturity, _
                maturityAdjustmentDays, _
                holdingDays, _
                gInitialNotional, _
                interestAmount, _
                closingValue, _
                holdingReturnPct, _
                annualizedEquivalentPct, _
                gCurveDates(maturityRateIndex), _
                maturityRate, _
                rateChangeBps, _
                gCurveDates(maturityRateIndex), _
                maturityRate, _
                rateChangeBps)

            outputRows.Add rowData

NextStartObservation:
        Next startCurveIndex
    Next tenorIndex

    If outputRows.Count > 0 Then
        outputMatrix = AuditCollectionToMatrix(outputRows, 21)
        ws.Range("A5").Resize(outputRows.Count, 21).Value = outputMatrix
        FormatAuditDatabase ws, outputRows.Count
    End If
End Sub

' ============================================================================
' SECTION 04 - RETURN AND VOLATILITY SUMMARY
' ============================================================================

Private Sub BuildAuditSummary()
    Dim ws As Worksheet
    Dim databaseNames As Variant
    Dim databaseLabel As Variant
    Dim sourceSheet As Worksheet
    Dim lastRow As Long
    Dim tenorIndex As Long
    Dim summaryRow As Long
    Dim sourceRow As Long
    Dim observationCount As Long
    Dim holdingDaysValues() As Double
    Dim holdingReturnValues() As Double
    Dim annualizedReturnValues() As Double
    Dim rateChangeValues() As Double
    Dim indexValue As Long
    Dim totalInterest As Double
    Dim endingPrincipal As Double

    Set ws = ThisWorkbook.Worksheets("Rollover_Audit_Summary")

    ws.Range("A1:Q1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = "Rollover Return and Volatility Audit Summary"
    StyleAuditTitle ws.Range("A1:Q1")

    ws.Range("A2").Value = _
        "Holding-period return volatility and annualized-equivalent return volatility are calculated separately by tenor."
    ws.Range("A2:Q2").Interior.Color = COLOR_PALE
    ws.Range("A2:Q2").WrapText = True

    WriteAuditRow ws.Range("A4"), Array( _
        "Database", _
        "Tenor", _
        "Observations", _
        "Average Holding Days", _
        "Average Start Rate", _
        "Average Holding Return", _
        "Holding Return Volatility", _
        "Average Annualized Equivalent Return", _
        "Annualized Equivalent Return Volatility", _
        "5th Percentile Annualized Return", _
        "Median Annualized Return", _
        "95th Percentile Annualized Return", _
        "Average Rate Change (bps)", _
        "Rate Change Volatility (bps)", _
        "Total Interest ($)", _
        "Final Sequential Principal ($)", _
        "Audit Note"))

    StyleAuditHeader ws.Range("A4:Q4")

    databaseNames = Array("Rollover_Path_DB", "Rollover_All_Starts")
    summaryRow = 5

    For Each databaseLabel In databaseNames
        Set sourceSheet = ThisWorkbook.Worksheets(CStr(databaseLabel))
        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row

        For tenorIndex = 1 To TENOR_COUNT
            observationCount = CountAuditTenorRows( _
                sourceSheet, lastRow, AuditTenorName(tenorIndex))

            If observationCount > 0 Then
                ReDim holdingDaysValues(1 To observationCount)
                ReDim holdingReturnValues(1 To observationCount)
                ReDim annualizedReturnValues(1 To observationCount)
                ReDim rateChangeValues(1 To observationCount)

                indexValue = 0
                totalInterest = 0#
                endingPrincipal = 0#

                For sourceRow = 5 To lastRow
                    If CStr(sourceSheet.Cells(sourceRow, 2).Value2) = _
                       AuditTenorName(tenorIndex) Then

                        indexValue = indexValue + 1
                        holdingDaysValues(indexValue) = _
                            CDbl(sourceSheet.Cells(sourceRow, 10).Value2)
                        holdingReturnValues(indexValue) = _
                            CDbl(sourceSheet.Cells(sourceRow, 14).Value2)
                        annualizedReturnValues(indexValue) = _
                            CDbl(sourceSheet.Cells(sourceRow, 15).Value2)
                        rateChangeValues(indexValue) = _
                            CDbl(sourceSheet.Cells(sourceRow, 18).Value2)

                        totalInterest = totalInterest + _
                            CDbl(sourceSheet.Cells(sourceRow, 12).Value2)
                        endingPrincipal = _
                            CDbl(sourceSheet.Cells(sourceRow, 13).Value2)
                    End If
                Next sourceRow

                SortAuditDoubleArray annualizedReturnValues

                ws.Cells(summaryRow, 1).Value = AuditDatabaseDisplayName( _
                    CStr(databaseLabel))
                ws.Cells(summaryRow, 2).Value = _
                    AuditTenorName(tenorIndex)
                ws.Cells(summaryRow, 3).Value = observationCount
                ws.Cells(summaryRow, 4).Value = _
                    AuditAverage(holdingDaysValues)
                ws.Cells(summaryRow, 5).Value = _
                    AverageAuditStartRate(sourceSheet, lastRow, _
                                          AuditTenorName(tenorIndex))
                ws.Cells(summaryRow, 6).Value = _
                    AuditAverage(holdingReturnValues)
                ws.Cells(summaryRow, 7).Value = _
                    AuditSampleStdDev(holdingReturnValues)
                ws.Cells(summaryRow, 8).Value = _
                    AuditAverage(annualizedReturnValues)
                ws.Cells(summaryRow, 9).Value = _
                    AuditSampleStdDev(annualizedReturnValues)
                ws.Cells(summaryRow, 10).Value = _
                    AuditPercentileSorted(annualizedReturnValues, 0.05)
                ws.Cells(summaryRow, 11).Value = _
                    AuditPercentileSorted(annualizedReturnValues, 0.5)
                ws.Cells(summaryRow, 12).Value = _
                    AuditPercentileSorted(annualizedReturnValues, 0.95)
                ws.Cells(summaryRow, 13).Value = _
                    AuditAverage(rateChangeValues)
                ws.Cells(summaryRow, 14).Value = _
                    AuditSampleStdDev(rateChangeValues)
                ws.Cells(summaryRow, 15).Value = totalInterest

                If CStr(databaseLabel) = "Rollover_Path_DB" Then
                    ws.Cells(summaryRow, 16).Value = endingPrincipal
                    ws.Cells(summaryRow, 17).Value = _
                        "Actual sequential rollover path."
                Else
                    ws.Cells(summaryRow, 16).Value = vbNullString
                    ws.Cells(summaryRow, 17).Value = _
                        "Overlapping starts; use for return-distribution audit."
                End If

                summaryRow = summaryRow + 1
            End If
        Next tenorIndex
    Next databaseLabel

    FormatAuditSummary ws, summaryRow - 1
End Sub

' ============================================================================
' SECTION 05 - OUTPUT FORMATTING
' ============================================================================

Private Sub WriteAuditDatabaseHeader(ByVal ws As Worksheet, _
                                     ByVal titleText As String, _
                                     ByVal subtitleText As String)
    ws.Range("A1:U1").Interior.Color = COLOR_NAVY
    ws.Range("A1").Value = titleText
    StyleAuditTitle ws.Range("A1:U1")

    ws.Range("A2").Value = subtitleText
    ws.Range("A2:U2").Interior.Color = COLOR_PALE
    ws.Range("A2:U2").WrapText = True

    WriteAuditRow ws.Range("A4"), Array( _
        "Database Type", _
        "Tenor", _
        "Observation / Transaction ID", _
        "Start Date", _
        "Start Rate Observation Date", _
        "Start Rate", _
        "Target Maturity", _
        "Actual Maturity", _
        "Maturity Adjustment Days", _
        "Holding Days", _
        "Opening Principal ($)", _
        "Interest ($)", _
        "Closing Principal / Value ($)", _
        "Holding Period Return", _
        "Annualized Equivalent Return", _
        "Maturity Rate Observation Date", _
        "Same-Tenor Rate at Maturity", _
        "Same-Tenor Rate Change (bps)", _
        "Next Rollover Rate Date", _
        "Next Rollover Rate", _
        "Next Rollover Rate Change (bps)"))

    StyleAuditHeader ws.Range("A4:U4")
End Sub

Private Sub FormatAuditDatabase(ByVal ws As Worksheet, _
                                ByVal dataRowCount As Long)
    Dim lastRow As Long

    lastRow = dataRowCount + 4

    ws.Range("D5:E" & lastRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("G5:H" & lastRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("P5:P" & lastRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("S5:S" & lastRow).NumberFormat = "mm/dd/yyyy"

    ws.Range("F5:F" & lastRow).NumberFormat = "0.0000"
    ws.Range("Q5:Q" & lastRow).NumberFormat = "0.0000"
    ws.Range("T5:T" & lastRow).NumberFormat = "0.0000"

    ws.Range("K5:M" & lastRow).NumberFormat = _
        "$#,##0.00;[Red]($#,##0.00);-"

    ws.Range("N5:O" & lastRow).NumberFormat = _
        "0.000000\%;[Red](0.000000\%);-"

    ws.Range("R5:R" & lastRow).NumberFormat = _
        "0.000;[Red](0.000);-"
    ws.Range("U5:U" & lastRow).NumberFormat = _
        "0.000;[Red](0.000);-"

    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D:E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 11
    ws.Columns("G:H").ColumnWidth = 15
    ws.Columns("I:J").ColumnWidth = 14
    ws.Columns("K:M").ColumnWidth = 20
    ws.Columns("N:O").ColumnWidth = 18
    ws.Columns("P").ColumnWidth = 17
    ws.Columns("Q").ColumnWidth = 17
    ws.Columns("R").ColumnWidth = 18
    ws.Columns("S").ColumnWidth = 17
    ws.Columns("T").ColumnWidth = 15
    ws.Columns("U").ColumnWidth = 20

    ws.Range("A4:U" & lastRow).AutoFilter
End Sub

Private Sub FormatAuditSummary(ByVal ws As Worksheet, _
                               ByVal lastRow As Long)
    ws.Range("C5:C" & lastRow).NumberFormat = "0"
    ws.Range("D5:D" & lastRow).NumberFormat = "0.0"
    ws.Range("E5:E" & lastRow).NumberFormat = "0.0000"

    ws.Range("F5:L" & lastRow).NumberFormat = _
        "0.000000\%;[Red](0.000000\%);-"

    ws.Range("M5:N" & lastRow).NumberFormat = _
        "0.000;[Red](0.000);-"

    ws.Range("O5:P" & lastRow).NumberFormat = _
        "$#,##0.00;[Red]($#,##0.00);-"

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C:N").ColumnWidth = 18
    ws.Columns("O:P").ColumnWidth = 22
    ws.Columns("Q").ColumnWidth = 44
    ws.Range("Q5:Q" & lastRow).WrapText = True
End Sub

Private Sub PrepareAuditSheet(ByVal sheetName As String)
    Dim ws As Worksheet

    If AuditSheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        ws.Cells.Clear

        On Error Resume Next
        ws.AutoFilterMode = False
        On Error GoTo 0
    Else
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

' ============================================================================
' SECTION 06 - MATURITY AND DATE HELPERS
' ============================================================================

Private Sub ResolveAuditMaturity(ByVal tenorIndex As Long, _
                                 ByVal startDate As Double, _
                                 ByRef targetMaturity As Double, _
                                 ByRef actualMaturity As Double)
    If tenorIndex = 1 Then
        targetMaturity = startDate + 1
        actualMaturity = NextAuditCurveDateAfter(startDate)
    Else
        targetMaturity = AddAuditMonthsPreserveEndOfMonth( _
            startDate, AuditTenorMonths(tenorIndex))

        If targetMaturity <= gCurveDates(gCurveCount) Then
            actualMaturity = AuditCurveDateOnOrBefore(targetMaturity)
        Else
            actualMaturity = 0#
        End If
    End If
End Sub

Private Function FindAuditIndexOnOrBefore( _
    ByVal targetDate As Double) As Long

    Dim lowValue As Long
    Dim highValue As Long
    Dim middleValue As Long
    Dim resultValue As Long

    lowValue = 1
    highValue = gCurveCount

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
        Err.Raise vbObjectError + 200, , _
            "No curve date exists on or before " & _
            Format$(targetDate, "dd-mmm-yyyy") & "."
    End If

    FindAuditIndexOnOrBefore = resultValue
End Function

Private Function FindAuditIndexOnOrAfter( _
    ByVal targetDate As Double) As Long

    Dim lowValue As Long
    Dim highValue As Long
    Dim middleValue As Long
    Dim resultValue As Long

    lowValue = 1
    highValue = gCurveCount

    Do While lowValue <= highValue
        middleValue = (lowValue + highValue) \ 2

        If gCurveDates(middleValue) >= targetDate Then
            resultValue = middleValue
            highValue = middleValue - 1
        Else
            lowValue = middleValue + 1
        End If
    Loop

    FindAuditIndexOnOrAfter = resultValue
End Function

Private Function FirstAuditCurveDateOnOrAfter( _
    ByVal targetDate As Double) As Double

    Dim curveIndex As Long

    curveIndex = FindAuditIndexOnOrAfter(targetDate)

    If curveIndex > 0 Then
        FirstAuditCurveDateOnOrAfter = gCurveDates(curveIndex)
    End If
End Function

Private Function AuditCurveDateOnOrBefore( _
    ByVal targetDate As Double) As Double

    AuditCurveDateOnOrBefore = _
        gCurveDates(FindAuditIndexOnOrBefore(targetDate))
End Function

Private Function NextAuditCurveDateAfter( _
    ByVal currentDate As Double) As Double

    Dim curveIndex As Long

    curveIndex = FindAuditIndexOnOrAfter(currentDate + 0.0000001)

    If curveIndex > 0 Then
        NextAuditCurveDateAfter = gCurveDates(curveIndex)
    End If
End Function

Private Function AddAuditMonthsPreserveEndOfMonth( _
    ByVal startDate As Double, _
    ByVal monthCount As Long) As Double

    Dim sourceDate As Date
    Dim targetDate As Date
    Dim sourceLastDay As Long
    Dim targetLastDay As Long
    Dim targetDay As Long

    sourceDate = CDate(startDate)
    targetDate = DateAdd("m", monthCount, sourceDate)

    sourceLastDay = Day(DateSerial( _
        Year(sourceDate), Month(sourceDate) + 1, 0))

    targetLastDay = Day(DateSerial( _
        Year(targetDate), Month(targetDate) + 1, 0))

    If Day(sourceDate) = sourceLastDay Then
        targetDay = targetLastDay
    Else
        targetDay = AuditMinimumLong(Day(sourceDate), targetLastDay)
    End If

    AddAuditMonthsPreserveEndOfMonth = CDbl(DateSerial( _
        Year(targetDate), Month(targetDate), targetDay))
End Function

' ============================================================================
' SECTION 07 - STATISTICS AND AUDIT HELPERS
' ============================================================================

Private Function CountAuditTenorRows(ByVal ws As Worksheet, _
                                     ByVal lastRow As Long, _
                                     ByVal tenorText As String) As Long
    Dim rowNumber As Long

    For rowNumber = 5 To lastRow
        If CStr(ws.Cells(rowNumber, 2).Value2) = tenorText Then
            CountAuditTenorRows = CountAuditTenorRows + 1
        End If
    Next rowNumber
End Function

Private Function AverageAuditStartRate(ByVal ws As Worksheet, _
                                       ByVal lastRow As Long, _
                                       ByVal tenorText As String) As Double
    Dim rowNumber As Long
    Dim totalValue As Double
    Dim observationCount As Long

    For rowNumber = 5 To lastRow
        If CStr(ws.Cells(rowNumber, 2).Value2) = tenorText Then
            totalValue = totalValue + CDbl(ws.Cells(rowNumber, 6).Value2)
            observationCount = observationCount + 1
        End If
    Next rowNumber

    If observationCount > 0 Then
        AverageAuditStartRate = totalValue / observationCount
    End If
End Function

Private Function AuditAverage(ByRef values() As Double) As Double
    Dim indexValue As Long

    For indexValue = LBound(values) To UBound(values)
        AuditAverage = AuditAverage + values(indexValue)
    Next indexValue

    AuditAverage = AuditAverage / _
        (UBound(values) - LBound(values) + 1)
End Function

Private Function AuditSampleStdDev(ByRef values() As Double) As Double
    Dim valueCount As Long
    Dim indexValue As Long
    Dim averageValue As Double
    Dim varianceSum As Double

    valueCount = UBound(values) - LBound(values) + 1

    If valueCount < 2 Then Exit Function

    averageValue = AuditAverage(values)

    For indexValue = LBound(values) To UBound(values)
        varianceSum = varianceSum + _
            (values(indexValue) - averageValue) ^ 2
    Next indexValue

    AuditSampleStdDev = Sqr(varianceSum / (valueCount - 1))
End Function

Private Function AuditPercentileSorted( _
    ByRef sortedValues() As Double, _
    ByVal percentileValue As Double) As Double

    Dim positionValue As Double
    Dim lowerIndex As Long
    Dim upperIndex As Long
    Dim lowerBoundValue As Long

    lowerBoundValue = LBound(sortedValues)
    positionValue = _
        (UBound(sortedValues) - lowerBoundValue) * percentileValue

    lowerIndex = lowerBoundValue + Int(positionValue)
    upperIndex = lowerBoundValue + _
        Application.WorksheetFunction.RoundUp(positionValue, 0)

    If lowerIndex = upperIndex Then
        AuditPercentileSorted = sortedValues(lowerIndex)
    Else
        AuditPercentileSorted = sortedValues(lowerIndex) + _
            (positionValue - Int(positionValue)) * _
            (sortedValues(upperIndex) - sortedValues(lowerIndex))
    End If
End Function

Private Sub SortAuditDoubleArray(ByRef values() As Double)
    QuickSortAuditDouble values, LBound(values), UBound(values)
End Sub

Private Sub QuickSortAuditDouble(ByRef values() As Double, _
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
        QuickSortAuditDouble values, firstIndex, highIndex
    End If

    If lowIndex < lastIndex Then
        QuickSortAuditDouble values, lowIndex, lastIndex
    End If
End Sub

' ============================================================================
' SECTION 08 - GENERAL HELPERS
' ============================================================================

Private Function AuditTenorName(ByVal tenorIndex As Long) As String
    Select Case tenorIndex
        Case 1
            AuditTenorName = "ON"
        Case 2
            AuditTenorName = "1M"
        Case 3
            AuditTenorName = "2M"
        Case 4
            AuditTenorName = "3M"
        Case 5
            AuditTenorName = "6M"
        Case Else
            Err.Raise vbObjectError + 300, , "Invalid tenor index."
    End Select
End Function

Private Function AuditTenorMonths(ByVal tenorIndex As Long) As Long
    Select Case tenorIndex
        Case 1
            AuditTenorMonths = 0
        Case 2
            AuditTenorMonths = 1
        Case 3
            AuditTenorMonths = 2
        Case 4
            AuditTenorMonths = 3
        Case 5
            AuditTenorMonths = 6
        Case Else
            Err.Raise vbObjectError + 301, , "Invalid tenor index."
    End Select
End Function

Private Function AuditDatabaseDisplayName( _
    ByVal sheetName As String) As String

    If sheetName = "Rollover_Path_DB" Then
        AuditDatabaseDisplayName = "Sequential Path"
    Else
        AuditDatabaseDisplayName = "All Eligible Starts"
    End If
End Function

Private Function AuditSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    AuditSheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function FindAuditHeaderRow(ByVal ws As Worksheet) As Long
    Dim rowNumber As Long

    For rowNumber = 1 To 20
        If FindAuditHeaderColumn(ws, rowNumber, "DATE") > 0 And _
           FindAuditHeaderColumn(ws, rowNumber, "ON") > 0 Then

            FindAuditHeaderRow = rowNumber
            Exit Function
        End If
    Next rowNumber
End Function

Private Function FindAuditHeaderColumn(ByVal ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal headerText As String) As Long
    Dim columnNumber As Long
    Dim cellText As String

    For columnNumber = 1 To 30
        cellText = UCase$(Trim$(CStr( _
            ws.Cells(headerRow, columnNumber).Value2)))

        If cellText = UCase$(headerText) Then
            FindAuditHeaderColumn = columnNumber
            Exit Function
        End If
    Next columnNumber
End Function

Private Function AuditCollectionToMatrix( _
    ByVal rows As Collection, _
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

    AuditCollectionToMatrix = matrix
End Function

Private Sub QuickSortAuditCurve(ByVal firstIndex As Long, _
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
        QuickSortAuditCurve firstIndex, highIndex
    End If

    If lowIndex < lastIndex Then
        QuickSortAuditCurve lowIndex, lastIndex
    End If
End Sub

Private Function AuditMinimumLong(ByVal firstValue As Long, _
                                  ByVal secondValue As Long) As Long
    If firstValue < secondValue Then
        AuditMinimumLong = firstValue
    Else
        AuditMinimumLong = secondValue
    End If
End Function

Private Sub WriteAuditRow(ByVal firstCell As Range, _
                          ByVal values As Variant)
    Dim matrix() As Variant
    Dim itemCount As Long
    Dim indexValue As Long

    itemCount = UBound(values) - LBound(values) + 1
    ReDim matrix(1 To 1, 1 To itemCount)

    For indexValue = 1 To itemCount
        matrix(1, indexValue) = _
            values(LBound(values) + indexValue - 1)
    Next indexValue

    firstCell.Resize(1, itemCount).Value2 = matrix
End Sub

Private Sub StyleAuditTitle(ByVal target As Range)
    With target
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StyleAuditHeader(ByVal target As Range)
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
