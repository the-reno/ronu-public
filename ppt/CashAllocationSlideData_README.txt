Cash Allocation Slide Data VBA
================================

File
----
CashAllocationSlideData.bas

Purpose
-------
Prepares the chart-ready data used in the first three cash-allocation slides.
It leaves the Curve sheet unchanged and creates or refreshes dedicated output
sheets.

Required Curve sheet
--------------------
The workbook must contain a worksheet named Curve. The module searches rows
1:25 for these required headers:

Date | ON | 1M | 3M | 6M

Additional columns such as 2M are allowed and ignored. Dates must be strictly
ascending. Rates can be stored as Excel decimals (3.55% = 0.0355) or percentage
points (3.55% = 3.55); the macro detects the scale for the full data set.

Installation
------------
1. Save the workbook as an Excel Macro-Enabled Workbook (*.xlsm).
2. Press Alt+F11 to open the Visual Basic Editor.
3. In the editor, choose File > Import File.
4. Select CashAllocationSlideData.bas.
5. Return to Excel and run BuildCashAllocationSlideData from Alt+F8.

Generated sheets
----------------
Slide_Setup
    User inputs and methodology. Blue/yellow cells B4:B7 are editable.

Slide2_6M
    Daily starts with inception 6M-ON premium, rolled ON return, locked 6M
    return, realized 6M excess return and winner. Use columns E and H for the
    slide-2 time series and scatter plot.

Slide3_Paths
    Realized-direction and inverted-direction ON/1M/3M/6M rates, normalized to
    a common level and automatically kept above the minimum plotted rate.

Slide3_Returns
    Six-month holding returns and excess returns for the first available curve
    date in each calendar month under both directional paths.

Slide3_Summary
    Chart-ready average returns, excess returns versus ON and winning-start
    counts by tenor.

Slide_Checks
    Header, observation-count, rate-scale, scenario-floor and output checks.

Important conventions
---------------------
- Common six-month maturity.
- ACT/360 simple accrual within each holding interval.
- ON resets on each available curve date.
- 1M and 3M reset at tenor dates, snapped backward to the latest available
  curve date.
- 6M is locked at the decision date.
- The inverted path reverses observed curve changes; it is a stress, not a
  forecast.
- Rerunning the macro overwrites only the generated Slide_* sheets. It does not
  edit the Curve sheet.
