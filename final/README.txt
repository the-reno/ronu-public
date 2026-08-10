FINAL CASH TENOR ANALYSIS — VBA PACKAGE
=======================================

Required workbook inputs
------------------------
1. Worksheet named:
   Curve

2. Curve row 1 must include:
   Date | ON | 1M | 3M | 6M

   Extra columns are allowed.
   Curve can contain more or fewer rows.
   Dates must be strictly ascending.

3. Public R_CURVE module in the same VBA project.

No other pre-existing sheet, add-in or data dependency is required.

Files
-----
00_Master.txt
01_Slide1_Historical.txt
02_Slide2_Direction.txt
03_Slide3_Regret.txt
04_Slide4_Robust.txt
05_Slide5_Execution.txt
06_Slide_Wording.txt

Import / use
------------
1. Paste each VBA .txt file 00 through 05 into a separate standard VBA module.
2. Keep the supplied R_CURVE module in the same workbook.
3. Run:
   Debug > Compile VBAProject
4. Run macro:
   FINAL_BuildAll

The master creates
------------------
Final_Control
Final_Input
F_Curve_Normalized
F_Curve_Inverted
F_Core_Returns
F_Portfolios
S1_Historical
S2_Direction
S3_Regret
S4_Robust
S5_Execution

Analysis window
---------------
Start:
First available Curve date on/after 01-Jan-2022.

End:
Last available Curve date.

Historical return observations:
First available date in each calendar month, only where a complete common 6M horizon fits inside the curve.

Rate units
----------
The master leaves Curve unchanged.

It creates F_Curve_Normalized and detects whether rates were supplied as:
- decimal rates, e.g. 0.054 = 5.40%
- percent points, e.g. 5.40 = 5.40%

R_CURVE is run only on the normalized working curve.

Directional stress
------------------
The inverted curve preserves the exact historical daily changes and cross-tenor movement, but reverses the sign of those changes.

Default starting ON:
5.50%

Default inversion magnitude:
1.0x

Both can be reviewed in Final_Control.

Return engine
-------------
All official strategy returns use R_CURVE.

Common explicit maturity:
Start + 6 calendar months, snapped backward using adj=0.

Tenors:
ON, 1M, 3M, 6M.

Optimization
------------
Portfolio grid:
Long only.
10% increments.
Weights sum to 100%.
286 portfolios.

Observation regret:
Best single-tenor return for that scenario/start
minus portfolio return.

Historical average regret:
Average regret across historical monthly starts.

Inverted average regret:
Average regret across inverted monthly starts.

Robust allocation:
Minimize the larger of the two directional average regrets.

Tie-breakers:
1. Lower mean directional regret.
2. Shorter weighted tenor, preserving more liquidity.

Slide 5 input
-------------
The master creates Final_Input.

Enter:
B3  Credit rating
B4  Global coverage
B5  Additional proof point

Paste actual premium history starting at row 10:
Date | 3M Premium (bp) | 6M Premium (bp)

Only observations from Jan-2025 onward are plotted.

Presentation discipline
-----------------------
Slide 1:
What actually happened?

Slide 2:
What changes when rate direction changes?

Slide 3:
What is the risk we should manage?

Slide 4:
Which allocation best protects against being wrong?

Slide 5:
How is the framework executed consistently?

The historical and inverted paths are directional stress tests, not forecasts.
