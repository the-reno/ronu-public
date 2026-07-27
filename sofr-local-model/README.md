# SOFR Local Model

This folder contains a self-contained installer for the SOFR curve, return, risk and FOMC analysis model.

The model runs locally in Visual Studio, Visual Studio Code, or a terminal on Windows, macOS, and Linux. It does not need an internet connection after the repository is cloned.

## Install in Visual Studio

1. Open this folder in Visual Studio.
2. Open **View > Terminal**.
3. Extract the package:

```powershell
py setup_sofr_local_model.py
```

4. Create and activate a virtual environment:

```powershell
py -m venv .venv
.venv\Scripts\activate
```

5. Install dependencies:

```powershell
py -m pip install -r requirements_sofr_local.txt
```

6. Create a workbook template:

```powershell
py sofr_local_model.py --create-template SOFR_Local_Model.xlsx
```

7. Open `SOFR_Local_Model.xlsx`, paste the curve into the **Curve** sheet, and save and close the workbook.

Required columns:

```text
Date | ON | 1M | 2M | 3M | 6M
```

Rates may be entered as either `3.5500` or `0.0355`.

8. Run the analysis:

```powershell
py sofr_local_model.py --input SOFR_Local_Model.xlsx
```

To update the same workbook and create a timestamped backup:

```powershell
py sofr_local_model.py --input SOFR_Local_Model.xlsx --in-place
```

## Input sheets

- **Curve**: user-provided historical curve.
- **Config**: analysis dates, regime thresholds, start periods, target range, and next-meeting dates.
- **FOMC**: meeting date, decision, action in basis points, and SEP indicator.

## Output sheets

- Executive
- Historical Curve
- Premium vs Realized
- Starting Periods
- Risk Return
- Efficient Frontier
- Next FOMC
- Sources

The analysis period is determined dynamically from the data and configuration. The curve can cover a different period each time the model is run.
