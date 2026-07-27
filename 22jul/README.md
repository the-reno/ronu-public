# SOFR VBA Model — 22jul

This folder separates the process into two VBA steps:

1. Build the Excel input template.
2. Run the simulation after the user enters the curve.

## Repository files

### `01_Build_SOFR_Template.bas`
Creates the input workbook structure:

- `Read Me`
- `Curve`
- `Config`
- `FOMC`

Main macro:

```text
BuildSOFRTemplate
```

### `02_Run_SOFR_Simulation.bas`
Runs the complete analysis and recreates:

- `Executive`
- `Historical Curve`
- `Premium vs Realized`
- `Starting Periods`
- `Risk Return`
- `Efficient Frontier`
- `Next FOMC`
- `Sources`

Main macro:

```text
RunSOFRSimulation
```

The simulation module is stored in `simulation_parts` to preserve the full VBA source in GitHub. Run this once in PowerShell to assemble it:

```powershell
powershell -ExecutionPolicy Bypass -File .\assemble_simulation.ps1
```

This creates `02_Run_SOFR_Simulation.bas` in the same folder.

## Excel setup

1. Download or clone the `22jul` folder.
2. Run `assemble_simulation.ps1` once.
3. Open a blank Excel workbook.
4. Save it as **Excel Macro-Enabled Workbook (`.xlsm`)**.
5. Press `Alt + F11`.
6. Select **File > Import File**.
7. Import:
   - `01_Build_SOFR_Template.bas`
   - `02_Run_SOFR_Simulation.bas`
8. Run `BuildSOFRTemplate`.

## Add the data

Paste the curve into the `Curve` sheet with these exact headers:

```text
Date | ON | 1M | 2M | 3M | 6M
```

Rates may be entered as either:

```text
3.5500
```

or:

```text
0.0355
```

Review the optional parameters in `Config` and the policy history in `FOMC`.

## Run the simulation

Run:

```text
RunSOFRSimulation
```

or click **Run SOFR Simulation** on the `Read Me` sheet.

## Normal workflow

```text
Run BuildSOFRTemplate once
        ↓
Paste or update Curve data
        ↓
Review Config and FOMC
        ↓
Run RunSOFRSimulation whenever the data changes
```

The Excel model is fully local. Python, internet access, external files, and Excel add-ins are not required after the VBA modules are imported.
