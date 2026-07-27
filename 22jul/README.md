# SOFR VBA Model — 22jul

This folder separates the Excel process into two VBA steps:

1. Build the input template.
2. Run the simulation after the user enters the curve.

## Files

### `01_Build_SOFR_Template.bas`

Creates and formats the input sheets:

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

The complete simulation source is stored in `simulation_parts`. Create the importable `.bas` file once by double-clicking:

```text
BUILD_SIMULATION_MODULE.bat
```

The batch file runs `assemble_simulation.ps1` and creates:

```text
02_Run_SOFR_Simulation.bas
```

## Initial Excel setup

1. Download or clone the `22jul` folder.
2. Double-click `BUILD_SIMULATION_MODULE.bat`.
3. Open a blank Excel workbook.
4. Save it as **Excel Macro-Enabled Workbook (`.xlsm`)**.
5. Press `Alt + F11`.
6. Select **File > Import File**.
7. Import:
   - `01_Build_SOFR_Template.bas`
   - `02_Run_SOFR_Simulation.bas`
8. Run:

```text
BuildSOFRTemplate
```

The first macro creates the spreadsheet structure and a **Run SOFR Simulation** button.

## User data

Paste the curve into the `Curve` sheet with these exact headers:

```text
Date | ON | 1M | 2M | 3M | 6M
```

Rates may be entered as either percentage points:

```text
3.5500
```

or Excel decimal rates:

```text
0.0355
```

Review the optional parameters in `Config` and the policy history in `FOMC`.

## Run the analysis

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

The model is fully local. Python, internet access, external files, Excel add-ins and external references are not required after the two VBA modules are imported.
