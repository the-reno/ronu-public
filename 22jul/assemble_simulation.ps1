$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$partsPath = Join-Path $root "simulation_parts"
$outputPath = Join-Path $root "02_Run_SOFR_Simulation.bas"

$parts = Get-ChildItem -Path $partsPath -Filter "part*.txt" | Sort-Object Name
if ($parts.Count -eq 0) {
    throw "No simulation code parts were found in $partsPath"
}

$builder = New-Object System.Text.StringBuilder
foreach ($part in $parts) {
    [void]$builder.Append([System.IO.File]::ReadAllText($part.FullName))
}

[System.IO.File]::WriteAllText(
    $outputPath,
    $builder.ToString(),
    (New-Object System.Text.UTF8Encoding($true))
)

Write-Host "Created: $outputPath"
Write-Host "Import 01_Build_SOFR_Template.bas and 02_Run_SOFR_Simulation.bas into a blank .xlsm workbook."
