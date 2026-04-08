param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\Workbook.xlsm",

    [Parameter(Mandatory = $false)]
    [string]$ModulesPath = ".\Workbook.xlsm.modules",

    [Parameter(Mandatory = $false)]
    [string]$DocumentModulesPath = ".\Workbook.xlsm.document-modules"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "== Close Excel ==" -ForegroundColor Cyan
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Write-Host "== Sync VBA -> Workbook ==" -ForegroundColor Cyan
python .\scripts\sync_vba_from_modules.py $WorkbookPath $ModulesPath $DocumentModulesPath
if ($LASTEXITCODE -ne 0) {
    throw "VBA sync failed."
}

Write-Host "== Apply Ribbon XML ==" -ForegroundColor Cyan
python .\scripts\apply_custom_ui.py $WorkbookPath
if ($LASTEXITCODE -ne 0) {
    throw "Ribbon apply failed."
}

Write-Host "== Smoke Test ==" -ForegroundColor Cyan
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1 -WorkbookPath $WorkbookPath

if ($LASTEXITCODE -ne 0) {
    throw "Smoke test failed."
}

Write-Host "Workbook repair completed successfully." -ForegroundColor Green
