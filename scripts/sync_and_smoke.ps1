param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\CreateLetter.xlsm",

    [Parameter(Mandatory = $false)]
    [string]$ModulesPath = ".\CreateLetter.xlsm.modules",

    [Parameter(Mandatory = $false)]
    [string]$DocumentModulesPath = ".\CreateLetter.xlsm.document-modules"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "== Sync VBA -> Workbook ==" -ForegroundColor Cyan
python .\scripts\sync_vba_from_modules.py $WorkbookPath $ModulesPath $DocumentModulesPath
if ($LASTEXITCODE -ne 0) {
    throw "VBA sync failed."
}

Write-Host "== Smoke Test ==" -ForegroundColor Cyan
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1 `
    -WorkbookPath $WorkbookPath `
    -RequireLocalizationModule `
    -RequireStructuredTables `
    -RequireLocalizationSheet `
    -RequireRibbonCustomization `
    -RequireAddressGroupColumn

if ($LASTEXITCODE -ne 0) {
    throw "Smoke test failed."
}

Write-Host "Sync and smoke completed successfully." -ForegroundColor Green
