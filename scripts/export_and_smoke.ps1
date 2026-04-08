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

Write-Host "== Export Workbook VBA -> Source Files ==" -ForegroundColor Cyan
python .\scripts\export_vba_to_modules.py $WorkbookPath $ModulesPath $DocumentModulesPath
if ($LASTEXITCODE -ne 0) {
    throw "VBA export failed."
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

Write-Host "Export and smoke completed successfully." -ForegroundColor Green
