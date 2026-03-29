param(
    [Parameter(Mandatory = $false)]
    [string]$Label = "manual",

    [Parameter(Mandatory = $false)]
    [string]$WorkbookName = "Workbook.xlsm",

    [Parameter(Mandatory = $false)]
    [string]$ModulesDirectoryName = "Workbook.xlsm.modules"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$safeLabel = ($Label -replace "[^A-Za-z0-9._-]", "-").Trim("-")

if ([string]::IsNullOrWhiteSpace($safeLabel)) {
    $safeLabel = "manual"
}

$restoreDir = Join-Path $repoRoot ("filesarchive\restore-point-{0}-{1}" -f $safeLabel, $timestamp)
New-Item -ItemType Directory -Path $restoreDir -Force | Out-Null

$workbookPath = Join-Path $repoRoot $WorkbookName
$modulesPath = Join-Path $repoRoot $ModulesDirectoryName

if (-not (Test-Path $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

if (-not (Test-Path $modulesPath)) {
    throw "Modules directory not found: $modulesPath"
}

Copy-Item $workbookPath -Destination (Join-Path $restoreDir $WorkbookName)
Copy-Item $modulesPath -Destination (Join-Path $restoreDir $ModulesDirectoryName) -Recurse

$readmePath = Join-Path $restoreDir "README.txt"
@(
    "Excel VBA restore point"
    "Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "Label: $safeLabel"
    "Workbook: $WorkbookName"
    "Modules: $ModulesDirectoryName"
) | Set-Content -Path $readmePath

Write-Output $restoreDir
