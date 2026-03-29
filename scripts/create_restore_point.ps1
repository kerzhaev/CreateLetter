param(
    [Parameter(Mandatory = $false)]
    [string]$Label = "manual",

    [Parameter(Mandatory = $false)]
    [string]$BranchName = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

if ([string]::IsNullOrWhiteSpace($BranchName)) {
    try {
        $BranchName = (git -C $repoRoot branch --show-current).Trim()
    }
    catch {
        $BranchName = "unknown-branch"
    }
}

$safeLabel = ($Label -replace "[^A-Za-z0-9._-]", "-").Trim("-")
if ([string]::IsNullOrWhiteSpace($safeLabel)) {
    $safeLabel = "manual"
}

$safeBranch = ($BranchName -replace "[^A-Za-z0-9._-]", "-").Trim("-")
if ([string]::IsNullOrWhiteSpace($safeBranch)) {
    $safeBranch = "unknown-branch"
}

$restoreDir = Join-Path $repoRoot ("filesarchive\restore-point-{0}-{1}" -f $safeLabel, $timestamp)
New-Item -ItemType Directory -Path $restoreDir -Force | Out-Null

$workbookPath = Join-Path $repoRoot "CreateLetter.xlsm"
$modulesPath = Join-Path $repoRoot "CreateLetter.xlsm.modules"
$documentModulesPath = Join-Path $repoRoot "CreateLetter.xlsm.document-modules"

if (-not (Test-Path $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

if (-not (Test-Path $modulesPath)) {
    throw "Modules directory not found: $modulesPath"
}

if (-not (Test-Path $documentModulesPath)) {
    throw "Document-modules directory not found: $documentModulesPath"
}

Copy-Item $workbookPath -Destination (Join-Path $restoreDir "CreateLetter.xlsm")
Copy-Item $modulesPath -Destination (Join-Path $restoreDir "CreateLetter.xlsm.modules") -Recurse
Copy-Item $documentModulesPath -Destination (Join-Path $restoreDir "CreateLetter.xlsm.document-modules") -Recurse

$readmePath = Join-Path $restoreDir "README.txt"
@(
    "CreateLetter restore point"
    "Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "Label: $safeLabel"
    "Branch: $safeBranch"
    "Rollback inputs:"
    "  - Workbook snapshot: CreateLetter.xlsm"
    "  - Module snapshot: CreateLetter.xlsm.modules"
    "  - Document-module snapshot: CreateLetter.xlsm.document-modules"
    ""
    "Usage:"
    "  Restore the workbook file and module folder from this directory before retrying or rolling back a feature stage."
) | Set-Content -Path $readmePath

Write-Output $restoreDir
