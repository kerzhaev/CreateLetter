param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\Workbook.xlsm",

    [Parameter(Mandatory = $false)]
    [string[]]$RequiredSheets = @("Addresses", "Letters", "Settings"),

    [Parameter(Mandatory = $false)]
    [string[]]$RequiredTables = @(),

    [Parameter(Mandatory = $false)]
    [switch]$RequireRibbonCustomization
)

$ErrorActionPreference = "Stop"

function Add-Result {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [string]$Name,
        [string]$Status,
        [string]$Details
    )

    $Results.Add([PSCustomObject]@{
        Name = $Name
        Status = $Status
        Details = $Details
    }) | Out-Null
}

$resolvedWorkbookPath = Resolve-Path $WorkbookPath
$results = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$workbook = $null
$failed = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath.Path, $false, $true)

    Add-Result -Results $results -Name "WorkbookOpen" -Status "PASS" -Details "Workbook opened in read-only mode."

    foreach ($sheetName in $RequiredSheets) {
        try {
            $null = $workbook.Worksheets.Item($sheetName)
            Add-Result -Results $results -Name ("Worksheet:" + $sheetName) -Status "PASS" -Details "Worksheet is present."
        }
        catch {
            Add-Result -Results $results -Name ("Worksheet:" + $sheetName) -Status "FAIL" -Details "Worksheet is missing."
            $failed = $true
        }
    }

    foreach ($tableName in $RequiredTables) {
        $found = $false
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $ws = $workbook.Worksheets.Item($i)
            for ($j = 1; $j -le $ws.ListObjects.Count; $j++) {
                if ([string]$ws.ListObjects.Item($j).Name -eq $tableName) {
                    $found = $true
                    break
                }
            }
            if ($found) { break }
        }

        if ($found) {
            Add-Result -Results $results -Name ("StructuredTable:" + $tableName) -Status "PASS" -Details "Structured table is present."
        }
        else {
            Add-Result -Results $results -Name ("StructuredTable:" + $tableName) -Status "FAIL" -Details "Structured table is missing."
            $failed = $true
        }
    }

    if ($RequireRibbonCustomization) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $tempCopy = Join-Path $env:TEMP ("excel-vba-ribbon-check-" + [guid]::NewGuid().ToString() + ".xlsm")
        Copy-Item -LiteralPath $resolvedWorkbookPath.Path -Destination $tempCopy -Force

        try {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($tempCopy)
            $customUiEntry = $archive.GetEntry("customUI/customUI.xml")
            $rootRelsEntry = $archive.GetEntry("_rels/.rels")
            $hasRibbonRelationship = $false

            if ($null -ne $rootRelsEntry) {
                $stream = $rootRelsEntry.Open()
                $reader = New-Object System.IO.StreamReader($stream)
                try {
                    $relsText = $reader.ReadToEnd()
                    $hasRibbonRelationship = $relsText -like "*http://schemas.microsoft.com/office/2006/relationships/ui/extensibility*"
                }
                finally {
                    $reader.Dispose()
                }
            }

            if ($null -ne $customUiEntry -and $hasRibbonRelationship) {
                Add-Result -Results $results -Name "RibbonCustomization" -Status "PASS" -Details "customUI package markup is present."
            }
            else {
                Add-Result -Results $results -Name "RibbonCustomization" -Status "FAIL" -Details "customUI package markup is missing."
                $failed = $true
            }
        }
        finally {
            if ($null -ne $archive) { $archive.Dispose() }
            Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
        }
    }
}
catch {
    Add-Result -Results $results -Name "SmokeHarness" -Status "FAIL" -Details $_.Exception.Message
    $failed = $true
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}

$results | Format-Table -AutoSize

if ($failed) {
    exit 1
}
