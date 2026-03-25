param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\CreateLetter.xlsm",

    [Parameter(Mandatory = $false)]
    [switch]$RequireLocalizationModule,

    [Parameter(Mandatory = $false)]
    [switch]$AllowLegacyRussianSheetNames
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

function Test-WorksheetVariants {
    param(
        [object]$Workbook,
        [string[]]$Variants
    )

    foreach ($sheetName in $Variants) {
        try {
            $sheet = $Workbook.Worksheets.Item($sheetName)
            if ($null -ne $sheet) {
                return $sheetName
            }
        }
        catch {
        }
    }

    return $null
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

    $requiredSheets = @(
        @{
            LogicalName = "Addresses"
            Variants = @("Addresses")
        },
        @{
            LogicalName = "Letters"
            Variants = @("Letters")
        },
        @{
            LogicalName = "Settings"
            Variants = @("Settings")
        }
    )

    if ($AllowLegacyRussianSheetNames) {
        $requiredSheets[0].Variants += "Адреса"
        $requiredSheets[1].Variants += "Письма"
        $requiredSheets[2].Variants += "Настройки"
    }

    foreach ($sheetRequirement in $requiredSheets) {
        $resolvedSheetName = Test-WorksheetVariants -Workbook $workbook -Variants $sheetRequirement.Variants

        if ($null -eq $resolvedSheetName) {
            Add-Result -Results $results -Name ("Worksheet:" + $sheetRequirement.LogicalName) -Status "FAIL" -Details ("Required worksheet not found. Checked: " + ($sheetRequirement.Variants -join ", "))
            $failed = $true
        }
        else {
            Add-Result -Results $results -Name ("Worksheet:" + $sheetRequirement.LogicalName) -Status "PASS" -Details ("Worksheet is present as '" + $resolvedSheetName + "'.")
        }
    }

    $formatPhoneResult = $excel.Run("'" + $workbook.Name + "'!FormatPhoneNumber", "89281234567")
    if ($formatPhoneResult -eq "8-928-123-45-67") {
        Add-Result -Results $results -Name "FormatPhoneNumber" -Status "PASS" -Details "Returned expected normalized phone."
    }
    else {
        Add-Result -Results $results -Name "FormatPhoneNumber" -Status "FAIL" -Details ("Unexpected result: " + [string]$formatPhoneResult)
        $failed = $true
    }

    $validPhoneResult = $excel.Run("'" + $workbook.Name + "'!IsPhoneNumberValid", "8-928-123-45-67")
    if ([bool]$validPhoneResult) {
        Add-Result -Results $results -Name "IsPhoneNumberValid" -Status "PASS" -Details "Accepted expected valid phone."
    }
    else {
        Add-Result -Results $results -Name "IsPhoneNumberValid" -Status "FAIL" -Details "Expected True for a valid phone."
        $failed = $true
    }

    $formattedDate = [string]$excel.Run("'" + $workbook.Name + "'!FormatLetterDate", "25.03.2026")
    if ([string]::IsNullOrWhiteSpace($formattedDate)) {
        Add-Result -Results $results -Name "FormatLetterDate" -Status "FAIL" -Details "Returned an empty formatted date."
        $failed = $true
    }
    else {
        Add-Result -Results $results -Name "FormatLetterDate" -Status "PASS" -Details ("Returned: " + $formattedDate)
    }

    try {
        $localizationStats = [string]$excel.Run("'" + $workbook.Name + "'!GetLocalizationStats")
        Add-Result -Results $results -Name "LocalizationModule" -Status "PASS" -Details $localizationStats

        $cancelText = [string]$excel.Run("'" + $workbook.Name + "'!T", "common.cancel", "fallback")
        if ([string]::IsNullOrWhiteSpace($cancelText) -or $cancelText -eq "fallback") {
            Add-Result -Results $results -Name "LocalizationLookup" -Status "FAIL" -Details "Localization lookup returned an empty value or fallback."
            $failed = $true
        }
        else {
            Add-Result -Results $results -Name "LocalizationLookup" -Status "PASS" -Details ("Lookup returned a non-fallback value with length " + $cancelText.Length + ".")
        }
    }
    catch {
        if ($RequireLocalizationModule) {
            Add-Result -Results $results -Name "LocalizationModule" -Status "FAIL" -Details "Localization module is not available in the workbook. Import modules first."
            $failed = $true
        }
        else {
            Add-Result -Results $results -Name "LocalizationModule" -Status "WARN" -Details "Localization module not found in workbook yet. Import modules before validating localization stages."
        }
    }
}
catch {
    Add-Result -Results $results -Name "SmokeHarness" -Status "FAIL" -Details $_.Exception.Message
    $failed = $true
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($null -ne $excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$results | Format-Table -AutoSize | Out-String | Write-Output

if ($failed) {
    exit 1
}

exit 0
