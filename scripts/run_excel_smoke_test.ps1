param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\CreateLetter.xlsm",

    [Parameter(Mandatory = $false)]
    [switch]$RequireLocalizationModule,

    [Parameter(Mandatory = $false)]
    [switch]$RequireStructuredTables,

    [Parameter(Mandatory = $false)]
    [switch]$RequireLocalizationSheet,

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

function Get-WorksheetTableNames {
    param(
        [object]$Worksheet
    )

    $tableNames = New-Object 'System.Collections.Generic.List[string]'

    for ($i = 1; $i -le $Worksheet.ListObjects.Count; $i++) {
        $tableNames.Add([string]$Worksheet.ListObjects.Item($i).Name) | Out-Null
    }

    return $tableNames
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

    if ($RequireLocalizationSheet) {
        $requiredSheets += @{
            LogicalName = "Localization"
            Variants = @("Localization")
        }
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

    $structuredTableRequirements = @(
        @{ Sheet = "Addresses"; Table = "tblAddresses" },
        @{ Sheet = "Letters"; Table = "tblLetters" },
        @{ Sheet = "Settings"; Table = "tblLetterTexts" }
    )

    foreach ($tableRequirement in $structuredTableRequirements) {
        try {
            $sheet = $workbook.Worksheets.Item($tableRequirement.Sheet)
            $tableNames = Get-WorksheetTableNames -Worksheet $sheet
            if ($tableNames.Contains($tableRequirement.Table)) {
                Add-Result -Results $results -Name ("StructuredTable:" + $tableRequirement.Sheet) -Status "PASS" -Details ("Structured table '" + $tableRequirement.Table + "' is present.")
            }
            elseif ($RequireStructuredTables) {
                Add-Result -Results $results -Name ("StructuredTable:" + $tableRequirement.Sheet) -Status "FAIL" -Details ("Expected structured table '" + $tableRequirement.Table + "'. Found: " + (($tableNames | Select-Object -First 20) -join ", "))
                $failed = $true
            }
            else {
                Add-Result -Results $results -Name ("StructuredTable:" + $tableRequirement.Sheet) -Status "WARN" -Details ("Structured table '" + $tableRequirement.Table + "' is missing. Found: " + (($tableNames | Select-Object -First 20) -join ", "))
            }
        }
        catch {
            Add-Result -Results $results -Name ("StructuredTable:" + $tableRequirement.Sheet) -Status "FAIL" -Details $_.Exception.Message
            $failed = $true
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
        $aliasModule = $workbook.VBProject.VBComponents.Item("ModuleMain").CodeModule
        $initializeModule = $workbook.VBProject.VBComponents.Item("mdlInicialize").CodeModule
        $moduleMainText = [string]$aliasModule.Lines(1, $aliasModule.CountOfLines)
        $initializeModuleText = [string]$initializeModule.Lines(1, $initializeModule.CountOfLines)
        $requiredAliasNames = @(
            "Public Sub ShowLetterCreator()",
            "Public Sub ShowLetterCreatorDeferred()",
            "Public Sub BootstrapWorkbookSheets()",
            "Public Sub ResetWorkbookSheets()"
        )

        $missingAliases = New-Object 'System.Collections.Generic.List[string]'
        foreach ($aliasName in $requiredAliasNames) {
            if (($moduleMainText -notlike ("*" + $aliasName + "*")) -and ($initializeModuleText -notlike ("*" + $aliasName + "*"))) {
                $missingAliases.Add($aliasName) | Out-Null
            }
        }

        if ($missingAliases.Count -eq 0) {
            Add-Result -Results $results -Name "EnglishAliases" -Status "PASS" -Details "English-safe public aliases are present in ModuleMain and mdlInicialize."
        }
        else {
            Add-Result -Results $results -Name "EnglishAliases" -Status "FAIL" -Details ("Missing aliases: " + ($missingAliases -join ", "))
            $failed = $true
        }
    }
    catch {
        Add-Result -Results $results -Name "EnglishAliases" -Status "FAIL" -Details ("Alias inspection failed: " + $_.Exception.Message)
        $failed = $true
    }

    try {
        $documentTypeLabel = [string]$excel.Run("'" + $workbook.Name + "'!GetDocumentTypeDisplayLabel", "confirmed_documents")
        if ([string]::IsNullOrWhiteSpace($documentTypeLabel) -or $documentTypeLabel -eq "confirmed_documents") {
            Add-Result -Results $results -Name "DocumentTypeDisplay" -Status "FAIL" -Details "Internal document type key leaked instead of a display label."
            $failed = $true
        }
        else {
            Add-Result -Results $results -Name "DocumentTypeDisplay" -Status "PASS" -Details ("Display label resolved to: " + $documentTypeLabel)
        }
    }
    catch {
        Add-Result -Results $results -Name "DocumentTypeDisplay" -Status "FAIL" -Details ("Display label lookup failed: " + $_.Exception.Message)
        $failed = $true
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
