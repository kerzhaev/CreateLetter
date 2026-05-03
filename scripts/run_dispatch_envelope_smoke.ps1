param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = ".\CreateLetter.xlsm",

    [Parameter(Mandatory = $false)]
    [switch]$KeepTemp
)

$ErrorActionPreference = "Stop"

function Add-SmokeResult {
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

function Get-ExcelOpenPath {
    param(
        [string]$Path
    )

    try {
        return ([System.Uri](Resolve-Path $Path).Path).AbsoluteUri
    } catch {
        return $Path
    }
}

function Get-Table {
    param(
        [object]$Workbook,
        [string]$SheetName,
        [string]$TableName
    )

    $sheet = $Workbook.Worksheets.Item($SheetName)
    return $sheet.ListObjects.Item($TableName)
}

function Clear-TableRows {
    param(
        [object]$Table
    )

    if ($null -ne $Table.DataBodyRange) {
        $Table.DataBodyRange.Delete() | Out-Null
    }
}

function Set-TableRowValues {
    param(
        [object]$Row,
        [hashtable]$Values
    )

    foreach ($key in $Values.Keys) {
        $columnIndex = $Row.Parent.ListColumns.Item($key).Index
        $Row.Range.Cells.Item(1, $columnIndex).Value2 = $Values[$key]
    }
}

function Add-TableRow {
    param(
        [object]$Table,
        [hashtable]$Values
    )

    $row = $Table.ListRows.Add()
    Set-TableRowValues -Row $row -Values $Values
}

function Get-ShapeText {
    param(
        [object]$Shape
    )

    try {
        return [string]$Shape.TextFrame.Characters().Text
    }
    catch {
        return ""
    }
}

function Test-EnvelopeSheet {
    param(
        [object]$Workbook,
        [string]$FormatKey,
        [string]$ExpectedOutgoing,
        [int]$ExpectedPages
    )

    $sheetName = "DispatchLayout_" + $FormatKey.ToUpperInvariant()
    $sheet = $Workbook.Worksheets.Item($sheetName)
    $barPrefix = "EnvelopeDynamic_PostalBar_" + $FormatKey.ToLowerInvariant() + "_"
    $outgoingPrefix = "EnvelopeDynamic_Outgoing_" + $FormatKey.ToLowerInvariant() + "_"
    $postalBarCount = 0
    $hasOutgoingText = $false
    $hasDigitOneSlant = $false

    for ($shapeIndex = 1; $shapeIndex -le $sheet.Shapes.Count; $shapeIndex++) {
        $shape = $sheet.Shapes.Item($shapeIndex)
        $shapeName = [string]$shape.Name

        if ($shapeName.StartsWith($barPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $postalBarCount++
        }

        if ($shapeName.StartsWith($outgoingPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $shapeText = Get-ShapeText -Shape $shape
            if ($shapeText.Contains($ExpectedOutgoing)) {
                $hasOutgoingText = $true
            }
        }

        if ($shapeName.Contains("_SLANT")) {
            $hasDigitOneSlant = $true
        }
    }

    $printArea = [string]$sheet.PageSetup.PrintArea
    return [PSCustomObject]@{
        SheetName = $sheetName
        PostalBarCount = $postalBarCount
        HasOutgoingText = $hasOutgoingText
        HasDigitOneSlant = $hasDigitOneSlant
        PrintArea = $printArea
        Visible = [int]$sheet.Visible
        ExpectedPostalBarCount = $ExpectedPages * 7
    }
}

function Invoke-WorkbookRepair {
    param(
        [string]$TargetPath,
        [string]$ScriptsDirectory
    )

    $repairScript = Join-Path $ScriptsDirectory "repair_workbook_package.py"
    if (Test-Path -LiteralPath $repairScript) {
        python $repairScript $TargetPath | Out-Null
    }
}

$resolvedWorkbookPath = Resolve-Path $WorkbookPath
$projectDirectory = Split-Path -Parent $resolvedWorkbookPath.Path
$scriptsDirectory = Join-Path $projectDirectory "scripts"
$tempDirectory = Join-Path $projectDirectory "filesarchive\temp-com-tests"
$results = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$workbook = $null
$failed = $false
$tempWorkbookPath = $null

try {
    New-Item -ItemType Directory -Path $tempDirectory -Force | Out-Null

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $tempWorkbookPath = Join-Path $tempDirectory ("dispatch-envelope-smoke-" + $timestamp + ".xlsm")
    Copy-Item -LiteralPath $resolvedWorkbookPath.Path -Destination $tempWorkbookPath -Force
    Invoke-WorkbookRepair -TargetPath $tempWorkbookPath -ScriptsDirectory $scriptsDirectory

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open((Get-ExcelOpenPath -Path $tempWorkbookPath))
    Add-SmokeResult -Results $results -Name "WorkbookTempOpen" -Status "PASS" -Details $tempWorkbookPath

    $sendersTable = Get-Table -Workbook $workbook -SheetName "Senders" -TableName "tblSenders"
    $dispatchItemsTable = Get-Table -Workbook $workbook -SheetName "DispatchItems" -TableName "tblDispatchItems"
    $dispatchRegistryTable = Get-Table -Workbook $workbook -SheetName "DispatchRegistry" -TableName "tblDispatchRegistry"

    Clear-TableRows -Table $sendersTable
    Clear-TableRows -Table $dispatchItemsTable
    Clear-TableRows -Table $dispatchRegistryTable

    Add-TableRow -Table $sendersTable -Values @{
        SenderName = "AIF Sender"
        AddressLine1 = "Sender city"
        AddressLine2 = "Sender street"
        AddressLine3 = ""
        PostalCode = "364029"
        Phone = ""
        IsDefault = "TRUE"
    }

    $registryNumber = "AIF-SMOKE"
    $registryDate = Get-Date -Format "dd.MM.yyyy"
    $formats = @("c4", "c5", "dl")

    foreach ($formatKey in $formats) {
        $batchId = "aif-envelope-smoke-" + $formatKey
        Add-TableRow -Table $dispatchItemsTable -Values @{
            DispatchId = $batchId + "-1"
            LetterNumber = "7/101"
            LetterDate = "01.05.2026"
            LetterRowNumber = "101"
            Addressee = "AIF Addressee " + $formatKey.ToUpperInvariant()
            AddressLine = "Recipient street, Recipient city, 355017"
            PostalCode = "355017"
            SenderName = "AIF Sender"
            EnvelopeFormatKey = $formatKey
            MailType = "registered"
            Mass = ""
            DeclaredValue = ""
            Comment = ""
            Phone = ""
            BatchId = $batchId
            Status = "packed"
            CreatedAt = $registryDate
            RegistryNumber = $registryNumber
            RegistryDate = $registryDate
        }
        Add-TableRow -Table $dispatchItemsTable -Values @{
            DispatchId = $batchId + "-2"
            LetterNumber = "7/102"
            LetterDate = "01.05.2026"
            LetterRowNumber = "102"
            Addressee = "AIF Addressee " + $formatKey.ToUpperInvariant()
            AddressLine = "Recipient street, Recipient city, 355017"
            PostalCode = "355017"
            SenderName = "AIF Sender"
            EnvelopeFormatKey = $formatKey
            MailType = "registered"
            Mass = ""
            DeclaredValue = ""
            Comment = ""
            Phone = ""
            BatchId = $batchId
            Status = "packed"
            CreatedAt = $registryDate
            RegistryNumber = $registryNumber
            RegistryDate = $registryDate
        }

        if ($formatKey -eq "c5") {
            $secondBatchId = "aif-envelope-smoke-c5-second"
            Add-TableRow -Table $dispatchItemsTable -Values @{
                DispatchId = $secondBatchId + "-1"
                LetterNumber = "7/201"
                LetterDate = "02.05.2026"
                LetterRowNumber = "201"
                Addressee = "AIF Addressee C5 Second"
                AddressLine = "Second recipient street, Second city, 355017"
                PostalCode = "355017"
                SenderName = "AIF Sender"
                EnvelopeFormatKey = $formatKey
                MailType = "registered"
                Mass = ""
                DeclaredValue = ""
                Comment = ""
                Phone = ""
                BatchId = $secondBatchId
                Status = "packed"
                CreatedAt = $registryDate
                RegistryNumber = $registryNumber
                RegistryDate = $registryDate
            }
        }
    }

    $macroPrefix = "'" + $workbook.Name + "'!"
    $registryRows = [int]$excel.Run($macroPrefix + "BuildDispatchRegistry")
    if ($registryRows -eq 4) {
        Add-SmokeResult -Results $results -Name "BuildDispatchRegistry" -Status "PASS" -Details "Built 4 grouped registry rows."
    }
    else {
        Add-SmokeResult -Results $results -Name "BuildDispatchRegistry" -Status "FAIL" -Details ("Expected 4 rows, got " + $registryRows)
        $failed = $true
    }

    $preparedCount = [int]$excel.Run($macroPrefix + "PrepareEnvelopePrint")
    if ($preparedCount -eq 4) {
        Add-SmokeResult -Results $results -Name "PrepareEnvelopePrint" -Status "PASS" -Details "Prepared C4, two C5, and DL envelope pages."
    }
    else {
        Add-SmokeResult -Results $results -Name "PrepareEnvelopePrint" -Status "FAIL" -Details ("Expected 4 prepared envelopes, got " + $preparedCount)
        $failed = $true
    }

    $expectedPagesByFormat = @{
        c4 = 1
        c5 = 2
        dl = 1
    }

    foreach ($formatKey in $formats) {
        $sheetCheck = Test-EnvelopeSheet -Workbook $workbook -FormatKey $formatKey -ExpectedOutgoing "7/102" -ExpectedPages $expectedPagesByFormat[$formatKey]
        if ($sheetCheck.PostalBarCount -eq $sheetCheck.ExpectedPostalBarCount -and $sheetCheck.HasOutgoingText -and $sheetCheck.HasDigitOneSlant -and -not [string]::IsNullOrWhiteSpace($sheetCheck.PrintArea)) {
            Add-SmokeResult -Results $results -Name ("EnvelopeSheet:" + $formatKey) -Status "PASS" -Details ("bars=" + $sheetCheck.PostalBarCount + "; slantOne=" + $sheetCheck.HasDigitOneSlant + "; printArea=" + $sheetCheck.PrintArea)
        }
        else {
            Add-SmokeResult -Results $results -Name ("EnvelopeSheet:" + $formatKey) -Status "FAIL" -Details ("bars=" + $sheetCheck.PostalBarCount + "; expectedBars=" + $sheetCheck.ExpectedPostalBarCount + "; outgoing=" + $sheetCheck.HasOutgoingText + "; slantOne=" + $sheetCheck.HasDigitOneSlant + "; printArea=" + $sheetCheck.PrintArea)
            $failed = $true
        }
    }

    $workbook.Close($false)
    $workbook = $null
    Invoke-WorkbookRepair -TargetPath $tempWorkbookPath -ScriptsDirectory $scriptsDirectory
}
catch {
    Add-SmokeResult -Results $results -Name "DispatchEnvelopeSmoke" -Status "FAIL" -Details $_.Exception.Message
    $failed = $true
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if (-not $KeepTemp -and -not [string]::IsNullOrWhiteSpace($tempWorkbookPath)) {
        Remove-Item -LiteralPath $tempWorkbookPath -Force -ErrorAction SilentlyContinue
    }
}

foreach ($result in $results) {
    Write-Host ("[{0}] {1} - {2}" -f $result.Status, $result.Name, $result.Details)
}

if ($failed) {
    exit 1
}

