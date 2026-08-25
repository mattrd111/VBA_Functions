<#
.SYNOPSIS
    Builds WorkbookDoctor.xlam from the VBA source in this repository.

.DESCRIPTION
    Excel add-ins are binary, so the .xlam is not kept in source control - this
    script assembles one on your machine by driving Excel.

    It creates a new workbook, imports every module from \src and \addin, copies
    the ThisWorkbook code into place, flips the workbook into add-in mode and
    saves it as .xlam.

    Requires:
      - Windows with Excel installed
      - "Trust access to the VBA project object model" switched on:
        Excel > File > Options > Trust Center > Trust Center Settings >
        Macro Settings > tick "Trust access to the VBA project object model"
        (turn it back off afterwards if you would rather)

.PARAMETER OutputPath
    Where to write the .xlam. Defaults to \dist\WorkbookDoctor.xlam.

.PARAMETER Install
    After building, copy the add-in to your Excel AddIns folder and switch it on.

.EXAMPLE
    .\Build-AddIn.ps1

.EXAMPLE
    .\Build-AddIn.ps1 -Install
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$Install
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputPath) {
    $OutputPath = Join-Path $repoRoot 'dist\WorkbookDoctor.xlam'
}

# Library modules first, then the add-in itself.
$moduleFiles = @(
    'src\modApp.bas',
    'src\modArray.bas',
    'src\modDate.bas',
    'src\modDictionary.bas',
    'src\modFile.bas',
    'src\modRange.bas',
    'src\modString.bas',
    'src\modWorkbook.bas',
    'addin\modDoctorCommon.bas',
    'addin\modDoctorScan.bas',
    'addin\modAuditFormula.bas',
    'addin\modAuditCore.bas',
    'addin\modModelAudit.bas',
    'addin\modDoctorNames.bas',
    'addin\modDoctorStyles.bas',
    'addin\modDoctorSheets.bas',
    'addin\modDoctorLinks.bas',
    'addin\modDoctorTools.bas',
    'addin\modDoctorAudit.bas',
    'addin\modDoctorRunner.bas',
    'addin\modWrangleStack.bas',
    'addin\modWrangleShape.bas',
    'addin\modWrangleMatch.bas',
    'addin\modDoctorMenu.bas'
)
$thisWorkbookFile = Join-Path $repoRoot 'addin\ThisWorkbook.cls'

# --- checks ----------------------------------------------------------------
foreach ($relative in $moduleFiles) {
    $full = Join-Path $repoRoot $relative
    if (-not (Test-Path -LiteralPath $full)) {
        throw "Missing source file: $full"
    }
}
if (-not (Test-Path -LiteralPath $thisWorkbookFile)) {
    throw "Missing source file: $thisWorkbookFile"
}

$outputFolder = Split-Path -Parent $OutputPath
if ($outputFolder -and -not (Test-Path -LiteralPath $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

Write-Host "Building $OutputPath" -ForegroundColor Cyan

# --- drive Excel -----------------------------------------------------------
$excel = $null
$workbook = $null
$previousSheetCount = $null

try {
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Could not start Excel. This script needs Windows with Excel installed."
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $previousSheetCount = $excel.SheetsInNewWorkbook
    $excel.SheetsInNewWorkbook = 1

    $workbook = $excel.Workbooks.Add()

    try {
        $project = $workbook.VBProject
    }
    catch {
        throw @"
Excel would not let the script touch the VBA project.

Switch on: File > Options > Trust Center > Trust Center Settings >
           Macro Settings > Trust access to the VBA project object model

then run this script again. (You can switch it back off afterwards - it is only
needed while building.)
"@
    }

    foreach ($relative in $moduleFiles) {
        $full = Join-Path $repoRoot $relative
        Write-Host "  importing $relative"
        $project.VBComponents.Import($full) | Out-Null
    }

    # A document module cannot be imported as a new component, so its code is
    # copied into the ThisWorkbook that already exists.
    Write-Host "  merging addin\ThisWorkbook.cls"
    $lines = Get-Content -LiteralPath $thisWorkbookFile
    $start = ($lines | Select-String -SimpleMatch 'Option Explicit' | Select-Object -First 1).LineNumber
    if (-not $start) {
        throw "addin\ThisWorkbook.cls does not contain an 'Option Explicit' line to start from."
    }
    $code = ($lines | Select-Object -Skip ($start - 1)) -join "`r`n"
    $project.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString($code)

    # --- turn it into an add-in --------------------------------------------
    try {
        $workbook.BuiltinDocumentProperties.Item('Title').Value = 'Workbook Doctor'
        $workbook.BuiltinDocumentProperties.Item('Comments').Value =
            'Clean-up tools for heavy Excel workbooks. github.com/mattrd111/VBA_Functions'
    }
    catch {
        Write-Warning "Could not set the document title - carrying on."
    }
    $workbook.IsAddin = $true

    $xlOpenXMLAddIn = 55
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }
    $workbook.SaveAs($OutputPath, $xlOpenXMLAddIn)
    $workbook.Close($false)
    $workbook = $null

    Write-Host "Built $OutputPath" -ForegroundColor Green

    # --- optional install ---------------------------------------------------
    if ($Install) {
        $addInFolder = Join-Path $env:APPDATA 'Microsoft\AddIns'
        if (-not (Test-Path -LiteralPath $addInFolder)) {
            New-Item -ItemType Directory -Path $addInFolder -Force | Out-Null
        }
        $installed = Join-Path $addInFolder (Split-Path -Leaf $OutputPath)
        Copy-Item -LiteralPath $OutputPath -Destination $installed -Force

        $addIn = $excel.AddIns.Add($installed, $false)
        $addIn.Installed = $true

        Write-Host "Installed to $installed and switched on." -ForegroundColor Green
        Write-Host "Open Excel and look for 'Workbook Doctor' on the Add-ins tab." -ForegroundColor Green
    }
    else {
        Write-Host ""
        Write-Host "To use it:" -ForegroundColor Cyan
        Write-Host "  Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse"
        Write-Host "  pick $OutputPath, then tick it."
        Write-Host "  The menu appears as 'Workbook Doctor' on the Add-ins tab."
        Write-Host ""
        Write-Host "Or re-run this script with -Install to do that for you."
    }
}
finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch { }
    }
    if ($excel) {
        try {
            if ($null -ne $previousSheetCount) { $excel.SheetsInNewWorkbook = $previousSheetCount }
            $excel.DisplayAlerts = $true
            $excel.Quit()
        }
        catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
