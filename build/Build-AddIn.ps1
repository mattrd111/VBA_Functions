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

.EXAMPLE
    .\Build-AddIn.ps1 -Loader -Install
    Builds the small loader instead, which each person installs once when the
    add-in is being kept up to date from a shared folder. See loader\README.md.
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$Install,
    [switch]$Loader
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$addInFileName = if ($Loader) { 'WorkbookDoctorLoader.xlam' } else { 'WorkbookDoctor.xlam' }
if (-not $OutputPath) {
    $OutputPath = Join-Path $repoRoot "dist\$addInFileName"
}

# Library modules first, then the add-in itself.
$moduleFiles = if ($Loader) { @('loader\modLoader.bas') } else { @(
    'src\modApp.bas',
    'src\modArray.bas',
    'src\modDate.bas',
    'src\modDictionary.bas',
    'src\modFinance.bas',
    'src\modWaterfall.bas',
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
    'addin\modWrangleBlocks.bas',
    'addin\modWrangleMatch.bas',
    'addin\modFundHelper.bas',
    'addin\modHouseStyle.bas',
    'addin\modDoctorMenu.bas'
) }
$thisWorkbookFile = if ($Loader) {
    Join-Path $repoRoot 'loader\ThisWorkbook.cls'
} else {
    Join-Path $repoRoot 'addin\ThisWorkbook.cls'
}

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

    # When the trust setting is off, PowerShell gets $null back from this rather
    # than an exception - so a catch alone would let a null through and the
    # failure would surface later as "cannot call a method on a null-valued
    # expression", which tells you nothing.
    $project = $null
    try { $project = $workbook.VBProject } catch { $project = $null }

    if (($null -eq $project) -or ($null -eq $project.VBComponents)) {
        throw @"
Excel would not let the script reach the workbook's VBA project.

This is the setting, and it is off by default:

  Excel > File > Options > Trust Center > Trust Center Settings >
  Macro Settings > tick "Trust access to the VBA project object model"

Tick it and run this script again. You can untick it afterwards - it is only
needed while the add-in is being built.

If that box is greyed out, your machine has it locked by policy. In that case
neither builder can work: import the modules by hand instead, following
addin\INSTALL.md.
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
        $workbook.BuiltinDocumentProperties.Item('Title').Value =
            $(if ($Loader) { 'Workbook Doctor Loader' } else { 'Workbook Doctor' })
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
