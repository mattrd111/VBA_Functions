<#
.SYNOPSIS
  Runs the VBA test suite against a workbook, headless, on Windows.

.DESCRIPTION
  Opens Excel via COM, imports every module from src\ and tests\ into a scratch
  copy of the workbook, runs modTestRunner.RunAllTests, then reads the JSON
  result file the runner writes. Exit code 0 = all green, 1 = failures,
  2 = could not run (no Excel, VBOM access disabled, missing workbook).

  Requires: Windows, Excel, and Trust Center > Macro Settings >
  "Trust access to the VBA project object model" enabled.

.EXAMPLE
  pwsh -File tools/Invoke-VbaTests.ps1 -Workbook .\Book.xlsm
#>
[CmdletBinding()]
param(
  [string]$Workbook = "",
  [string]$RepoRoot = (Resolve-Path "$PSScriptRoot\..").Path,
  [switch]$KeepScratch
)

$ErrorActionPreference = 'Stop'

function Fail($message, $code = 2) {
  Write-Error $message
  exit $code
}

if (-not $IsWindows -and $PSVersionTable.PSEdition -eq 'Core') {
  Fail "Excel COM automation only runs on Windows. Run the static gate (tools/vba_lint.py) here and this script on a Windows host."
}

# A scratch macro-enabled workbook keeps the tracked workbook untouched.
$scratch = Join-Path $env:TEMP ("vba-tests-" + [guid]::NewGuid().ToString('N') + ".xlsm")
if ($Workbook -and (Test-Path $Workbook)) {
  Copy-Item -LiteralPath $Workbook -Destination $scratch -Force
}

$excel = $null
$book = $null
try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  Fail "Could not start Excel: $($_.Exception.Message)"
}

try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable - no macros auto-run on open

  if (Test-Path $scratch) {
    $book = $excel.Workbooks.Open($scratch)
  } else {
    $book = $excel.Workbooks.Add()
    $book.SaveAs($scratch, 52)  # xlOpenXMLWorkbookMacroEnabled
  }

  try { $null = $book.VBProject.VBComponents.Count } catch {
    Fail "No access to the VBA project. Enable Trust Center > Macro Settings > 'Trust access to the VBA project object model'."
  }

  $sources = @()
  foreach ($dir in @("src", "tests")) {
    $path = Join-Path $RepoRoot $dir
    if (Test-Path $path) {
      $sources += Get-ChildItem -Path $path -Include *.bas, *.cls, *.frm -File -Recurse
    }
  }
  if (-not $sources) { Fail "No .bas/.cls/.frm files found under src\ or tests\." }

  foreach ($file in $sources) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
    $existing = $book.VBProject.VBComponents | Where-Object { $_.Name -eq $name }
    if ($existing -and $existing.Type -ne 100) { $book.VBProject.VBComponents.Remove($existing) }
    $null = $book.VBProject.VBComponents.Import($file.FullName)
  }
  Write-Host "Imported $($sources.Count) component(s)."

  $resultFile = Join-Path (Split-Path $scratch -Parent) "vba-test-results.json"
  if (Test-Path $resultFile) { Remove-Item $resultFile -Force }

  $excel.Run("modTestRunner.RunAllTests")

  if (-not (Test-Path $resultFile)) { Fail "Test runner produced no result file - RunAllTests likely errored before writing." }
  $results = Get-Content $resultFile -Raw | ConvertFrom-Json

  Write-Host $results.log
  Write-Host "passed=$($results.passed) failed=$($results.failed) assertions=$($results.assertions)"
  foreach ($failure in $results.failures) { Write-Host "  FAIL: $failure" }

  $summary = Join-Path $RepoRoot "vba-test-results.json"
  Copy-Item $resultFile $summary -Force
  Write-Host "Results written to $summary"

  if ($results.failed -gt 0) { exit 1 }
  exit 0
}
finally {
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  if (-not $KeepScratch -and (Test-Path $scratch)) { Remove-Item $scratch -Force -ErrorAction SilentlyContinue }
}
