# build_xlam.ps1 - assemble PerfDiag.xlam from the .bas modules + inject the ribbon.
# Requires: Excel installed; "Trust access to the VBA project object model" ON.
# Run from this folder:  powershell -File build_xlam.ps1
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlam = Join-Path $here "PerfDiag.xlam"
$mods = @("modTimer.bas", "modCollect.bas", "modReport.bas", "modRibbon.bas") |
        ForEach-Object { Join-Path $here $_ }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    foreach ($m in $mods) {
        $wb.VBProject.VBComponents.Import($m) | Out-Null
        Write-Host "imported $(Split-Path -Leaf $m)"
    }
    $wb.IsAddin = $true
    $wb.SaveAs($xlam, 55)   # 55 = xlOpenXMLAddIn (.xlam)
    $wb.Close($false)
    Write-Host "built $xlam"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

# Inject the ribbon part (VBComponents.Import cannot add custom UI).
python (Join-Path $here "inject_ribbon.py")
Write-Host "done. Load PerfDiag.xlam in Excel; use the 'Perf Diagnostic' ribbon tab."
