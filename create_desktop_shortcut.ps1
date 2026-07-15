# Creates a Desktop shortcut that launches the PDF -> PowerPoint tool.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Definition
$bat  = Join-Path $root "PDF to PowerPoint.bat"
$desktop = [Environment]::GetFolderPath("Desktop")
$lnk = Join-Path $desktop "PDF to PowerPoint.lnk"

$shell = New-Object -ComObject WScript.Shell
$sc = $shell.CreateShortcut($lnk)
$sc.TargetPath = $bat
$sc.WorkingDirectory = $root
$sc.WindowStyle = 7   # minimized
$sc.Description = "Convert a PDF into an editable PowerPoint (runs locally)"

# Use PowerPoint's icon if we can find it, otherwise a generic document icon.
$ppCandidates = @(
  "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
  "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
  "C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE"
)
$icon = $null
foreach ($c in $ppCandidates) { if (Test-Path $c) { $icon = "$c,0"; break } }
if (-not $icon) { $icon = "$env:SystemRoot\System32\shell32.dll,70" }
$sc.IconLocation = $icon
$sc.Save()
Write-Host "Shortcut created: $lnk"
