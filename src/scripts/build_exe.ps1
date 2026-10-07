# Build the single-file Windows exe with PyInstaller.
#
# The build runs from a copy in a neutral folder so that no personal path (your user
# name) is baked into the compiled code, then copies the exe to ./dist.
#
#   ./scripts/build_exe.ps1
#
$root = Split-Path -Parent $PSScriptRoot
$work = Join-Path $env:PUBLIC "usb_tool_build"
$name = "USB_Performance_Tool"

if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory $work | Out-Null
foreach ($f in "usb_gui.py", "usb_test.py", "excel_report.py", "browser_utils.py", "app_paths.py",
               "generate_test_files.py", "chart.umd.min.js", "file_set_manifest.json") {
    Copy-Item (Join-Path $root $f) $work
}

Push-Location $work
try {
    # PyInstaller reports progress on stderr; judge success by its exit code, not by that output.
    pyinstaller --name $name --onefile --windowed --collect-all playwright `
        --hidden-import usb_test --hidden-import browser_utils --hidden-import generate_test_files `
        --hidden-import app_paths --hidden-import excel_report --hidden-import xlsxwriter `
        usb_gui.py 2>&1 | ForEach-Object { "$_" } | Select-Object -Last 3
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }
    New-Item -ItemType Directory (Join-Path $root "dist") -Force | Out-Null
    Copy-Item (Join-Path $work "dist\$name.exe") (Join-Path $root "dist") -Force
    Write-Host "Built dist\$name.exe"
} finally {
    Pop-Location
}

# Run the exe next to chart.umd.min.js, file_set_manifest.json and test_config.example.json
# (the app finds its data files beside the exe, not inside it).
