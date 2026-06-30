param(
  [string]$Python = "python",
  [string]$ZipPath = "AutoExcelKit-Windows.zip"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host "Installing dependencies..."
& $Python -m pip install -r requirements.txt
& $Python -m pip install pyinstaller

Write-Host "Building Windows executables..."
& $Python -m PyInstaller -y autoexcel-fill.spec

$ReleaseDir = Join-Path $Root "AutoExcelKit-Windows"
if (Test-Path $ReleaseDir) {
  Remove-Item $ReleaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $ReleaseDir | Out-Null

Copy-Item -Path (Join-Path $Root "dist\AutoExcelKit") -Destination (Join-Path $ReleaseDir "AutoExcelKit") -Recurse
Copy-Item -Path (Join-Path $Root "run-autoexcel-fill.bat") -Destination $ReleaseDir
Copy-Item -Path (Join-Path $Root "run-diff-orders.bat") -Destination $ReleaseDir
Copy-Item -Path (Join-Path $Root "run-fetch-orders.bat") -Destination $ReleaseDir
Copy-Item -Path (Join-Path $Root "fetch-orders-user-guide.html") -Destination $ReleaseDir
Copy-Item -Path (Join-Path $Root "RELEASE_README_Windows.txt") -Destination (Join-Path $ReleaseDir "README.txt")

if (Test-Path $ZipPath) {
  Remove-Item $ZipPath -Force
}
Compress-Archive -Path $ReleaseDir -DestinationPath $ZipPath -Force

Write-Host "Done: $ZipPath"
