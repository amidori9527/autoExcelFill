param(
  [string]$Python = "python",
  [string]$ZipPath = "AutoExcelKit-Windows.zip"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host "Installing dependencies..."
& $Python -m pip install -r requirements.txt "pyinstaller>=6.11,<7"
if ($LASTEXITCODE -ne 0) {
  throw "Dependency installation failed with exit code $LASTEXITCODE"
}

Write-Host "Running tests..."
$env:PYTHONPATH = Join-Path $Root "src"
& $Python -m unittest discover -s tests
if ($LASTEXITCODE -ne 0) {
  throw "Tests failed with exit code $LASTEXITCODE"
}

Write-Host "Building Windows executables..."
& $Python -m PyInstaller --clean -y autoexcel-fill.spec
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$ExpectedExecutables = @(
  "autoexcel-fill.exe",
  "diff-orders.exe",
  "fetch-orders.exe",
  "add-cards.exe",
  "add-b2b.exe"
)
foreach ($Executable in $ExpectedExecutables) {
  $ExecutablePath = Join-Path $Root "dist\AutoExcelKit\$Executable"
  if (-not (Test-Path $ExecutablePath)) {
    throw "Build output is missing: $ExecutablePath"
  }
}

$ReleaseDir = Join-Path $Root "AutoExcelKit-Windows"
if (Test-Path $ReleaseDir) {
  Remove-Item $ReleaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $ReleaseDir | Out-Null

Copy-Item -Path (Join-Path $Root "dist\AutoExcelKit") -Destination (Join-Path $ReleaseDir "AutoExcelKit") -Recurse
$KitDir = Join-Path $ReleaseDir "AutoExcelKit"
Copy-Item -Path (Join-Path $Root "run-autoexcel-fill.bat") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "run-diff-orders.bat") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "run-fetch-orders.bat") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "run-add-cards.bat") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "run-add-b2b.bat") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "fetch-orders-user-guide.html") -Destination $KitDir
Copy-Item -Path (Join-Path $Root "RELEASE_README_Windows.txt") -Destination (Join-Path $KitDir "README.txt")

if (Test-Path $ZipPath) {
  Remove-Item $ZipPath -Force
}
Compress-Archive -Path (Join-Path $ReleaseDir "*") -DestinationPath $ZipPath -Force

Write-Host "Done: $ZipPath"
