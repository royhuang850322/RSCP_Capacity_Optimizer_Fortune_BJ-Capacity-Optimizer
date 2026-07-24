$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot
python -m PyInstaller packaging\FortuneBJOptimizer.spec --clean --noconfirm
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}
$DistRoot = Join-Path $ProjectRoot "dist\FortuneBJOptimizer"
$DataDirName = -join ([char[]](0x6570, 0x636e, 0x5bfc, 0x5165))
$DemoDirName = -join ([char[]](0x6a21, 0x62df, 0x6570, 0x636e, 0x5bfc, 0x5165))
$ReportDirName = -join ([char[]](0x62a5, 0x544a))

$DataTarget = Join-Path $DistRoot $DataDirName
New-Item -ItemType Directory -Force -Path $DataTarget | Out-Null
$DataSource = Join-Path $ProjectRoot $DataDirName
foreach ($Pattern in @("*.csv", "*.xlsx", "*.xls")) {
    Copy-Item -Path (Join-Path $DataSource $Pattern) -Destination $DataTarget -Force -ErrorAction SilentlyContinue
}

$DemoSource = Join-Path $ProjectRoot $DemoDirName
if (Test-Path -LiteralPath $DemoSource) {
    $DemoTarget = Join-Path $DistRoot $DemoDirName
    New-Item -ItemType Directory -Force -Path $DemoTarget | Out-Null
    foreach ($Pattern in @("*.csv", "*.xlsx", "*.xls")) {
        Copy-Item -Path (Join-Path $DemoSource $Pattern) -Destination $DemoTarget -Force -ErrorAction SilentlyContinue
    }
}

$LicenseTarget = Join-Path $DistRoot "licenses\active"
New-Item -ItemType Directory -Force -Path $LicenseTarget | Out-Null
Copy-Item -LiteralPath (Join-Path $ProjectRoot "licenses\active\license.json") -Destination $LicenseTarget -Force
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot "licenses\requests") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot $ReportDirName) | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot "logs") | Out-Null
Write-Host "Fortune BJ Optimizer build finished: dist\FortuneBJOptimizer"
