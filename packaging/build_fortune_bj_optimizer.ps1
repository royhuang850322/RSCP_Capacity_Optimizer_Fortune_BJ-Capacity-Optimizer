$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot
python -m PyInstaller packaging\FortuneBJOptimizer.spec --clean --noconfirm
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}
$DistRoot = Join-Path $ProjectRoot "dist\FortuneBJOptimizer"
$DataDirName = -join ([char[]](0x6570, 0x636e, 0x5bfc, 0x5165))
$ReportDirName = -join ([char[]](0x62a5, 0x544a))

$DataTarget = Join-Path $DistRoot $DataDirName
New-Item -ItemType Directory -Force -Path $DataTarget | Out-Null
$DataSource = Join-Path $ProjectRoot $DataDirName
foreach ($Pattern in @("*.csv", "*.xlsx", "*.xls")) {
    Copy-Item -Path (Join-Path $DataSource $Pattern) -Destination $DataTarget -Force -ErrorAction SilentlyContinue
}

$LicenseTarget = Join-Path $DistRoot "licenses\active"
New-Item -ItemType Directory -Force -Path $LicenseTarget | Out-Null
Copy-Item -LiteralPath (Join-Path $ProjectRoot "licenses\active\license.json") -Destination $LicenseTarget -Force
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot "licenses\requests") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot $ReportDirName) | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DistRoot "logs") | Out-Null
$GuideTitle = -join ([char[]](0x64CD, 0x4F5C, 0x5458, 0x5347, 0x7EA7, 0x8BF4, 0x660E, 0x4E0E, 0x4EEA, 0x8868, 0x76D8, 0x9605, 0x8BFB, 0x6307, 0x5F15))
$GuideName = "Fortune_BJ_${GuideTitle}_CN_V07.docx"
$GuideSource = Join-Path $ProjectRoot "docs\$GuideName"
if (Test-Path -LiteralPath $GuideSource) {
    Copy-Item -LiteralPath $GuideSource -Destination (Join-Path $DistRoot $GuideName) -Force
}
Write-Host "Fortune BJ Optimizer build finished: dist\FortuneBJOptimizer"
