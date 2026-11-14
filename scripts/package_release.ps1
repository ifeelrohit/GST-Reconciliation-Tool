$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$ProjectRoot = "C:\GST-Reconciliation-Tool-main"
$Staging = Join-Path $ProjectRoot "staging_package"
$ReleaseDist = Join-Path $ProjectRoot "release_distribution"
$ZipTarget = Join-Path $ReleaseDist "GST-Reconciliation-Tool-Windows-x64.zip"

Write-Host "1. Refreshing staging package directory..."
if (Test-Path $Staging) { Remove-Item -Recurse -Force $Staging }
New-Item -ItemType Directory -Path $Staging -Force | Out-Null

Write-Host "2. Copying binaries and documentation..."
Copy-Item (Join-Path $ProjectRoot "dist\GST-Reconciliation-Tool.exe") -Destination $Staging
Copy-Item (Join-Path $ProjectRoot "gst_recon.py") -Destination $Staging
Copy-Item (Join-Path $ProjectRoot "Run-GST-Tool.bat") -Destination $Staging
Copy-Item (Join-Path $ProjectRoot "Run-GST-Tool.vbs") -Destination $Staging
Copy-Item (Join-Path $ProjectRoot "ACCOUNTANT-INSTRUCTIONS.txt") -Destination $Staging
Copy-Item (Join-Path $ProjectRoot "README-WINDOWS.txt") -Destination $Staging

# Copy demo sample files
$StagingSamples = Join-Path $Staging "Sample Files"
New-Item -ItemType Directory -Path $StagingSamples -Force | Out-Null
Copy-Item (Join-Path $ProjectRoot "Sample Files\Demo_*.xlsx") -Destination $StagingSamples

Write-Host "3. Generating SHA256SUMS.txt manifest..."
$ManifestLines = @()
Get-ChildItem -Path $Staging -Recurse -File | ForEach-Object {
    $relPath = $_.FullName.Substring($Staging.Length).TrimStart('\', '/')
    $hash = (Get-FileHash -Path $_.FullName -Algorithm SHA256).Hash.ToLower()
    $ManifestLines += "$hash  $relPath"
}
$ManifestPath = Join-Path $Staging "SHA256SUMS.txt"
$ManifestLines | Out-File -FilePath $ManifestPath -Encoding utf8

Write-Host "4. Creating ZIP archive..."
if (Test-Path $ZipTarget) { Remove-Item -Force $ZipTarget }
if (-not (Test-Path $ReleaseDist)) { New-Item -ItemType Directory -Path $ReleaseDist -Force | Out-Null }
Compress-Archive -Path "$Staging\*" -DestinationPath $ZipTarget -CompressionLevel Optimal

$ZipHash = (Get-FileHash -Path $ZipTarget -Algorithm SHA256).Hash.ToLower()
$ZipSize = (Get-Item $ZipTarget).Length

Write-Host "========================================================================" -ForegroundColor Green
Write-Host " PACKAGING COMPLETE!" -ForegroundColor Green
Write-Host " ZIP Archive : $ZipTarget" -ForegroundColor Green
Write-Host " ZIP Size    : $ZipSize bytes ($([Math]::Round($ZipSize / 1MB, 2)) MB)" -ForegroundColor Green
Write-Host " SHA-256     : $ZipHash" -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Green
