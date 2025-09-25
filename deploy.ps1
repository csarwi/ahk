# Deploy AHK Files to Startup Folder
# This script copies all .ahk files from the current directory to the Windows startup folder

Write-Host "AHK Files Deployment Script" -ForegroundColor Green
Write-Host "============================" -ForegroundColor Green

# Get the current directory (where the script is located)
$sourceFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Source folder: $sourceFolder" -ForegroundColor Cyan

# Get the startup folder path
$startupFolder = [Environment]::GetFolderPath('Startup')
Write-Host "Startup folder: $startupFolder" -ForegroundColor Cyan

# Find all .ahk files in the current directory
$ahkFiles = Get-ChildItem -Path $sourceFolder -Filter "*.ahk" -File

if ($ahkFiles.Count -eq 0) {
    Write-Host "No .ahk files found in the current directory." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Found $($ahkFiles.Count) AHK file(s):" -ForegroundColor Yellow
foreach ($file in $ahkFiles) {
    Write-Host "  - $($file.Name)" -ForegroundColor White
}

Write-Host ""
Write-Host "Deploying files..." -ForegroundColor Green

# Copy each .ahk file to the startup folder
$successCount = 0
$errorCount = 0

foreach ($file in $ahkFiles) {
    try {
        $destinationPath = Join-Path $startupFolder $file.Name
        Copy-Item -Path $file.FullName -Destination $destinationPath -Force
        Write-Host "  [OK] Deployed: $($file.Name)" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "  [ERROR] Failed to deploy: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}

# Summary
Write-Host ""
Write-Host "============================" -ForegroundColor Green
Write-Host "Deployment Summary:" -ForegroundColor Green
Write-Host "  Successfully deployed: $successCount" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "  Failed: $errorCount" -ForegroundColor Red
}
Write-Host "  Target folder: $startupFolder" -ForegroundColor Cyan

if ($successCount -gt 0) {
    Write-Host ""
    Write-Host "AHK files will now start automatically with Windows!" -ForegroundColor Green
}

Write-Host ""
Write-Host "Press Enter to continue..." -ForegroundColor Gray
Read-Host