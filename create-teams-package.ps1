# PowerShell script to create Teams App Package
# Run this script from the root of the teams-meeting-app directory

Write-Host "=== Teams AI Meeting App Packager ===" -ForegroundColor Green
Write-Host ""

# Check if we're in the correct directory
if (!(Test-Path "teams\manifest.json")) {
    Write-Host "Error: teams\manifest.json not found. Please run this script from the teams-meeting-app root directory." -ForegroundColor Red
    exit 1
}

# Create output directory
$outputDir = "dist"
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Write-Host "Creating Teams app package..." -ForegroundColor Yellow

# Check for icon files
$iconFiles = @(
    "teams\a1b2c3d4-e5f6-789a-bcde-f0123456789a_color.png",
    "teams\a1b2c3d4-e5f6-789a-bcde-f0123456789a_outline.png"
)

$missingIcons = @()
foreach ($icon in $iconFiles) {
    if (!(Test-Path $icon)) {
        $missingIcons += $icon
    }
}

if ($missingIcons.Count -gt 0) {
    Write-Host "Warning: Missing icon files:" -ForegroundColor Yellow
    foreach ($icon in $missingIcons) {
        Write-Host "  - $icon" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Please create the following icon files:" -ForegroundColor Yellow
    Write-Host "  - Color icon: 96x96 pixels, full color PNG" -ForegroundColor Yellow
    Write-Host "  - Outline icon: 32x32 pixels, white outline on transparent PNG" -ForegroundColor Yellow
    Write-Host ""
    
    $continue = Read-Host "Continue without icons? (y/N)"
    if ($continue -ne "y" -and $continue -ne "Y") {
        Write-Host "Aborted." -ForegroundColor Red
        exit 1
    }
}

# Files to include in the package
$files = @("teams\manifest.json")
foreach ($icon in $iconFiles) {
    if (Test-Path $icon) {
        $files += $icon
    }
}

# Create the ZIP package
$packagePath = "$outputDir\TeamsSPFxApp.zip"
if (Test-Path $packagePath) {
    Remove-Item $packagePath -Force
}

try {
    # Create temporary directory for packaging
    $tempDir = "temp_package"
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tempDir | Out-Null
    
    # Copy files to temp directory with correct names
    Copy-Item "teams\manifest.json" "$tempDir\manifest.json"
    
    foreach ($icon in $iconFiles) {
        if (Test-Path $icon) {
            $iconName = Split-Path $icon -Leaf
            Copy-Item $icon "$tempDir\$iconName"
        }
    }
    
    # Create ZIP from temp directory contents
    Compress-Archive -Path "$tempDir\*" -DestinationPath $packagePath -Force
    
    # Cleanup temp directory
    Remove-Item $tempDir -Recurse -Force
    
    Write-Host "Success! Teams app package created: $packagePath" -ForegroundColor Green
    
    # Also copy to teams directory for SPFx integration
    $spfxPackagePath = "teams\TeamsSPFxApp.zip"
    Copy-Item $packagePath $spfxPackagePath -Force
    Write-Host "Package also copied to: $spfxPackagePath" -ForegroundColor Green
    
} catch {
    Write-Host "Error creating package: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Next Steps ===" -ForegroundColor Cyan
Write-Host "1. Build SPFx solution:" -ForegroundColor White
Write-Host "   gulp bundle --ship" -ForegroundColor Gray
Write-Host "   gulp package-solution --ship" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Deploy SPFx package to SharePoint App Catalog" -ForegroundColor White
Write-Host ""
Write-Host "3. Sync to Teams or upload $packagePath to Teams admin center" -ForegroundColor White
Write-Host ""
Write-Host "Package contents:"
Get-ChildItem $tempDir -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Done!" -ForegroundColor Green
