# PowerShell script to properly build and package the SPFx solution
# This fixes the file structure issues

Write-Host "=== SPFx Build and Package Script ===" -ForegroundColor Green
Write-Host ""

# Step 1: Clean
Write-Host "Step 1: Cleaning previous build..." -ForegroundColor Yellow
gulp clean | Out-Null

# Step 2: Build
Write-Host "Step 2: Building solution..." -ForegroundColor Yellow
gulp build

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed. Exiting." -ForegroundColor Red
    exit 1
}

# Step 3: Fix file locations (workaround for TypeScript compilation issue)
Write-Host "Step 3: Fixing file locations..." -ForegroundColor Yellow
if (Test-Path "lib\TeamsAiMeetingAppWebPart.js") {
    Move-Item "lib\TeamsAiMeetingAppWebPart.js" "lib\webparts\teamsAiMeetingApp\TeamsAiMeetingAppWebPart.js" -Force
    Write-Host "  Moved main JS file" -ForegroundColor Gray
}
if (Test-Path "lib\TeamsAiMeetingAppWebPart.js.map") {
    Move-Item "lib\TeamsAiMeetingAppWebPart.js.map" "lib\webparts\teamsAiMeetingApp\TeamsAiMeetingAppWebPart.js.map" -Force
    Write-Host "  Moved source map" -ForegroundColor Gray
}
if (Test-Path "lib\TeamsAiMeetingAppWebPart.module.scss.js") {
    Move-Item "lib\TeamsAiMeetingAppWebPart.module.scss.js" "lib\webparts\teamsAiMeetingApp\TeamsAiMeetingAppWebPart.module.scss.js" -Force
    Write-Host "  Moved SCSS JS file" -ForegroundColor Gray
}

# Step 4: Bundle for production
Write-Host "Step 4: Bundling for production..." -ForegroundColor Yellow
gulp bundle --ship

if ($LASTEXITCODE -ne 0) {
    Write-Host "Bundle failed. Exiting." -ForegroundColor Red
    exit 1
}

# Step 5: Package solution
Write-Host "Step 5: Creating solution package..." -ForegroundColor Yellow
gulp package-solution --ship

if ($LASTEXITCODE -ne 0) {
    Write-Host "Package creation failed. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== Build Complete! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Files created:" -ForegroundColor Cyan
if (Test-Path "sharepoint\solution\teams-meeting-app.sppkg") {
    Write-Host "  ✅ sharepoint\solution\teams-meeting-app.sppkg" -ForegroundColor Green
    $size = (Get-Item "sharepoint\solution\teams-meeting-app.sppkg").Length
    Write-Host "     Size: $([math]::Round($size/1024, 2)) KB" -ForegroundColor Gray
}
if (Test-Path "teams\TeamsSPFxApp.zip") {
    Write-Host "  ✅ teams\TeamsSPFxApp.zip (Teams app package)" -ForegroundColor Green
}

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "1. Upload the .sppkg file to SharePoint App Catalog" -ForegroundColor White
Write-Host "2. Deploy with 'Make available to all sites' checked" -ForegroundColor White
Write-Host "3. Sync to Teams or upload the Teams package" -ForegroundColor White
