# Excel2Markdown Batch Test Script
# Test various functions of the tool

Write-Host "Excel2Markdown Batch Test Tool" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green
Write-Host ""

# Get project root directory
$projectRoot = Split-Path $PSScriptRoot -Parent
$publishPath = Join-Path $projectRoot "Publish\Excel2Markdown.exe"
$testFile = Join-Path $projectRoot "test.xlsx"

# Check if executable exists
if (-not (Test-Path $publishPath)) {
    Write-Host "Error: Cannot find Excel2Markdown.exe" -ForegroundColor Red
    Write-Host "Please run publish command first: dotnet publish ..." -ForegroundColor Yellow
    exit 1
}

# Step 1: Create test file
Write-Host "Step 1: Creating test file..." -ForegroundColor Cyan
dotnet run --project (Join-Path $PSScriptRoot "TestFileCreator.csproj")

if (-not (Test-Path $testFile)) {
    Write-Host "Error: Test file creation failed" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Test single file conversion
Write-Host "Step 2: Testing single file conversion..." -ForegroundColor Cyan
& $publishPath $testFile

# Check output file
$outputFile = Join-Path $projectRoot "test.md"
if (Test-Path $outputFile) {
    Write-Host "Single file conversion test passed" -ForegroundColor Green
    
    # Show partial output
    $content = Get-Content $outputFile -Head 10
    Write-Host "Generated Markdown content preview:" -ForegroundColor Yellow
    $content | ForEach-Object { Write-Host "   $_" -ForegroundColor Gray }
    Write-Host ""
} else {
    Write-Host "Single file conversion test failed" -ForegroundColor Red
}

# Step 3: Test help information
Write-Host "Step 3: Testing help information..." -ForegroundColor Cyan
& $publishPath

Write-Host ""

# Step 4: Clean up test files
Write-Host "Step 4: Cleaning up test files..." -ForegroundColor Cyan
if (Test-Path $testFile) {
    Remove-Item $testFile -Force
    Write-Host "Deleted test file: test.xlsx" -ForegroundColor Green
}

if (Test-Path $outputFile) {
    Remove-Item $outputFile -Force
    Write-Host "Deleted output file: test.md" -ForegroundColor Green
}

Write-Host ""
Write-Host "All tests completed!" -ForegroundColor Green