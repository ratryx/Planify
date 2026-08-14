$ErrorActionPreference = "Stop"

Write-Host "============================================="
Write-Host "     Planify Excel Runtime Validation        "
Write-Host "============================================="

Write-Host "`n1. Generating Runtime Fixtures A, B, C..."
python scripts/excel_runtime/generate_runtime_fixtures.py
if ($LASTEXITCODE -ne 0) { Write-Error "Generation failed"; exit 1 }

$fixtures = @(
    "runtime-output/finance_personal_runtime_a.xlsx",
    "runtime-output/finance_personal_runtime_b.xlsx",
    "runtime-output/finance_personal_runtime_c.xlsx"
)

$totalFormulas = 0
$totalErrors = 0

foreach ($fixture in $fixtures) {
    Write-Host "`n---------------------------------------------"
    Write-Host "Processing $fixture"
    Write-Host "---------------------------------------------"

    # 2. Recalculate and Save
    powershell -ExecutionPolicy Bypass -File scripts/excel_runtime/recalculate_excel.ps1 -FilePath $fixture
    if ($LASTEXITCODE -ne 0) { Write-Error "Recalculation failed for $fixture"; exit 1 }

    # 3. Scan Errors
    $scanOutput = python scripts/excel_runtime/scan_excel_errors.py $fixture 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host $scanOutput
        Write-Error "Error scanner found errors in $fixture";
        exit 1
    }
    Write-Host $scanOutput

    $scanText = $scanOutput -join "`n"
    if ($scanText -match "Scan complete\. (\d+) formulas checked\. (\d+) errors\.") {
        $totalFormulas += [int]$matches[1]
        $totalErrors += [int]$matches[2]
    }
    
    # 4. Status Validation (Only for B and C)
    if ($fixture -match "_runtime_[bc]\.xlsx$") {
        Write-Host "`n4. Validating semantic status for $fixture..."
        $statusOutput = python scripts/excel_runtime/validate_status.py $fixture 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host $statusOutput
            Write-Error "Status validation failed in $fixture"
            exit 1
        }
        Write-Host $statusOutput
    }
}

Write-Host "`n============================================="
Write-Host "             VALIDATION PASSED               "
Write-Host "Total Formulas Scanned : $totalFormulas"
Write-Host "Total Excel Errors     : $totalErrors"
Write-Host "============================================="
exit 0
