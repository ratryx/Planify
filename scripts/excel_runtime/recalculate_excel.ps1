param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)

$ErrorActionPreference = "Stop"
$xlDone = 0
$timeoutSeconds = 60

# Resolve absolute path
$absolutePath = [System.IO.Path]::GetFullPath((Resolve-Path -Path $FilePath -ErrorAction SilentlyContinue).Path)
if (-not (Test-Path $absolutePath)) {
    Write-Error "File not found: $absolutePath"
    exit 1
}

Write-Host "Starting Excel COM for $absolutePath"

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Excel COM Created. Version: $($excel.Version)"

    $workbook = $excel.Workbooks.Open($absolutePath, 0, $false)
    Write-Host "Workbook opened. Triggering CalculateFullRebuild()..."

    $excel.CalculateFullRebuild()

    $startTime = Get-Date
    while ($excel.CalculationState -ne $xlDone) {
        $elapsed = (Get-Date) - $startTime
        if ($elapsed.TotalSeconds -gt $timeoutSeconds) {
            Write-Error "Timeout waiting for calculation to finish after $timeoutSeconds seconds."
            exit 1
        }
        Start-Sleep -Milliseconds 500
    }

    Write-Host "Calculation complete."

    $workbook.Save()
    Write-Host "Workbook saved."
}
catch {
    Write-Error "Error during Excel COM execution: $_"
    exit 1
}
finally {
    if ($workbook -ne $null) {
        try {
            $workbook.Close($false)
        } catch {}
    }
    if ($excel -ne $null) {
        try {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Remove-Variable excel -ErrorAction SilentlyContinue
        } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Excel COM process finished successfully."
exit 0
