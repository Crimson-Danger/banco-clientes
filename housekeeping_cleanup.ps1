param(
    [int]$RetentionDays = 90,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = $PSScriptRoot
$archiveDir = Join-Path $projectRoot "data\archive"
$logsDir = Join-Path $projectRoot "logs"
$cleanupLogPath = Join-Path $logsDir "housekeeping_cleanup.log"
$cutoff = (Get-Date).AddDays(-1 * $RetentionDays)

if (-not (Test-Path -LiteralPath $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir | Out-Null
}

function Write-CleanupLog {
    param(
        [string]$Level,
        [string]$Message
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-ddTHH:mm:ssK"), $Level, $Message
    Add-Content -Path $cleanupLogPath -Value $line
    Write-Host $line
}

function Remove-OldFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetDir,
        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )
    if (-not (Test-Path -LiteralPath $TargetDir)) {
        return [pscustomobject]@{
            scanned = 0
            deleted = 0
            freedBytes = 0
        }
    }

    $files = @(Get-ChildItem -LiteralPath $TargetDir -File -Recurse | Where-Object {
        $matchesPattern = $false
        foreach ($pattern in $Patterns) {
            if ($_.Name -like $pattern) {
                $matchesPattern = $true
                break
            }
        }
        $matchesPattern -and $_.LastWriteTime -lt $cutoff
    })

    $freedBytes = 0L
    $deleted = 0
    foreach ($file in $files) {
        $freedBytes += [int64]$file.Length
        if (-not $DryRun) {
            Remove-Item -LiteralPath $file.FullName -Force
        }
        $deleted += 1
    }

    return [pscustomobject]@{
        scanned = $files.Count
        deleted = $deleted
        freedBytes = $freedBytes
    }
}

try {
    Write-CleanupLog -Level "INFO" -Message ("Inicio housekeeping: retention_days={0}, cutoff={1}, dry_run={2}" -f $RetentionDays, $cutoff.ToString("yyyy-MM-dd"), [bool]$DryRun)

    $archiveStats = Remove-OldFiles -TargetDir $archiveDir -Patterns @("*")
    $rootLogStats = Remove-OldFiles -TargetDir $projectRoot -Patterns @("*.log")
    $appLogStats = Remove-OldFiles -TargetDir $logsDir -Patterns @("*.log")

    $totalDeleted = $archiveStats.deleted + $rootLogStats.deleted + $appLogStats.deleted
    $totalFreed = $archiveStats.freedBytes + $rootLogStats.freedBytes + $appLogStats.freedBytes

    Write-CleanupLog -Level "INFO" -Message ("Archive: deleted={0}, bytes={1}" -f $archiveStats.deleted, $archiveStats.freedBytes)
    Write-CleanupLog -Level "INFO" -Message ("RootLogs: deleted={0}, bytes={1}" -f $rootLogStats.deleted, $rootLogStats.freedBytes)
    Write-CleanupLog -Level "INFO" -Message ("AppLogs: deleted={0}, bytes={1}" -f $appLogStats.deleted, $appLogStats.freedBytes)
    Write-CleanupLog -Level "INFO" -Message ("Fim housekeeping: total_deleted={0}, total_freed_bytes={1}" -f $totalDeleted, $totalFreed)
}
catch {
    Write-CleanupLog -Level "ERROR" -Message ("Falha housekeeping: {0}" -f $_.Exception.Message)
    throw
}
