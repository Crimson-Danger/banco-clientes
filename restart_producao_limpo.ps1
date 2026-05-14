$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

function Test-TcpPort {
    param(
        [string]$HostName = "127.0.0.1",
        [int]$Port = 6379,
        [int]$TimeoutMs = 1200
    )
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($HostName, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $connected = $ok -and $tcp.Connected
        $tcp.Close()
        return $connected
    } catch {
        return $false
    }
}

function Stop-TargetPythonProcess {
    param([string]$Needle)
    $targets = Get-CimInstance Win32_Process -Filter "Name='python.exe'" | Where-Object {
        ($_.CommandLine -as [string]) -like "*$Needle*"
    }
    foreach ($p in $targets) {
        Write-Host "Finalizando processo PID=$($p.ProcessId) ($Needle)"
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }
}


Write-Host "==============================================="
Write-Host " RESTART LIMPO - PRODUCAO (API + WORKER)"
Write-Host "==============================================="

Stop-TargetPythonProcess -Needle "run_inss_app.py"
Stop-TargetPythonProcess -Needle "run_rq_worker.py"
Start-Sleep -Seconds 2

$redisReady = Test-TcpPort -HostName "127.0.0.1" -Port 6379
if (-not $redisReady) {
    Write-Host "Redis 127.0.0.1:6379 indisponivel. Tentando iniciar via Docker..."
    $containerLine = docker ps -a --format "{{.ID}}|{{.Names}}|{{.Ports}}|{{.Status}}" | Select-Object -First 1
    if ($containerLine) {
        $containers = docker ps -a --format "{{.ID}}|{{.Names}}|{{.Ports}}|{{.Status}}"
        $redisCandidate = $containers | Where-Object { $_ -match "6379->6379" } | Select-Object -First 1
        if ($redisCandidate) {
            $parts = $redisCandidate -split "\|"
            $redisName = $parts[1]
            Write-Host "Iniciando container Redis: $redisName"
            docker start $redisName | Out-Null
            Start-Sleep -Seconds 3
            $redisReady = Test-TcpPort -HostName "127.0.0.1" -Port 6379
        }
    }
}

if (-not $redisReady) {
    Write-Host "Redis continua indisponivel. API sera iniciada em fallback thread." -ForegroundColor Yellow
}

$queueBackend = if ($redisReady) { "rq" } else { "thread" }
$appScriptPath = Join-Path $PSScriptRoot "run_inss_app.py"
$workerScriptPath = Join-Path $PSScriptRoot "run_rq_worker.py"
$sessionSecret = [Environment]::GetEnvironmentVariable("APP_SESSION_SECRET", "Process")
if ([string]::IsNullOrWhiteSpace($sessionSecret)) {
    $sessionSecret = [Environment]::GetEnvironmentVariable("APP_SESSION_SECRET", "User")
}
if ([string]::IsNullOrWhiteSpace($sessionSecret)) {
    $sessionSecret = "trocar-antes-de-vps-2026-chave-segura-32chars"
}
$defaultAdminPassword = [Environment]::GetEnvironmentVariable("APP_DEFAULT_ADMIN_PASSWORD", "Process")
if ([string]::IsNullOrWhiteSpace($defaultAdminPassword)) {
    $defaultAdminPassword = [Environment]::GetEnvironmentVariable("APP_DEFAULT_ADMIN_PASSWORD", "User")
}
if ([string]::IsNullOrWhiteSpace($defaultAdminPassword)) {
    $defaultAdminPassword = "TrocarAdminForte2026!"
}

$apiCommand = @"
`$env:APP_ENV='production';
`$env:PGHOST='192.168.1.245';
`$env:PGPORT='5432';
`$env:PGDATABASE='inss_clientes';
`$env:PGUSER='postgres';
`$env:PGPASSWORD='651010';
`$env:APP_SESSION_SECRET='$sessionSecret';
`$env:APP_DEFAULT_ADMIN_PASSWORD='$defaultAdminPassword';
`$env:APP_HOST='0.0.0.0';
`$env:APP_PORT='8000';
`$env:APP_ENABLE_CEP_ENRICHMENT='1';
`$env:APP_ENABLE_DDD_ENRICHMENT='1';
`$env:APP_ENABLE_STARTUP_SCHEMA_MAINTENANCE='1';
`$env:APP_ENABLE_STARTUP_ADDRESS_BACKFILL='1';
`$env:APP_SKIP_RUNTIME_SCHEMA_VALIDATION='0';
`$env:C6_WORKER_ENABLED='1';
`$env:C6_USERNAME='01278252584_003238';
`$env:C6_PASSWORD='Brasil@2026';
`$env:APP_QUEUE_REDIS_URL='redis://127.0.0.1:6379/0';
`$env:APP_QUEUE_EXPORTS_NAME='inss_exports';
`$env:APP_QUEUE_IMPORTS_NAME='inss_imports';
`$env:APP_QUEUE_DEFAULT_NAME='inss_default';
`$env:APP_QUEUE_BACKEND='$queueBackend';
& 'python' '$appScriptPath';
"@
Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $apiCommand
Start-Sleep -Seconds 2

if ($redisReady) {
    $workerCommand = @"
`$env:APP_ENV='production';
`$env:PGHOST='192.168.1.245';
`$env:PGPORT='5432';
`$env:PGDATABASE='inss_clientes';
`$env:PGUSER='postgres';
`$env:PGPASSWORD='651010';
`$env:APP_SESSION_SECRET='$sessionSecret';
`$env:APP_DEFAULT_ADMIN_PASSWORD='$defaultAdminPassword';
`$env:APP_QUEUE_BACKEND='rq';
`$env:APP_QUEUE_REDIS_URL='redis://127.0.0.1:6379/0';
`$env:APP_QUEUE_EXPORTS_NAME='inss_exports';
`$env:APP_QUEUE_IMPORTS_NAME='inss_imports';
`$env:APP_QUEUE_DEFAULT_NAME='inss_default';
& 'python' '$workerScriptPath';
"@
    Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $workerCommand
}

$apiUp = $false
for ($i = 0; $i -lt 30; $i++) {
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:8000/health" -UseBasicParsing -TimeoutSec 2
        if ($response.StatusCode -eq 200) {
            $apiUp = $true
            break
        }
    } catch {
    }
    Start-Sleep -Seconds 1
}

$workerUp = $false
if ($redisReady) {
    for ($j = 0; $j -lt 20; $j++) {
        $workerUp = (Get-CimInstance Win32_Process -Filter "Name='python.exe'" | Where-Object {
            ($_.CommandLine -as [string]) -like "*run_rq_worker.py*"
        }).Count -gt 0
        if ($workerUp) {
            break
        }
        Start-Sleep -Seconds 1
    }
}

Write-Host ""
Write-Host "STATUS FINAL:"
Write-Host "API health: $apiUp"
Write-Host "Redis: $redisReady"
Write-Host "Worker RQ ativo: $workerUp"
Write-Host "==============================================="
