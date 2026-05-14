$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$env:APP_ENV = "production"
$env:PGHOST = "192.168.1.245"
$env:PGPORT = "5432"
$env:PGDATABASE = "inss_clientes"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "651010"
$env:APP_QUEUE_BACKEND = "rq"
$env:APP_QUEUE_REDIS_URL = "redis://127.0.0.1:6379/0"
$env:APP_QUEUE_EXPORTS_NAME = "inss_exports"
$env:APP_QUEUE_IMPORTS_NAME = "inss_imports"
$env:APP_QUEUE_DEFAULT_NAME = "inss_default"

Write-Host "==============================================="
Write-Host " WORKER RQ - PRODUCAO"
Write-Host "==============================================="
Write-Host "APP_QUEUE_BACKEND=$env:APP_QUEUE_BACKEND"
Write-Host "APP_QUEUE_REDIS_URL=$env:APP_QUEUE_REDIS_URL"
Write-Host "APP_QUEUE_EXPORTS_NAME=$env:APP_QUEUE_EXPORTS_NAME"
Write-Host "APP_QUEUE_IMPORTS_NAME=$env:APP_QUEUE_IMPORTS_NAME"
Write-Host ""

$redisAvailable = $false
try {
    $tcp = New-Object System.Net.Sockets.TcpClient
    $iar = $tcp.BeginConnect("127.0.0.1", 6379, $null, $null)
    $ok = $iar.AsyncWaitHandle.WaitOne(1000, $false)
    if ($ok -and $tcp.Connected) {
        $redisAvailable = $true
    }
    $tcp.Close()
} catch {
    $redisAvailable = $false
}

if (-not $redisAvailable) {
    throw "Redis indisponivel em 127.0.0.1:6379. Inicie o Redis antes do worker RQ."
}

& "C:\Python312\python.exe" (Join-Path $PSScriptRoot "run_rq_worker.py")
