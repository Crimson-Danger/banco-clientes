$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$env:APP_ENV = "production"
$env:PGHOST = "192.168.1.245"
$env:PGPORT = "5432"
$env:PGDATABASE = "inss_clientes"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "651010"
$env:APP_HOST = "0.0.0.0"
$env:APP_PORT = "8000"
$env:APP_ENABLE_CEP_ENRICHMENT = "1"
$env:APP_ENABLE_DDD_ENRICHMENT = "1"
$env:APP_ENABLE_STARTUP_SCHEMA_MAINTENANCE = "0"
$env:APP_ENABLE_STARTUP_ADDRESS_BACKFILL = "1"
$env:APP_SKIP_RUNTIME_SCHEMA_VALIDATION = "0"
$env:C6_WORKER_ENABLED = "1"
$env:C6_USERNAME = "01278252584_003238"
$env:C6_PASSWORD = "Brasil@2026"
$env:APP_QUEUE_REDIS_URL = "redis://127.0.0.1:6379/0"
$env:APP_QUEUE_EXPORTS_NAME = "inss_exports"
$env:APP_QUEUE_IMPORTS_NAME = "inss_imports"
$env:APP_QUEUE_DEFAULT_NAME = "inss_default"

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

if ($redisAvailable) {
    $env:APP_QUEUE_BACKEND = "rq"
} else {
    $env:APP_QUEUE_BACKEND = "thread"
}

if ($env:PGDATABASE -match "(?i)test|homolog") {
    throw "Bloqueado: PGDATABASE de producao nao pode conter 'test' ou 'homolog'."
}

Write-Host "==============================================="
Write-Host " AMBIENTE: PRODUCAO"
Write-Host "==============================================="
Write-Host "APP_ENV=$env:APP_ENV"
Write-Host "PGHOST=$env:PGHOST"
Write-Host "PGPORT=$env:PGPORT"
Write-Host "PGDATABASE=$env:PGDATABASE"
Write-Host "PGUSER=$env:PGUSER"
Write-Host "APP_PORT=$env:APP_PORT"
Write-Host "APP_ENABLE_CEP_ENRICHMENT=$env:APP_ENABLE_CEP_ENRICHMENT"
Write-Host "APP_ENABLE_DDD_ENRICHMENT=$env:APP_ENABLE_DDD_ENRICHMENT"
Write-Host "APP_ENABLE_STARTUP_SCHEMA_MAINTENANCE=$env:APP_ENABLE_STARTUP_SCHEMA_MAINTENANCE"
Write-Host "APP_ENABLE_STARTUP_ADDRESS_BACKFILL=$env:APP_ENABLE_STARTUP_ADDRESS_BACKFILL"
Write-Host "APP_QUEUE_BACKEND=$env:APP_QUEUE_BACKEND"
if ($env:APP_QUEUE_BACKEND -eq "thread") {
    Write-Host "REDIS indisponivel. API segue em fallback seguro (thread)." -ForegroundColor Yellow
}
Write-Host "ROOT=$PSScriptRoot"
Write-Host ""

$confirm = Read-Host "Digite PRODUCAO para confirmar subida no ambiente real"
if ($confirm -cne "PRODUCAO") {
    throw "Operacao cancelada: confirmacao invalida."
}

& "C:\Python312\python.exe" (Join-Path $PSScriptRoot "run_inss_app.py")
