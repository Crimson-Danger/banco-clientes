$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$env:APP_ENV = "test"
$env:PGHOST = "192.168.1.245"
$env:PGPORT = "5432"
$env:PGDATABASE = "inss_clientes_teste"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "651010"
$env:APP_HOST = "0.0.0.0"
$env:APP_PORT = "8001"

Write-Host "==============================================="
Write-Host " AMBIENTE: TESTE"
Write-Host "==============================================="
Write-Host "APP_ENV=$env:APP_ENV"
Write-Host "PGHOST=$env:PGHOST"
Write-Host "PGPORT=$env:PGPORT"
Write-Host "PGDATABASE=$env:PGDATABASE"
Write-Host "PGUSER=$env:PGUSER"
Write-Host "APP_PORT=$env:APP_PORT"
Write-Host ""

& "C:\Python312\python.exe" ".\run_inss_app.py"
