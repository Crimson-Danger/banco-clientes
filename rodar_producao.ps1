$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$env:PGHOST = "192.168.1.245"
$env:PGPORT = "5432"
$env:PGDATABASE = "inss_clientes_teste"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "651010"
$env:APP_HOST = "0.0.0.0"
$env:APP_PORT = "8010"

Write-Host "Ambiente de TESTE - simulacao de producao"
Write-Host "PGHOST=$env:PGHOST"
Write-Host "PGPORT=$env:PGPORT"
Write-Host "PGDATABASE=$env:PGDATABASE"
Write-Host "PGUSER=$env:PGUSER"
Write-Host "APP_PORT=$env:APP_PORT"
Write-Host "ROOT=$PSScriptRoot"
Write-Host ""

& "C:\Python312\python.exe" (Join-Path $PSScriptRoot "run_inss_app.py")
