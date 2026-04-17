$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$env:PGHOST = "192.168.1.245"
$env:PGPORT = "5432"
$env:PGDATABASE = "inss_clientes"
$env:PGUSER = "postgres"
$env:PGPASSWORD = "651010"

Write-Host "Ambiente de PRODUCAO"
Write-Host "PGHOST=$env:PGHOST"
Write-Host "PGPORT=$env:PGPORT"
Write-Host "PGDATABASE=$env:PGDATABASE"
Write-Host "PGUSER=$env:PGUSER"
Write-Host ""

& "C:\Python312\python.exe" ".\run_inss_app.py"
