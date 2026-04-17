$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Ambiente virtual nao encontrado. Crie com: python -m venv .venv"
    exit 1
}

$env:PGHOST = Read-Host "Host PostgreSQL"
if (-not $env:PGHOST) { throw "PGHOST obrigatorio." }

$port = Read-Host "Porta PostgreSQL [5432]"
$env:PGPORT = if ($port) { $port } else { "5432" }

$db = Read-Host "Nome do banco novo [inss_clientes]"
$env:PGDATABASE = if ($db) { $db } else { "inss_clientes" }

$user = Read-Host "Usuario administrador [postgres]"
$env:PGADMINUSER = if ($user) { $user } else { "postgres" }

$adminPassword = Read-Host "Senha administrador" -AsSecureString
$env:PGADMINPASSWORD = [System.Net.NetworkCredential]::new("", $adminPassword).Password

& ".venv\Scripts\python.exe" ".\init_postgres_db.py"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$appUser = Read-Host "Usuario da aplicacao [$env:PGADMINUSER]"
$env:PGUSER = if ($appUser) { $appUser } else { $env:PGADMINUSER }

$useSamePassword = Read-Host "Usar a mesma senha do admin para a aplicacao? [S/n]"
if ($useSamePassword -match "^(n|N)$") {
    $appPassword = Read-Host "Senha da aplicacao" -AsSecureString
    $env:PGPASSWORD = [System.Net.NetworkCredential]::new("", $appPassword).Password
} else {
    $env:PGPASSWORD = $env:PGADMINPASSWORD
}

Write-Host ""
Write-Host "Variaveis definidas para esta sessao:"
Write-Host "PGHOST=$env:PGHOST"
Write-Host "PGPORT=$env:PGPORT"
Write-Host "PGUSER=$env:PGUSER"
Write-Host "PGDATABASE=$env:PGDATABASE"
Write-Host ""
Write-Host "Para subir a aplicacao nesta mesma janela:"
Write-Host '.\.venv\Scripts\python.exe .\run_inss_app.py'
