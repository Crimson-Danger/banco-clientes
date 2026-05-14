$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

function Get-RequiredEnvValue {
    param([string]$Name)
    $value = [Environment]::GetEnvironmentVariable($Name, "Process")
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($Name, "User")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable($Name, "Machine")
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Variavel obrigatoria ausente: $Name"
    }
    return $value
}

$env:PGHOST = Get-RequiredEnvValue "PGHOST"
$env:PGPORT = Get-RequiredEnvValue "PGPORT"
$env:PGDATABASE = Get-RequiredEnvValue "PGDATABASE"
$env:PGUSER = Get-RequiredEnvValue "PGUSER"
$env:PGPASSWORD = Get-RequiredEnvValue "PGPASSWORD"

$scriptPath = Join-Path $PSScriptRoot "sql\phase1_postgres_indexes.sql"
if (-not (Test-Path $scriptPath)) {
    throw "Script nao encontrado: $scriptPath"
}

Write-Host "Aplicando indices Fase 1 em $env:PGHOST:$env:PGPORT/$env:PGDATABASE ..."
& psql -h $env:PGHOST -p $env:PGPORT -U $env:PGUSER -d $env:PGDATABASE -f $scriptPath
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao aplicar indices (psql exit code $LASTEXITCODE)."
}

Write-Host "Indices aplicados com sucesso."
