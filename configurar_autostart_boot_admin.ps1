param(
    [string]$TaskName = "INSS-Banco-Clientes-Prod-Boot",
    [string]$RepoPath = "\\Win-7bicsiu84ik\d\DADOS\PEDRO FILTRAGEM\ANALISE DE DADOS\BANCO_CLIENTES_INSS"
)

$ErrorActionPreference = "Stop"

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    throw "Execute este script em um PowerShell aberto como Administrador."
}

$pythonPath = Join-Path $RepoPath ".venv\Scripts\python.exe"
$appScript = Join-Path $RepoPath "run_inss_app.py"
$logsDir = Join-Path $RepoPath "logs"
$launcherPath = Join-Path $RepoPath "start_inss_prod_boot.bat"

if (-not (Test-Path -LiteralPath $RepoPath)) {
    throw "Pasta do repositorio nao encontrada: $RepoPath"
}
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Python da venv nao encontrado: $pythonPath"
}
if (-not (Test-Path -LiteralPath $appScript)) {
    throw "Script principal nao encontrado: $appScript"
}

New-Item -ItemType Directory -Path $logsDir -Force | Out-Null

$launcherContent = @"
@echo off
cd /d "$RepoPath"
"$pythonPath" "$appScript" >> "$logsDir\autostart.out.log" 2>> "$logsDir\autostart.err.log"
"@
Set-Content -LiteralPath $launcherPath -Value $launcherContent -Encoding ASCII

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute $launcherPath -WorkingDirectory $RepoPath
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Sobe o sistema INSS automaticamente no boot (antes do login)." | Out-Null

Start-ScheduledTask -TaskName $TaskName

Write-Host "Tarefa criada com sucesso."
Write-Host "TaskName: $TaskName"
Write-Host "Launcher: $launcherPath"
Write-Host "Logs: $logsDir\autostart.out.log | $logsDir\autostart.err.log"
Get-ScheduledTask -TaskName $TaskName | Select-Object TaskName, State, TaskPath | Format-List
Get-ScheduledTaskInfo -TaskName $TaskName | Select-Object LastRunTime, LastTaskResult, NextRunTime | Format-List
