param(
    [string]$TaskName = "INSS_Housekeeping_Cleanup",
    [int]$RetentionDays = 90,
    [string]$DailyAt = "02:30"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = $PSScriptRoot
$scriptPath = Join-Path $projectRoot "housekeeping_cleanup.ps1"

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Script de cleanup nao encontrado: $scriptPath"
}

$parts = $DailyAt.Split(":")
if ($parts.Count -ne 2) {
    throw "Formato invalido para DailyAt. Use HH:mm, ex.: 02:30"
}

$hour = [int]$parts[0]
$minute = [int]$parts[1]
if ($hour -lt 0 -or $hour -gt 23 -or $minute -lt 0 -or $minute -gt 59) {
    throw "Horario DailyAt fora de faixa valida."
}

$triggerTime = Get-Date -Hour $hour -Minute $minute -Second 0
$actionArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -RetentionDays $RetentionDays"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $actionArgs -WorkingDirectory $projectRoot
$trigger = New-ScheduledTaskTrigger -Daily -At $triggerTime
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Limpeza automatica de data/archive e logs com retencao." `
    -Force | Out-Null

Write-Host "Task registrada: $TaskName"
Write-Host "Horario diario: $DailyAt"
Write-Host "Retencao: $RetentionDays dias"
Write-Host "Script: $scriptPath"
