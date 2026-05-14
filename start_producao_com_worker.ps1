$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host "Iniciando API de producao em nova janela..."
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $PSScriptRoot "start_producao.ps1")

Start-Sleep -Seconds 2

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
    Write-Host "Iniciando worker RQ em nova janela..."
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $PSScriptRoot "start_worker_producao.ps1")
} else {
    Write-Host "Redis indisponivel. Worker nao iniciado; API segue em fallback thread." -ForegroundColor Yellow
}

Write-Host "API + worker disparados."
