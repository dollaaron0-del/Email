# Watchdog: ueberwacht Bot und Dashboard, startet sie automatisch neu

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logFile   = Join-Path $scriptDir "watchdog.log"
$botScript = Join-Path $scriptDir "bot.ps1"
$dashScript = Join-Path $scriptDir "dashboard_server.ps1"

function Write-WatchLog {
    param([string]$Msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Msg"
    Add-Content $logFile -Value $line -Encoding UTF8
}

function Test-Port {
    param([int]$Port)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect("127.0.0.1", $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(1000)
        $tcp.Close()
        return $ok
    } catch { return $false }
}

function Get-BotProcess {
    # Eigene PID ausschliessen damit der Watchdog sich nicht selbst matcht
    $myPid = $PID
    return Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessId -ne $myPid -and $_.CommandLine -like "*-File*bot.ps1*" -and $_.CommandLine -notlike "*watchdog*" }
}

Write-WatchLog "Watchdog gestartet (PID $PID)"

while ($true) {
    Start-Sleep -Seconds 60

    # Dashboard pruefen (Port 5000)
    if (-not (Test-Port 5000)) {
        Write-WatchLog "Dashboard nicht erreichbar - starte neu"
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$dashScript`"" -WorkingDirectory $scriptDir
        Start-Sleep -Seconds 6
    }

    # Bot pruefen
    if (-not (Get-BotProcess)) {
        Write-WatchLog "Bot nicht aktiv - starte neu"
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$botScript`"" -WorkingDirectory $scriptDir
        Start-Sleep -Seconds 6
    }
}
