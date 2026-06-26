# Startet eine eigene Ollama-Instanz auf Port 11435 fuer den Email-Bot
# Andere Programme nutzen weiterhin Port 11434 (Standard)

$model = "llama3.2"
$port  = 11435

# Pruefen ob bereits eine Instanz auf Port 11435 laeuft
$existing = netstat -ano | Select-String ":$port "
if ($existing) {
    Write-Host "Ollama Email-Bot laeuft bereits auf Port $port." -ForegroundColor Green
    Write-Host "Kein Neustart noetig."
    pause; exit 0
}

Write-Host ""
Write-Host "  Starte eigene Ollama-Instanz fuer Email-Bot..." -ForegroundColor Cyan
Write-Host "  Port: $port  (andere Programme nutzen 11434)"
Write-Host ""

# Ollama auf Port 11435 starten mit erhöhter CPU-Priorität
$env:OLLAMA_HOST = "127.0.0.1:$port"
$ollamaProc = Start-Process -FilePath "ollama" -ArgumentList "serve" -PassThru -WindowStyle Minimized
Start-Sleep -Seconds 1
try {
    $ollamaProc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal
    Write-Host "  Ollama gestartet (PID $($ollamaProc.Id)) — Priorität: AboveNormal" -ForegroundColor Green
} catch {
    Write-Host "  Ollama gestartet (PID $($ollamaProc.Id)) — Priorität konnte nicht gesetzt werden" -ForegroundColor Yellow
}

# Warten bis Ollama bereit ist
Write-Host "  Warte auf Ollama..." -NoNewline
$ready = $false
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 1
    try {
        $r = Invoke-RestMethod -Uri "http://localhost:$port/api/tags" -TimeoutSec 2 -ErrorAction Stop
        $ready = $true; break
    } catch {}
    Write-Host "." -NoNewline
}

if (-not $ready) {
    Write-Host "`n  FEHLER: Ollama hat nicht geantwortet." -ForegroundColor Red
    pause; exit 1
}

Write-Host "`n  Ollama bereit." -ForegroundColor Green

# Modell voraufladen damit die erste Email sofort klassifiziert wird
Write-Host "  Lade Modell '$model' vor..." -NoNewline
try {
    $body = @{ model=$model; messages=@(@{role="user";content="hi"}); stream=$false } | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Uri "http://localhost:$port/api/chat" -Method POST -Body $body -ContentType "application/json" -TimeoutSec 120 | Out-Null
    Write-Host " bereit." -ForegroundColor Green
} catch {
    Write-Host " Warnung: Vorladen fehlgeschlagen ($_)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Email-Bot Ollama laeuft auf http://localhost:$port" -ForegroundColor Green
Write-Host "  Dieses Fenster offen lassen - es haelt Ollama am Laufen."
Write-Host "  Zum Beenden: Strg+C"
Write-Host ""

# Fenster offen halten
try { while ($true) { Start-Sleep -Seconds 60 } }
finally {
    Write-Host "Stoppe Ollama Email-Bot..."
    Stop-Process -Id $ollamaProc.Id -Force -ErrorAction SilentlyContinue
}
