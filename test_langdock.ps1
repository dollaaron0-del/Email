# Prueft die Langdock-Verbindung: API-Key, verfuegbare Modelle und eine Test-Analyse.
# Es wird KEINE echte E-Mail verschickt oder gelesen - nur ein Beispieltext an Langdock gesendet.
#
# Aufruf: powershell -ExecutionPolicy Bypass -File test_langdock.ps1
param([string]$ConfigFile = "")
$scriptDir = $PSScriptRoot
. (Join-Path $scriptDir "langdock.ps1")

if (-not $ConfigFile) {
    $ConfigFile = Join-Path $scriptDir "config.json"
    if (-not (Test-Path $ConfigFile)) { $ConfigFile = Join-Path $scriptDir "config.example.json" }
}
$config = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Host ""
Write-Host "Config: $ConfigFile"

# 1. API-Key
$keyName = Get-LangdockKeyName $config
$key = Get-LangdockApiKey $config
if (-not $key) {
    Write-Host "FEHLER: Kein API-Key in der Umgebungsvariable $keyName gefunden." -ForegroundColor Red
    Write-Host "Setzen mit (PowerShell):"
    Write-Host "  [Environment]::SetEnvironmentVariable('$keyName', 'DEIN-KEY', 'User')"
    exit 1
}
Write-Host "1) API-Key gefunden ($keyName)" -ForegroundColor Green

$ld = $config.langdock
$baseUrl = if ($ld.base_url) { "$($ld.base_url)".TrimEnd('/') } else { "https://api.langdock.com" }
$region  = if ($ld.region) { "$($ld.region)" } else { "eu" }

# 2. Modelle (nur ohne Agent)
if ($ld.agent_id) {
    Write-Host "2) Agent-Modus: Agent $($ld.agent_id) (Modellliste wird uebersprungen)"
} else {
    try {
        $resp = Invoke-WebRequest -Uri "$baseUrl/openai/$region/v1/models" -Method GET -UseBasicParsing -TimeoutSec 30 `
            -Headers @{ Authorization = "Bearer $key" }
        $models = ([Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray()) | ConvertFrom-Json).data | ForEach-Object { $_.id } | Sort-Object
        Write-Host "2) Verfuegbare Modelle:" -ForegroundColor Green
        $models | ForEach-Object { Write-Host "     $_" }
        if ($ld.model -and $ld.model -notin $models) {
            Write-Host "   WARNUNG: Das eingestellte Modell '$($ld.model)' ist nicht in der Liste." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "2) Modellliste nicht abrufbar: $(Get-HttpErrorDetail $_)" -ForegroundColor Yellow
    }
}

# 3. Test-Analyse mit einem erfundenen Beispiel
$hotelName = if ($config.hotels -and $config.hotels[0].name) { "$($config.hotels[0].name)" } else { "Beispiel Hotel" }
$hotels = ($config.hotels | ForEach-Object { "  - $($_.name) (Keywords: $($_.keywords -join ', '))" }) -join "`n"
$prompt = Get-AnalysisPrompt -HotelsBlock $hotels -ExamplesBlock "" `
    -SenderName "Max Mustermann" -SenderEmail "max.mustermann@example.com" `
    -Subject "Zimmeranfrage Maerz" -Attachments "keine" `
    -BodyText "Guten Tag,`n`nwir moechten vom 12. bis 14. Maerz ein Doppelzimmer im $hotelName buchen. Ist noch etwas frei?`n`nViele Gruesse`nMax Mustermann"

Write-Host "3) Test-Analyse (Buchungsanfrage fuer '$hotelName')..."
$sw = [Diagnostics.Stopwatch]::StartNew()
try {
    $raw = Invoke-LangdockJson -Prompt $prompt -Schema (Get-AnalysisSchema) -Config $config
    $res = ConvertTo-AnalysisResult $raw $config
    $sw.Stop()
    Write-Host ("   Antwort nach {0:N1} Sek.:" -f $sw.Elapsed.TotalSeconds) -ForegroundColor Green
    $res.GetEnumerator() | ForEach-Object { Write-Host ("     {0,-20} {1}" -f $_.Key, $_.Value) }
    if ($res.category -eq 'booking') { Write-Host "   OK - als Buchung erkannt." -ForegroundColor Green }
    else { Write-Host "   Hinweis: erwartet war 'booking'." -ForegroundColor Yellow }
} catch {
    Write-Host "   FEHLER: $(Get-HttpErrorDetail $_)" -ForegroundColor Red
    exit 1
}
Write-Host ""
