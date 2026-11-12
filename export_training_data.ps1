# Fuehrt die gesammelten Trainingsdaten zusammen: ein Datensatz pro Mail.
# Zeigt ausserdem, wie oft die KI richtig lag und wie gut die Auto-Freigabe gewesen waere.
#
# Aufruf:
#   powershell -ExecutionPolicy Bypass -File export_training_data.ps1
#   powershell -ExecutionPolicy Bypass -File export_training_data.ps1 -Dir D:\daten\training_data -Out D:\export.jsonl
param(
    [string]$Dir = "",
    [string]$Out = ""
)
$scriptDir = $PSScriptRoot

if (-not $Dir) {
    $Dir = Join-Path $scriptDir "training_data"
    $cfgFile = Join-Path $scriptDir "config.json"
    if (Test-Path $cfgFile) {
        try {
            $cfg = Get-Content $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($cfg.training_data.dir) {
                $Dir = if ([IO.Path]::IsPathRooted("$($cfg.training_data.dir)")) { "$($cfg.training_data.dir)" } else { Join-Path $scriptDir "$($cfg.training_data.dir)" }
            }
        } catch {}
    }
}
if (-not (Test-Path $Dir)) { Write-Host "Ordner nicht gefunden: $Dir" -ForegroundColor Red; exit 1 }
if (-not $Out) { $Out = Join-Path $Dir ("export_" + (Get-Date -Format "yyyy-MM-dd") + ".jsonl") }

# -- Ereignisse einlesen ----------------------------------------------
$preds = @{}
$decs  = @{}
$order = New-Object System.Collections.ArrayList
$bad   = 0
$files = @(Get-ChildItem $Dir -Filter "emails_*.jsonl" | Sort-Object Name)
if ($files.Count -eq 0) { Write-Host "Keine Dateien emails_*.jsonl in $Dir" -ForegroundColor Yellow; exit 0 }

foreach ($f in $files) {
    foreach ($line in [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)) {
        if (-not $line.Trim()) { continue }
        try { $ev = $line | ConvertFrom-Json } catch { $bad++; continue }
        if (-not $ev.id) { $bad++; continue }
        if ($ev.event -eq 'prediction') {
            if (-not $preds.ContainsKey($ev.id)) { [void]$order.Add($ev.id) }
            $preds[$ev.id] = $ev
        } elseif ($ev.event -eq 'decision') {
            if (-not $decs.ContainsKey($ev.id)) { $decs[$ev.id] = New-Object System.Collections.ArrayList }
            [void]$decs[$ev.id].Add($ev)
        }
    }
}

# -- Zusammenfuehren --------------------------------------------------
$count  = [ordered]@{ gold = 0; silver = 0; rejected = 0; ignored = 0; open = 0 }
$labels = @{}
$evalAi     = @{ total = 0; ok = 0 }   # KI-Vorschlag (Kategorie) vs. deine Entscheidung
$evalHotel  = @{ total = 0; ok = 0 }   # Hotel bei Buchungen
$evalShadow = @{ total = 0; ok = 0 }   # "Wuerde automatisch ausgefuehrt" vs. deine Entscheidung
$evalOllama = @{ total = 0; ok = 0 }   # optionaler Ollama-Vergleich
$utf8 = New-Object Text.UTF8Encoding($false)
$sb = New-Object System.Text.StringBuilder

foreach ($id in $order) {
    $p    = $preds[$id]
    $list = if ($decs.ContainsKey($id)) { @($decs[$id]) } else { @() }
    if ($list.Count -eq 0) { $count.open++; continue }

    $last      = $list[$list.Count - 1]
    $corrected = @($list | Where-Object { $_.decision -eq 'corrected' }).Count -gt 0
    $label = $null; $hotel = $null
    switch ("$($last.decision)") {
        'approved'      { $quality = 'gold';   $label = $last.final_category; $hotel = $last.final_hotel }
        'corrected'     { $quality = 'gold';   $label = $last.final_category; $hotel = $last.final_hotel }
        'auto_approved' { $quality = 'silver'; $label = $last.final_category; $hotel = $last.final_hotel }
        'rejected'      { $quality = 'rejected' }
        'ignored'       { $quality = 'ignored' }
        default         { $quality = 'open' }
    }
    $count[$quality]++

    # Testbetrieb auswerten: haette die Auto-Freigabe das Richtige getan?
    if ($p.auto -and $p.auto.eligible -eq $true -and $quality -in @('gold','rejected','ignored')) {
        $evalShadow.total++
        if ($quality -eq 'gold' -and -not $corrected) { $evalShadow.ok++ }
    }

    if (-not $label) { continue }
    $labels[$label] = 1 + [int]$labels[$label]

    if ($quality -eq 'gold') {
        # Verglichen wird der Vorschlag des Bots (inkl. Sicherheitswarnungen) mit deiner Entscheidung
        if ($p.queued_category) {
            $evalAi.total++
            if ("$($p.queued_category)" -eq "$label" -and -not $corrected) { $evalAi.ok++ }
        }
        if ($label -eq 'booking' -and $hotel) {
            $evalHotel.total++
            if ("$($p.queued_hotel)" -eq "$hotel") { $evalHotel.ok++ }
        }
        if ($p.extra -and $p.extra.ollama_prediction) {
            $evalOllama.total++
            if ("$($p.extra.ollama_prediction.category)" -eq "$label") { $evalOllama.ok++ }
        }
    }

    $rec = [ordered]@{
        id                = $id
        message_id        = $p.input.message_id
        received          = $p.input.received
        input             = $p.input
        label             = $label
        hotel             = $hotel
        label_quality     = $quality
        was_corrected     = $corrected
        prediction        = $p.prediction
        ollama_prediction = $p.extra.ollama_prediction
        queued_category   = $p.queued_category
        auto              = $p.auto
        prompt_version    = $p.prompt_version
        decided_at        = $last.time
    }
    [void]$sb.Append((ConvertTo-Json -InputObject $rec -Depth 10 -Compress)).Append("`n")
}
[IO.File]::WriteAllText($Out, $sb.ToString(), $utf8)

# -- Auswertung --------------------------------------------------------
function Format-Rate($e) {
    if ($e.total -eq 0) { return "noch keine Daten" }
    return "{0} von {1} ({2:P0})" -f $e.ok, $e.total, ($e.ok / $e.total)
}

Write-Host ""
Write-Host "Trainingsdaten zusammengefuehrt" -ForegroundColor Green
Write-Host "  Datei:            $Out"
Write-Host "  Mails gesamt:     $($order.Count)"
Write-Host "  gold (von dir bestaetigt/korrigiert): $($count.gold)"
Write-Host "  silver (automatisch ausgefuehrt):     $($count.silver)"
Write-Host "  abgelehnt ohne Korrektur:             $($count.rejected)"
Write-Host "  ignoriert:                            $($count.ignored)"
Write-Host "  noch offen:                           $($count.open)"
if ($bad) { Write-Host "  unlesbare Zeilen:                     $bad" -ForegroundColor Yellow }
Write-Host ""
Write-Host "Datensaetze je Kategorie:"
foreach ($k in ($labels.Keys | Sort-Object)) { Write-Host ("  {0,-16} {1}" -f $k, $labels[$k]) }
Write-Host ""
Write-Host "Wie oft lag die KI richtig? (nur von dir gepruefte Mails)"
Write-Host "  Kategorie:              $(Format-Rate $evalAi)"
Write-Host "  Hotel bei Buchungen:    $(Format-Rate $evalHotel)"
Write-Host "  Auto-Freigabe (Test):   $(Format-Rate $evalShadow)"
if ($evalOllama.total -gt 0) { Write-Host "  Ollama zum Vergleich:   $(Format-Rate $evalOllama)" }
Write-Host ""
