# Email-Processor Bot - PowerShell + Outlook + Ollama
# Emails werden NICHT automatisch versendet - zuerst im Dashboard genehmigen!
$ErrorActionPreference = "Continue"
$scriptDir       = $PSScriptRoot
$correctionsFile = Join-Path $scriptDir "corrections.json"

# Bot-Prozess auf hohe Priorität setzen
try { [System.Diagnostics.Process]::GetCurrentProcess().PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal } catch {}

# ── Logging ──────────────────────────────────────────────────────────
$logFile = Join-Path $scriptDir "email_processor.log"
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $($Level.PadRight(8)) $Msg"
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line
}

# ── Config ────────────────────────────────────────────────────────────
function Load-Config {
    Get-Content (Join-Path $scriptDir "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
}

# ── Korrekturen (Few-Shot Learning) ──────────────────────────────────
function Load-Corrections {
    if (-not (Test-Path $correctionsFile)) { return @() }
    try {
        $parsed = Get-Content $correctionsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($parsed -is [System.Collections.IEnumerable]) { return @($parsed) }
        return @($parsed)
    } catch { return @() }
}

function Save-Correction {
    param($Entry)
    $existing = @(Load-Corrections)
    $domain = if ($Entry.sender -match '@(.+)$') { $Matches[1] } else { $Entry.sender }
    $new = [PSCustomObject]@{
        time             = Get-Date -Format "yyyy-MM-dd HH:mm"
        subject          = $Entry.subject
        sender_domain    = $domain
        wrong_category   = $Entry.category
        correct_category = $Entry.correct_category
        correct_hotel    = if ($Entry.correct_hotel) { $Entry.correct_hotel } else { "" }
    }
    $all = @($new) + $existing | Select-Object -First 50
    $parts = @($all | ForEach-Object { ConvertTo-Json -InputObject $_ -Depth 4 -Compress })
    $json = "[`n" + ($parts -join ",`n") + "`n]"
    [IO.File]::WriteAllText($correctionsFile, $json, [Text.Encoding]::UTF8)
    Write-Log "  -> Korrektur gespeichert: $($Entry.category) -> $($Entry.correct_category)"
}

# ── Outlook ───────────────────────────────────────────────────────────
function Get-OutlookApp {
    try { [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") }
    catch { New-Object -ComObject Outlook.Application }
}

function Get-OrCreate-Folder {
    param($Parent, [string]$Name)
    foreach ($f in $Parent.Folders) { if ($f.Name -eq $Name) { return $f } }
    $Parent.Folders.Add($Name)
}

function Resolve-Folder {
    param($Inbox, [string]$Path)
    $f = $Inbox
    foreach ($part in $Path.Split('/')) { $f = Get-OrCreate-Folder $f $part }
    $f
}

function Get-UnsubscribeHeader {
    param($Item)
    try {
        $h = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
        if ($h -match "(?i)List-Unsubscribe:\s*(.+?)(\r?\n\S|\z)") { return $Matches[1].Trim() }
    } catch {}
    ""
}

function Find-Hotel {
    param([string]$HotelName, $Config)
    if (-not $HotelName -or $HotelName -in @('null','')) { return $null }
    foreach ($h in $Config.hotels) {
        if ($h.name -like "*$HotelName*" -or $HotelName -like "*$($h.name)*") { return $h }
    }
    foreach ($h in $Config.hotels) {
        foreach ($kw in $h.keywords) { if ($HotelName -like "*$kw*") { return $h } }
    }
    $null
}

# ── Ollama ────────────────────────────────────────────────────────────
function Classify-Email {
    param($Item, $Config)
    $hotels = ($Config.hotels | ForEach-Object { "  - $($_.name) (Keywords: $($_.keywords -join ', '))" }) -join "`n"

    # Korrekturen dedupliziert laden: max. 2 Beispiele pro Muster (wrong->correct), bis zu 15 gesamt
    $allCorrections = @(Load-Corrections)
    $seen = @{}
    $exampleSet = @()
    foreach ($c in $allCorrections) {
        $key = "$($c.wrong_category)->$($c.correct_category)"
        if (-not $seen.ContainsKey($key)) { $seen[$key] = 0 }
        if ($seen[$key] -lt 2) { $exampleSet += $c; $seen[$key]++ }
        if ($exampleSet.Count -ge 15) { break }
    }
    $examplesBlock = ""
    if ($exampleSet.Count -gt 0) {
        $lines = $exampleSet | ForEach-Object {
            $hotel = if ($_.correct_hotel) { " -> Hotel: $($_.correct_hotel)" } else { "" }
            "  - Betreff: `"$($_.subject)`" | Domain: $($_.sender_domain) | War: $($_.wrong_category) | Richtig: $($_.correct_category)$hotel"
        }
        $examplesBlock = "`nLernbeispiele (HOHE PRIORITAET - diese Muster wurden korrigiert):`n" + ($lines -join "`n") + "`n"
    }

    function Invoke-Classify($bodyText) {
        $prompt = @"
Klassifiziere diese E-Mail im Kontext Hotelmanagement.

Kategorien (waehle EXAKT eine):
- booking: Buchungsanfragen, Zimmerverfuegbarkeit, Reservierungen, Gruppenanfragen - AUCH von privaten Adressen wie gmail.com, hotmail.com etc.
- application: Bewerbungen, Lebenslauf-Einsendungen, Jobanfragen - erkennbar an Karrieredomains (stepstone, linkedin, hotelpersonal, dki-lehrlinge) ODER expliziten Bewerbungsformulierungen
- newsletter: Regelmaessige Rundschreiben, Magazine, Branchennews, oft mit Abmelde-Link
- advertisement: Werbung, Verkaufsangebote, B2B-Marketing, Dienstleistungsangebote, Produktpraesentation
- phishing: Verdacht auf Phishing/Betrug/Schadsoftware - Aufforderung zur Eingabe von Zugangsdaten, Kontosperrung-Drohungen, gefaelschte Rechnungen, CEO-Fraud, ungewoehnliche Zahlungsaufforderungen, Identitaetsdiebstahl-Versuche, verdaechtige Anhaenge/Links
- sonstiges: Wichtige Geschaeftskorrespondenz, Partneranfragen, Kooperationen, externe Emails die in keine andere Kategorie passen

Hotels:
$hotels
$examplesBlock
Email:
Von: $($Item.SenderEmailAddress) | Betreff: $($Item.Subject)
Text: $bodyText

Antworte NUR mit JSON: {"category":"...","hotel_name":null,"confidence":"high","reason":"..."}
"@
        $reqBody = @{
            model = $Config.ollama_model
            messages = @(@{ role = "user"; content = $prompt })
            stream = $false; format = "json"; options = @{ temperature = 0.1 }
        } | ConvertTo-Json -Depth 10 -Compress

        $ollamaUrl = if ($Config.ollama_url) { $Config.ollama_url } else { "http://localhost:11434/api/chat" }
        $resp = Invoke-RestMethod -Uri $ollamaUrl -Method POST `
            -Body ([Text.Encoding]::UTF8.GetBytes($reqBody)) `
            -ContentType "application/json; charset=utf-8" -TimeoutSec 180
        return $resp.message.content | ConvertFrom-Json
    }

    try {
        # Erst mit kuerzem Text versuchen (schneller, weniger Timeouts)
        $body = if ($Item.Body.Length -gt 800) { $Item.Body.Substring(0,800) } else { $Item.Body }
        $r = Invoke-Classify $body
        $cat = $r.category
        if ($cat -notin @('newsletter','application','advertisement','booking','phishing','sonstiges','unknown')) { $cat = 'unknown' }
        Write-Log "  -> $cat | Hotel: $($r.hotel_name) | $($r.confidence) | $($r.reason)"
        return @{ category = $cat; hotel = "$($r.hotel_name)"; reason = "$($r.reason)" }
    } catch {
        Write-Log "  -> Erster Versuch fehlgeschlagen, retry mit laengerem Text..." "WARN"
        try {
            $body = if ($Item.Body.Length -gt 2000) { $Item.Body.Substring(0,2000) } else { $Item.Body }
            $r = Invoke-Classify $body
            $cat = $r.category
            if ($cat -notin @('newsletter','application','advertisement','booking','phishing','sonstiges','unknown')) { $cat = 'unknown' }
            Write-Log "  -> $cat | Hotel: $($r.hotel_name) | $($r.confidence) | $($r.reason)"
            return @{ category = $cat; hotel = "$($r.hotel_name)"; reason = "$($r.reason)" }
        } catch {
            Write-Log "  -> Klassifizierungsfehler: $_" "ERROR"
            return @{ category = 'unknown'; hotel = $null }
        }
    }
}

# ── Rückfrage-Entwurf generieren ─────────────────────────────────────
function Detect-Language {
    param($Item, $Config)
    $snippet = if ($Item.Body.Length -gt 400) { $Item.Body.Substring(0,400) } else { $Item.Body }
    $prompt  = "Detect the language of this email. Reply with ONLY one of these values: de, en, es`n`nSubject: $($Item.Subject)`nText: $snippet"
    try {
        $reqBody = @{
            model    = $Config.ollama_model
            messages = @(@{ role = "user"; content = $prompt })
            stream   = $false
            options  = @{ temperature = 0.0 }
        } | ConvertTo-Json -Depth 10 -Compress
        $ollamaUrl = if ($Config.ollama_url) { $Config.ollama_url } else { "http://localhost:11434/api/chat" }
        $resp = Invoke-RestMethod -Uri $ollamaUrl -Method POST `
            -Body ([Text.Encoding]::UTF8.GetBytes($reqBody)) `
            -ContentType "application/json; charset=utf-8" -TimeoutSec 120
        $lang = ($resp.message.content.Trim().ToLower() -replace '[^a-z]','').Substring(0,2)
        if ($lang -in @('de','en','es')) { return $lang }
    } catch {}
    return "de"
}

function Generate-Rueckfrage {
    param($Item, $Config)

    # 1. Sprache erkennen
    $lang = Detect-Language $Item $Config
    Write-Log "  -> Sprache erkannt: $lang"

    # 2. Passendes Template aus Config laden
    $templates = $Config.reply_templates
    $template  = if ($templates -and $templates.$lang) { "$($templates.$lang)" } `
                 elseif ($templates -and $templates.de)  { "$($templates.de)" } `
                 else { "Sehr geehrte Damen und Herren,`n`nfuer welches Hotel gilt Ihre Anfrage?`n`n{hotels}`n`nMit freundlichen Gruessen" }

    # 3. {hotels} Platzhalter ersetzen
    $hotelList = ($Config.hotels | ForEach-Object { "- $($_.name)" }) -join "`n"
    return $template -replace '\{hotels\}', $hotelList
}

# ── Queue ─────────────────────────────────────────────────────────────
$queueFile = Join-Path $scriptDir "queue.json"

function Flatten-Queue {
    param($obj)
    # Rekursiv alle echten Queue-Eintraege (mit "id"-Feld) aus verschachtelten value-Strukturen holen
    $results = @()
    if ($null -eq $obj) { return $results }
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        foreach ($item in $obj) { $results += Flatten-Queue $item }
    } elseif ($obj.PSObject.Properties.Name -contains 'id') {
        $results += $obj
    } elseif ($obj.PSObject.Properties.Name -contains 'value') {
        $results += Flatten-Queue $obj.value
    }
    return $results
}

function Load-Queue {
    if (-not (Test-Path $queueFile)) { return @() }
    try {
        $parsed = Get-Content $queueFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $flat = @(Flatten-Queue $parsed)
        # Duplikate nach ID entfernen
        $seen = @{}
        return @($flat | Where-Object { $_.id -and -not $seen.ContainsKey($_.id) } | ForEach-Object { $seen[$_.id] = 1; $_ })
    }
    catch { return @() }
}

function Save-Queue {
    param($Q)
    # Jedes Item einzeln serialisieren - vermeidet PowerShell 5.1 Array-Serialisierungsbug
    $parts = @($Q | Where-Object { $null -ne $_ } | ForEach-Object {
        ConvertTo-Json -InputObject $_ -Depth 8 -Compress
    })
    $json = if ($parts.Count -eq 0) { '[]' } else { "[`n" + ($parts -join ",`n") + "`n]" }
    [IO.File]::WriteAllText($queueFile, $json, [Text.Encoding]::UTF8)
}

function Add-To-Queue {
    param($Entry)
    $existing = @(Load-Queue)
    $all = $existing + @($Entry)
    Save-Queue $all
}

function Build-QueueEntry {
    param($Item, [string]$Category, [string]$Hotel, $Config)

    $entry = [ordered]@{
        id          = [System.Guid]::NewGuid().ToString()
        entry_id    = $Item.EntryID
        subject     = if ($Item.Subject.Length -gt 80) { $Item.Subject.Substring(0,80) } else { $Item.Subject }
        sender      = $Item.SenderEmailAddress
        sender_name = $Item.SenderName
        received    = $Item.ReceivedTime.ToString("dd.MM.yyyy HH:mm")
        created     = (Get-Date -Format "dd.MM. HH:mm")
        category    = $Category
        hotel       = $Hotel
        status      = "pending"
        proposed_action  = ""
        target_folder    = "EmailProcessor/Fehler"
        target_email     = ""
        unsubscribe_header = ""
    }

    switch ($Category) {
        'newsletter' {
            $h = Get-UnsubscribeHeader $Item
            $entry.unsubscribe_header = $h
            if ($h -match '<mailto:') {
                $entry.unsubscribe_method = "mailto"
                $entry.proposed_action    = "Abmelde-Email senden"
                $entry.target_folder      = "EmailProcessor/Abgemeldet"
            } elseif ($h -match '<https?://') {
                $entry.unsubscribe_method = "http"
                $entry.proposed_action    = "Abmelde-Webseite aufrufen + Button klicken"
                $entry.target_folder      = "EmailProcessor/Abgemeldet"
            } else {
                $entry.unsubscribe_method = "none"
                $entry.proposed_action    = "Kein Abmelde-Link - manuell abmelden"
                $entry.target_folder      = "EmailProcessor/Newsletter-Manuell"
            }
        }
        'application' {
            $entry.proposed_action = "Weiterleiten an: $($Config.applications_email)"
            $entry.target_email    = $Config.applications_email
            $entry.target_folder   = "EmailProcessor/Bewerbung"
        }
        'advertisement' {
            $entry.proposed_action = "Archivieren - keine weitere Aktion"
            $entry.target_folder   = "EmailProcessor/Werbung"
        }
        'booking' {
            $h = Find-Hotel $Hotel $Config
            if ($h) {
                $entry.proposed_action = "Weiterleiten an $($h.name) ($($h.email))"
                $entry.action_type     = "forward"
                $entry.target_email    = $h.email
                $entry.target_folder   = "EmailProcessor/Buchung/$($h.name)"
                $entry.hotel           = $h.name
            } else {
                Write-Log "  -> Hotel unklar, generiere Rueckfrage-Entwurf..."
                $draft = Generate-Rueckfrage $Item $Config
                $entry.proposed_action = "Rueckfrage an Gast senden (Hotel nicht erkennbar)"
                $entry.action_type     = "reply"
                $entry.reply_draft     = $draft
                $entry.reply_to        = $Item.SenderEmailAddress
                $entry.target_folder   = "EmailProcessor/Buchung-Unbekannt"
            }
        }
        'sonstiges' {
            $entry.proposed_action = "Wichtige Geschäftskorrespondenz - manuell bearbeiten"
            $entry.target_folder   = "EmailProcessor/Sonstiges"
        }
        'phishing' {
            $entry.proposed_action = "ACHTUNG: Verdacht auf Phishing/Schadsoftware - Links/Anhänge NICHT öffnen. Bitte prüfen."
            $entry.target_folder   = "EmailProcessor/Sicherheitswarnung"
        }
        default {
            $entry.proposed_action = "Kategorie unbekannt - manuell prüfen"
            $entry.target_folder   = "EmailProcessor/Unbekannt"
        }
    }
    return $entry
}

# ── Verdächtige Email isolieren (Phishing/Schadsoftware) ─────────────
function Quarantine-Suspicious-Email {
    param($Item, $Inbox, [string[]]$Reasons, $Config)
    try { $Item.Importance = 2; $Item.FlagStatus = 2; $Item.Categories = "Phishing-Verdacht"; $Item.Save() } catch {}

    $entry = Build-QueueEntry $Item 'phishing' '' $Config
    $reasonText = if ($Reasons -and $Reasons.Count -gt 0) { $Reasons -join " | " } else { "Verdächtige Merkmale erkannt" }
    $entry.proposed_action = "ACHTUNG: Verdacht auf Phishing/Schadsoftware - Links/Anhänge NICHT öffnen. Gründe: $reasonText"
    $entry.security_reasons = $Reasons

    $moved = $Item.Move((Resolve-Folder $Inbox "EmailProcessor/Sicherheitswarnung"))
    $entry.entry_id = $moved.EntryID
    Add-To-Queue $entry
    return $entry
}

# ── Aktion ausführen (nach Genehmigung) ──────────────────────────────
function Invoke-WebUnsubscribe {
    param([string]$Url)
    try {
        # Schritt 1: Seite laden
        $resp = Invoke-WebRequest -Uri $Url -Method GET -TimeoutSec 20 -UseBasicParsing -SessionVariable sess
        $html = $resp.Content

        # Schritt 2: Direkte Bestätigung durch GET?
        if ($html -match '(?i)(erfolgreich abgemeldet|unsubscribed|successfully removed|opt.out successful|abmeldung erfolgreich|you.ve been unsubscribed|you have been unsubscribed)') {
            return @{ method = "get"; message = "Seite bestätigt Abmeldung direkt (kein Klick nötig)" }
        }

        # Schritt 3: Abmelde-Formular auf der Seite suchen und abschicken
        $formPattern = "(?is)<form[^>]*?action=""?([^\s"">]+)""?[^>]*?>(.*?)</form>"
        $formMatches = [regex]::Matches($html, $formPattern)
        foreach ($fm in $formMatches) {
            $action  = $fm.Groups[1].Value
            $content = $fm.Groups[2].Value
            if ($content -notmatch '(?i)(unsubscribe|abmeld|opt.?out|austragen|confirm|remove)') { continue }

            # Absolute URL aufbauen
            $base = [Uri]$Url
            $actionUrl = if ($action -match '^https?://') { $action }
                         elseif ($action.StartsWith('/'))  { "$($base.Scheme)://$($base.Host)$action" }
                         else                              { "$($base.Scheme)://$($base.Host)/$action" }

            # Alle Input-Felder sammeln (hidden + submit)
            $fields = @{}
            $inputPattern = "(?i)<input([^>]+)>"
            [regex]::Matches($content, $inputPattern) | ForEach-Object {
                $inp   = $_.Groups[1].Value
                $type  = if ($inp -match '(?i)type="?(\w+)') { $Matches[1] } else { "text" }
                $name  = if ($inp -match '(?i)name="([^">\s]+)"') { $Matches[1] } else { $null }
                $value = if ($inp -match '(?i)value="([^"]*)"')    { $Matches[1] } else { "" }
                if ($name -and $type -in @('hidden','submit','email','text')) { $fields[$name] = $value }
            }

            $postResp = Invoke-WebRequest -Uri $actionUrl -Method POST -Body $fields -TimeoutSec 20 -UseBasicParsing -WebSession $sess
            if ($postResp.Content -match '(?i)(erfolgreich|unsubscribed|success|removed|abgemeldet)') {
                return @{ method = "post"; message = "Formular abgeschickt - Seite bestaetigt Abmeldung" }
            }
            return @{ method = "post"; message = "Formular abgeschickt (keine Bestaetigung erkannt - ggf. manuell pruefen)" }
        }

        # Schritt 4: Kein Formular - nur GET war moeglich
        return @{ method = "get"; message = "Seite aufgerufen, kein Abmelde-Formular gefunden - ggf. manuell pruefen" }
    } catch {
        return @{ method = "error"; message = "Fehler beim Aufrufen der Abmelde-Seite: $_" }
    }
}

function Execute-Entry {
    param($Entry, $Namespace, $Inbox, $Outlook)

    $item = $null
    try { $item = $Namespace.GetItemFromID($Entry.entry_id) } catch {}
    if (-not $item) {
        # Fallback: in EmailProcessor/Zur Pruefung nach Betreff + Absender suchen
        try {
            $pruefOrdner = Resolve-Folder $Inbox "EmailProcessor/Zur Pruefung"
            foreach ($m in $pruefOrdner.Items) {
                try {
                    $senderMatch = ($m.SenderEmailAddress -eq $Entry.sender) -or ($m.SenderName -eq $Entry.sender_name)
                    if ($m.Subject -like "*$($Entry.subject.Substring(0,[Math]::Min(40,$Entry.subject.Length)))*" -and $senderMatch) { $item = $m; break }
                } catch {}
            }
        } catch {}
    }
    if (-not $item) { Write-Log "  -> E-Mail nicht mehr auffindbar (EntryID ungültig)" "WARN"; return "error" }

    try {
        switch ($Entry.category) {
            'newsletter' {
                if ($Entry.unsubscribe_header) {
                    $h = $Entry.unsubscribe_header
                    if ($h -match '<mailto:([^>?]+)') {
                        $addr = $Matches[1]; $subj = "unsubscribe"
                        if ($h -match 'subject=([^&>]+)') { $subj = $Matches[1] }
                        $m = $Outlook.CreateItem(0); $m.To = $addr; $m.Subject = $subj; $m.Body = "unsubscribe"; $m.Send()
                        Write-Log "  -> Abmelde-Email gesendet an $addr"
                    } elseif ($h -match '<(https?://[^>]+)>') {
                        $result = Invoke-WebUnsubscribe $Matches[1]
                        Write-Log "  -> Web-Abmeldung ($($result.method)): $($result.message)"
                    }
                }
            }
            'application' {
                $fwd = $item.Forward()
                $fwd.To = $Entry.target_email
                $fwd.Subject = "BEWERBUNG: $($item.Subject)"
                $fwd.Body = "Weitergeleitet von: $($item.SenderEmailAddress)`n$("`-"*60)`n" + $fwd.Body
                $fwd.Send()
            }
            'booking' {
                if ($Entry.action_type -eq 'reply') {
                    # Rueckfrage als echtes Reply auf die Originalemail senden
                    $replyText = if ($Entry.reply_text) { $Entry.reply_text } else { $Entry.reply_draft }
                    $replyTo   = if ($Entry.reply_to)  { $Entry.reply_to  } else { $Entry.sender }
                    $mail = $item.Reply()
                    $mail.Body = $replyText + "`n`n" + $mail.Body
                    $mail.Send()
                    Write-Log "  -> Rueckfrage gesendet an $replyTo"
                } elseif ($Entry.target_email) {
                    # Weiterleitung an Hotel
                    $fwd = $item.Forward()
                    $fwd.To = $Entry.target_email
                    $fwd.Subject = "BUCHUNGSANFRAGE - $($Entry.hotel): $($item.Subject)"
                    $fwd.Body = "Von: $($item.SenderEmailAddress)`n$("`-"*60)`n" + $fwd.Body
                    $fwd.Send()
                }
            }
            'sonstiges' {
                # Rot markieren + hohe Priorität damit die Email sofort auffällt
                try { $item.Importance = 2; $item.FlagStatus = 2; $item.Save() } catch {}
            }
            # 'advertisement' und andere: keine Sendung, nur verschieben
        }
        [void]$item.Move((Resolve-Folder $Inbox $Entry.target_folder))
        Write-Log "  -> Ausgefuehrt: $($Entry.proposed_action)"
        return "done"
    } catch {
        Write-Log "  -> Fehler beim Ausfuehren: $_" "ERROR"
        return "error"
    }
}

# ── Korrigierten Queue-Eintrag aufbauen ──────────────────────────────
function Build-CorrectedEntry {
    param($Entry, $Config)
    $corr = @{}
    $Entry.PSObject.Properties | ForEach-Object { $corr[$_.Name] = $_.Value }
    $corr.category   = $Entry.correct_category
    $corr.hotel      = if ($Entry.correct_hotel) { $Entry.correct_hotel } else { "" }
    $corr.action_type = ""

    switch ($Entry.correct_category) {
        'booking' {
            $h = if ($Entry.correct_hotel) { Find-Hotel $Entry.correct_hotel $Config } else { $null }
            if ($h) {
                $corr.target_folder   = "EmailProcessor/Buchung/$($h.name)"
                $corr.target_email    = $h.email
                $corr.proposed_action = "Weitergeleitet an $($h.name)"
                $corr.action_type     = "forward"
                $corr.hotel           = $h.name
            } else {
                # Kein Hotel bekannt -> Rueckfrage an Gast, genau wie beim normalen Buchungsflow
                $template = if ($Config.reply_templates -and $Config.reply_templates.de) { $Config.reply_templates.de } else {
                    "Sehr geehrte Damen und Herren,`n`nvielen Dank für Ihre Anfrage.`n`nKönnten Sie mir bitte mitteilen, in welchem Hotel Sie gerne eine Reservierung vornehmen möchten?`n`nIch freue mich auf Ihre Rückmeldung.`n`nMit freundlichen Grüßen"
                }
                $corr.target_folder   = "EmailProcessor/Buchung-Unbekannt"
                $corr.target_email    = ""
                $corr.action_type     = "reply"
                $corr.reply_draft     = $template
                $corr.reply_to        = $Entry.sender
                $corr.proposed_action = "Rückfrage an Gast senden (Hotel nicht erkennbar)"
            }
        }
        'application' {
            $corr.target_folder   = "EmailProcessor/Bewerbung"
            $corr.target_email    = $Config.applications_email
            $corr.proposed_action = "Weitergeleitet an $($Config.applications_email)"
        }
        'advertisement' {
            $corr.target_folder   = "EmailProcessor/Werbung"
            $corr.target_email    = ""
            $corr.proposed_action = "Archiviert als Werbung"
        }
        'newsletter' {
            $corr.target_folder   = "EmailProcessor/Newsletter-Manuell"
            $corr.target_email    = ""
            $corr.proposed_action = "Archiviert als Newsletter"
        }
        'sonstiges' {
            $corr.target_folder   = "EmailProcessor/Sonstiges"
            $corr.target_email    = ""
            $corr.proposed_action = "Wichtige Geschäftskorrespondenz - manuell bearbeiten"
        }
        'phishing' {
            $corr.target_folder   = "EmailProcessor/Sicherheitswarnung"
            $corr.target_email    = ""
            $corr.proposed_action = "ACHTUNG: Verdacht auf Phishing/Schadsoftware - Links/Anhänge NICHT öffnen. Bitte prüfen."
        }
        default {
            $corr.target_folder   = "EmailProcessor/Unbekannt"
            $corr.target_email    = ""
            $corr.proposed_action = "Archiviert (Kategorie unbekannt)"
        }
    }
    return [PSCustomObject]$corr
}

# ── Queue-Verarbeitung (alle 30 Sek.) ────────────────────────────────
function Process-Queue {
    param($Namespace, $Inbox, $Outlook, $Config)
    $q = Load-Queue
    $changed = $false

    foreach ($entry in $q) {
        if ($entry.status -eq "approved") {
            Write-Log "Genehmigt - fuehre aus: `"$($entry.subject)`""
            $result = Execute-Entry $entry $Namespace $Inbox $Outlook
            $entry.status = $result
            $changed = $true
            Record-Email $entry $result
        }
        elseif ($entry.status -eq "ignored") {
            Write-Log "Ignoriert: `"$($entry.subject)`""
            try {
                $item = $Namespace.GetItemFromID($entry.entry_id)
                [void]$item.Move((Resolve-Folder $Inbox "EmailProcessor/Ignoriert"))
                Write-Log "  -> Verschoben nach EmailProcessor/Ignoriert"
            } catch { Write-Log "  -> E-Mail nicht gefunden oder bereits verschoben" "WARN" }
            $entry.status = "done"
            $changed = $true
        }
        elseif ($entry.status -eq "rejected") {
            $reason = if ($entry.rejection_reason) { " | $($entry.rejection_reason)" } else { "" }
            Write-Log "Abgelehnt$reason`: `"$($entry.subject)`""

            if ($entry.correct_category -and $entry.correct_category -notin @('unknown','')) {
                # Korrektur vorhanden: Eintrag aktualisieren und erneut zur Pruefung stellen
                Write-Log "  -> Korrektur: $($entry.correct_category)$(if($entry.correct_hotel){" | $($entry.correct_hotel)"})"
                Save-Correction $entry
                $corrEntry = Build-CorrectedEntry $entry $Config
                $entry | Add-Member -NotePropertyName category        -NotePropertyValue $corrEntry.category        -Force
                $entry | Add-Member -NotePropertyName hotel           -NotePropertyValue $corrEntry.hotel           -Force
                $entry | Add-Member -NotePropertyName target_folder   -NotePropertyValue $corrEntry.target_folder   -Force
                $entry | Add-Member -NotePropertyName target_email    -NotePropertyValue $corrEntry.target_email    -Force
                $entry | Add-Member -NotePropertyName proposed_action -NotePropertyValue $corrEntry.proposed_action -Force
                $entry | Add-Member -NotePropertyName action_type     -NotePropertyValue $corrEntry.action_type     -Force
                $entry | Add-Member -NotePropertyName reply_to        -NotePropertyValue $corrEntry.reply_to        -Force
                $entry | Add-Member -NotePropertyName reply_draft     -NotePropertyValue $corrEntry.reply_draft     -Force
                $entry | Add-Member -NotePropertyName is_corrected    -NotePropertyValue $true                      -Force
                $entry | Add-Member -NotePropertyName correct_category -NotePropertyValue $null -Force
                $entry | Add-Member -NotePropertyName correct_hotel    -NotePropertyValue $null -Force
                $entry | Add-Member -NotePropertyName rejection_reason -NotePropertyValue $null -Force
                $entry.status = "pending"
                Write-Log "  -> Zurueck in Warteschlange fuer erneute Genehmigung"
            } else {
                # Keine Korrektur - einfach archivieren
                try {
                    $item = $Namespace.GetItemFromID($entry.entry_id)
                    [void]$item.Move((Resolve-Folder $Inbox "EmailProcessor/Abgelehnt"))
                    Write-Log "  -> Archiviert in EmailProcessor/Abgelehnt"
                } catch { Write-Log "  -> E-Mail nicht gefunden" "WARN" }
                $entry.status = "done"
            }
            $changed = $true
        }
    }

    if ($changed) { Save-Queue $q }
}

# ── Stats ─────────────────────────────────────────────────────────────
$statsFile = Join-Path $scriptDir "stats.json"

function Load-Stats {
    $def = @{ total_processed=0; today_processed=0; last_date=""; last_checked="-";
              categories=@{newsletter=0;application=0;advertisement=0;booking=0;phishing=0;unknown=0;error=0};
              hotels=@{}; recent=@() }
    if (-not (Test-Path $statsFile)) { return $def }
    try {
        $j = Get-Content $statsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $s = @{ total_processed=[int]$j.total_processed; today_processed=[int]$j.today_processed;
                last_date=[string]$j.last_date; last_checked=[string]$j.last_checked;
                categories=@{}; hotels=@{}; recent=@() }
        foreach ($k in @('newsletter','application','advertisement','booking','phishing','sonstiges','unknown','error')) {
            $s.categories[$k] = if ($j.categories.$k) { [int]$j.categories.$k } else { 0 }
        }
        if ($j.hotels) { $j.hotels.PSObject.Properties | ForEach-Object { $s.hotels[$_.Name] = [int]$_.Value } }
        if ($j.recent) { $s.recent = @($j.recent) }
        return $s
    } catch { return $def }
}

function Save-Stats { param($s); $s | ConvertTo-Json -Depth 5 | Set-Content $statsFile -Encoding UTF8 }

function Record-Email {
    param($Entry, [string]$Result = "done")
    if ($Result -ne "done") { return }
    $s = Load-Stats
    $today = Get-Date -Format "yyyy-MM-dd"
    if ($s.last_date -ne $today) { $s.today_processed = 0; $s.last_date = $today }
    $s.total_processed++; $s.today_processed++
    $s.categories[$Entry.category] = $s.categories[$Entry.category] + 1
    if ($Entry.category -eq 'booking' -and $Entry.hotel) {
        $s.hotels[$Entry.hotel] = ([int]$s.hotels[$Entry.hotel]) + 1
    }
    $rec = @{ time=(Get-Date -Format "dd.MM. HH:mm"); subject=$Entry.subject;
              sender=$Entry.sender_name; category=$Entry.category;
              hotel=$Entry.hotel; folder=$Entry.target_folder;
              corrected=([bool]$Entry.is_corrected) }
    $s.recent = @($rec) + $s.recent | Select-Object -First 100
    Save-Stats $s
}

# ── Verdächtige Spam-Emails erkennen ─────────────────────────────────
function Is-SuspiciousSpam {
    param($Item)
    $email = $Item.SenderEmailAddress.ToLower()
    $name  = $Item.SenderName

    # Nur bei öffentlichen Webmail-Providern prüfen
    if ($email -notmatch '@(gmail|hotmail|yahoo|outlook|web|gmx|aol|icloud)\.(com|de|at|ch)$') { return $false }

    $local = if ($email -match '^([^@]+)@') { $Matches[1] } else { '' }

    # Stärkstes Indiz: Absendername = E-Mail-Adresse oder lokaler Teil (kein echter Name gesetzt)
    if ($name -eq $Item.SenderEmailAddress -or $name.ToLower() -eq $local) { return $true }

    # Lokalteil > 12 Zeichen, keine Punkte/Unterstriche, enthält Zahlen → zufällig generiert
    if ($local.Length -gt 12 -and $local -notmatch '[._]' -and $local -match '\d' -and $local -notmatch '^\w+[12][0-9]{3}$') { return $true }

    return $false
}

# ── Phishing/Schadsoftware-Erkennung ─────────────────────────────────
$DangerousAttachmentExtensions = @('.exe','.scr','.bat','.cmd','.com','.pif','.vbs','.vbe','.js','.jse','.wsf','.wsh','.ps1','.psm1','.jar','.msi','.hta','.chm','.reg','.lnk','.iso','.img')
$MacroAttachmentExtensions     = @('.docm','.xlsm','.pptm','.dotm','.xltm','.xlsb')
$PhishingKeywordPattern = '(?i)(konto.{0,20}(gesperrt|deaktiviert|eingeschränkt)|passwort.{0,20}(ändern|bestätigen|aktualisieren)|verify your (account|identity|password)|dringend(er)?.{0,15}handlungsbedarf|klicken sie (hier|sofort|jetzt)|unusual (login|activity)|verdächtige (anmeldung|aktivität)|zahlung.{0,20}überfällig|rechnung.{0,15}(überfällig|sofort)|IBAN.{0,20}(geändert|aktualisiert|angepasst)|bankdaten.{0,20}(bestätigen|aktualisieren)|suspend(ed)? your account|account.{0,15}(locked|suspended)|bestätigen sie (ihre|jetzt) (daten|identität|zugangsdaten)|innerhalb von 24 stunden|within 24 hours|letzte mahnung|paket.{0,15}(zoll|gebühr).{0,15}(bezahlen|zahlen)|beschwerde.{0,20}(gast|unterkunft|buchung|über ihr)|complaint.{0,15}(against|about).{0,20}(your|the).{0,10}(property|listing|account)|guest.{0,10}complaint|ihr (angebot|inserat|profil).{0,15}(gemeldet|gesperrt)|neue bewertung.{0,15}erhalten|verifizierung.{0,15}fehlgeschlagen|ich bin kein roboter|verify you are human|nicht.?automatisiert.{0,10}best[äa]tigen)'

function Get-AttachmentRisk {
    param($Item)
    $hits = @()
    try {
        foreach ($att in $Item.Attachments) {
            $fn  = $att.FileName
            $ext = [IO.Path]::GetExtension($fn).ToLower()
            $hasDoubleExt = [regex]::IsMatch($fn, '\.[A-Za-z0-9]{2,5}\.[A-Za-z0-9]{2,5}$')
            if ($ext -in $DangerousAttachmentExtensions) {
                $hits += "Gefaehrlicher Anhangstyp ($ext): $fn"
            } elseif ($ext -in $MacroAttachmentExtensions) {
                $hits += "Makro-faehiges Office-Dokument: $fn"
            } elseif ($hasDoubleExt -and $ext -in @('.exe','.scr','.bat','.cmd','.js','.vbs','.pif','.com')) {
                $hits += "Verschleierte doppelte Dateiendung: $fn"
            }
        }
    } catch {}
    return $hits
}

function Get-LinkRisk {
    param($Item)
    $hits = @()
    try {
        $html = $Item.HTMLBody
        if (-not $html) { return $hits }
        $linkMatches = [regex]::Matches($html, '(?is)<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>')
        $seen = @{}
        foreach ($m in $linkMatches) {
            $href = $m.Groups[1].Value
            $text = ($m.Groups[2].Value -replace '<[^>]+>','').Trim()
            if ($href -notmatch '^https?://') { continue }
            try { $hrefDomain = ([Uri]$href).Host.ToLower() } catch { continue }

            # Link zeigt auf eine IP-Adresse statt auf eine Domain
            if ($hrefDomain -match '^\d{1,3}(\.\d{1,3}){3}$' -and -not $seen.ContainsKey('ip')) {
                $hits += "Link verweist auf IP-Adresse statt Domain: $hrefDomain"; $seen['ip'] = $true
            }

            # Sichtbarer Linktext nennt eine Domain, die vom tatsaechlichen Link abweicht
            if ($text -match '([a-z0-9\-]+\.)+[a-z]{2,}' -and -not $seen.ContainsKey('mismatch')) {
                $textDomain = $Matches[0].ToLower()
                if ($hrefDomain -notlike "*$textDomain*" -and $textDomain -notlike "*$hrefDomain*") {
                    $hits += "Linktext '$text' verweist tatsaechlich auf: $hrefDomain"; $seen['mismatch'] = $true
                }
            }
        }
    } catch {}
    return $hits
}

function Get-KeywordRisk {
    param($Item)
    $text = "$($Item.Subject) $($Item.Body)"
    if ($text -match $PhishingKeywordPattern) { return @("Verdaechtige Formulierung erkannt: '$($Matches[0].Trim())'") }
    return @()
}

# Bekannte Marken: falls im Absendernamen genannt, aber die Absenderdomain nicht dazu passt
$KnownBrandDomains = @{
    'paypal'=@('paypal.com'); 'microsoft'=@('microsoft.com'); 'amazon'=@('amazon.de','amazon.com');
    'dhl'=@('dhl.de','dhl.com'); 'apple'=@('apple.com'); 'google'=@('google.com'); 'ebay'=@('ebay.de','ebay.com');
    'sparkasse'=@('sparkasse.de'); 'volksbank'=@('volksbank.de'); 'postbank'=@('postbank.de'); 'ionos'=@('ionos.de');
    'booking'=@('booking.com'); 'airbnb'=@('airbnb.com','airbnb.de'); 'expedia'=@('expedia.com','expedia.de');
    'hrs'=@('hrs.com','hrs.de'); 'tripadvisor'=@('tripadvisor.com','tripadvisor.de')
}

function Get-SpoofRisk {
    param($Item)
    $hits = @()
    try {
        $displayName = "$($Item.SenderName)"
        $email  = $Item.SenderEmailAddress.ToLower()
        $domain = if ($email -match '@(.+)$') { $Matches[1] } else { '' }
        foreach ($brand in $KnownBrandDomains.Keys) {
            if ($displayName -match "(?i)\b$brand\b") {
                $matchesDomain = $false
                foreach ($d in $KnownBrandDomains[$brand]) { if ($domain -like "*$d") { $matchesDomain = $true; break } }
                if (-not $matchesDomain) { $hits += "Absendername nennt '$brand', Absenderdomain passt nicht: $domain" }
            }
        }
    } catch {}
    return $hits
}

function Get-UniqueLinks {
    param($Item)
    $links = @()
    try {
        $html = $Item.HTMLBody
        if (-not $html) { return $links }
        $linkMatches = [regex]::Matches($html, '(?is)href="(https?://[^"]+)"')
        $seen = @{}
        foreach ($m in $linkMatches) {
            $href = $m.Groups[1].Value
            if (-not $seen.ContainsKey($href)) { $seen[$href] = $true; $links += $href }
        }
    } catch {}
    return $links
}

# ── VirusTotal (Datei-Hash + Link-Reputation, benoetigt kostenlosen API-Key) ──
function Get-VTFileVerdict {
    param([string]$Sha256, $Config)
    if (-not $Config.virustotal_api_key) { return $null }
    try {
        $resp = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/files/$Sha256" `
            -Headers @{ 'x-apikey' = $Config.virustotal_api_key } -TimeoutSec 15 -ErrorAction Stop
        $stats = $resp.data.attributes.last_analysis_stats
        if ($stats -and ($stats.malicious -gt 0 -or $stats.suspicious -gt 0)) {
            return "VirusTotal: Anhang als Schadsoftware eingestuft ($($stats.malicious) Engines melden Malware, $($stats.suspicious) verdaechtig)"
        }
    } catch {}
    return $null
}

function Get-VTUrlVerdict {
    param([string]$Url, $Config)
    if (-not $Config.virustotal_api_key) { return $null }
    try {
        $urlId = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Url)).TrimEnd('=').Replace('+','-').Replace('/','_')
        $resp = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/urls/$urlId" `
            -Headers @{ 'x-apikey' = $Config.virustotal_api_key } -TimeoutSec 15 -ErrorAction Stop
        $stats = $resp.data.attributes.last_analysis_stats
        if ($stats -and ($stats.malicious -gt 0 -or $stats.suspicious -gt 0)) {
            return "VirusTotal: Link als gefaehrlich eingestuft ($($stats.malicious) Engines melden Malware/Phishing): $Url"
        }
    } catch {}
    return $null
}

# ── URLhaus (abuse.ch) — kostenlose Link-Reputation ohne API-Key ──
function Get-URLhausVerdict {
    param([string]$Url)
    try {
        $resp = Invoke-RestMethod -Uri "https://urlhaus-api.abuse.ch/v1/url/" -Method POST `
            -Body @{ url = $Url } -TimeoutSec 15 -ErrorAction Stop
        if ($resp.query_status -eq 'ok') {
            return "URLhaus: Link ist als bekannte Schadsoftware-/Phishing-URL gemeldet ($($resp.threat)): $Url"
        }
    } catch {}
    return $null
}

# ── Windows Defender — lokaler Anhang-Scan, kein Internet-Call ──
function Get-DefenderVerdict {
    param([string]$Path)
    try {
        Import-Module Defender -ErrorAction Stop
        $before = Get-Date
        Start-MpScan -ScanPath $Path -ScanType CustomScan -ErrorAction Stop
        Start-Sleep -Milliseconds 800
        $hit = Get-MpThreatDetection -ErrorAction Stop |
            Where-Object { $_.InitialDetectionTime -ge $before.AddSeconds(-3) } | Select-Object -First 1
        if ($hit) { return "Windows Defender: Bedrohung im Anhang erkannt ($($hit.ThreatName))" }
    } catch {}
    return $null
}

function Test-PhishingRisk {
    param($Item, $Config)
    $reasons = @()
    $reasons += Get-AttachmentRisk $Item
    $reasons += Get-LinkRisk $Item
    $reasons += Get-KeywordRisk $Item
    $reasons += Get-SpoofRisk $Item

    $isInternal = ($Item.SenderEmailType -eq "EX")

    # Anhaenge: lokaler Defender-Scan fuer alle, VirusTotal-Hash-Check nur fuer externe Absender
    try {
        foreach ($att in $Item.Attachments) {
            $tmp = Join-Path $env:TEMP ("epscan_" + [Guid]::NewGuid().ToString("N") + "_" + $att.FileName)
            try {
                $att.SaveAsFile($tmp)
                $defenderHit = Get-DefenderVerdict $tmp
                if ($defenderHit) { $reasons += $defenderHit }
                if (-not $isInternal -and $Config.virustotal_api_key) {
                    $sha = (Get-FileHash -Path $tmp -Algorithm SHA256 -ErrorAction Stop).Hash
                    $vtHit = Get-VTFileVerdict $sha $Config
                    if ($vtHit) { $reasons += $vtHit }
                }
            } catch {} finally {
                try { Remove-Item $tmp -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    } catch {}

    # Links: URLhaus + VirusTotal nur fuer externe Absender pruefen (Datenschutz + API-Kontingent),
    # maximal 3 pro Email
    if (-not $isInternal) {
        try {
            $links = @(Get-UniqueLinks $Item) | Select-Object -First 3
            foreach ($url in $links) {
                $uhHit = Get-URLhausVerdict $url
                if ($uhHit) { $reasons += $uhHit }
                if ($Config.virustotal_api_key) {
                    Start-Sleep -Milliseconds 300
                    $vtHit = Get-VTUrlVerdict $url $Config
                    if ($vtHit) { $reasons += $vtHit }
                }
            }
        } catch {}
    }

    return $reasons
}

# ── Geschäfts-/Kollegenemails erkennen und überspringen ──────────────
function Should-Skip {
    param($Item, $Config)

    # Interne Exchange-Emails (Arbeitskollegen auf demselben Server) immer überspringen
    if ($Item.SenderEmailType -eq "EX") { return $true }

    # SMTP-Adresse des Absenders ermitteln
    $senderEmail = $Item.SenderEmailAddress.ToLower()

    # Domains aus skip_domains-Liste prüfen
    if ($Config.skip_domains) {
        foreach ($domain in $Config.skip_domains) {
            if ($senderEmail -like "*@$($domain.ToLower())") { return $true }
        }
    }

    # Einzelne Adressen aus skip_senders-Liste prüfen
    if ($Config.skip_senders) {
        foreach ($addr in $Config.skip_senders) {
            if ($senderEmail -eq $addr.ToLower()) { return $true }
        }
    }

    return $false
}

# ── Posteingang scannen ───────────────────────────────────────────────
function Scan-Inbox {
    param($Config, $Inbox, $Namespace, $Outlook)
    $s = Load-Stats; $s.last_checked = Get-Date -Format "dd.MM.yyyy HH:mm"; Save-Stats $s

    # Bereits in der Queue vorhandene EntryIDs sammeln (nicht doppelt einreihen)
    $knownIds = @(Load-Queue | Where-Object { $_.status -in @('pending','approved','rejected','done','open_task','task_done','ignored','dismissed') } | ForEach-Object { $_.entry_id })

    $maxAgeDays = if ($Config.scan_max_age_days) { $Config.scan_max_age_days } else { 1 }
    $maxPerScan = if ($Config.scan_max_per_run)  { $Config.scan_max_per_run  } else { 50 }
    $cutoff     = (Get-Date).Date.AddDays(-($maxAgeDays - 1))

    $mails = @()
    $items = $Inbox.Items; $items.Sort("[ReceivedTime]", $true)
    foreach ($item in $items) {
        try {
            if ($item.Class -eq 43 -and $item.ReceivedTime -ge $cutoff -and $item.EntryID -notin $knownIds) {
                $mails += $item
                if ($mails.Count -ge $maxPerScan) { break }
            }
        } catch {}
    }

    if ($mails.Count -eq 0) { Write-Log "Keine neuen E-Mails (letzte $maxAgeDays Tage)."; return }

    Write-Log ("=" * 55)
    Write-Log "$($mails.Count) neue E-Mail(s) - werden zur Pruefung eingereiht"
    Write-Log ("=" * 55)

    $i = 0
    $processedKeys = @{}  # Duplikat-Schutz: Absender+Betreff innerhalb dieses Scans

    foreach ($mail in $mails) {
        $i++

        # Duplikat-Schutz: gleiche Email im selben Scan nicht doppelt verarbeiten
        $dedupKey = "$($mail.SenderEmailAddress.ToLower())|$($mail.Subject)"
        if ($processedKeys.ContainsKey($dedupKey)) {
            Write-Log "Duplikat uebersprungen: `"$($mail.Subject)`" | $($mail.SenderName)"
            continue
        }

        # Phishing/Schadsoftware-Pruefung: hoechste Prioritaet, vor allen anderen Checks
        $riskReasons = @(Test-PhishingRisk $mail $Config)
        if ($riskReasons.Count -gt 0) {
            Write-Log "[$i/$($mails.Count)] SICHERHEITSWARNUNG: `"$($mail.Subject)`" | $($mail.SenderName)" "WARN"
            foreach ($r in $riskReasons) { Write-Log "  -> $r" "WARN" }
            try {
                Quarantine-Suspicious-Email $mail $Inbox $riskReasons $Config | Out-Null
                $processedKeys[$dedupKey] = $true
                Write-Log "  -> In EmailProcessor/Sicherheitswarnung verschoben, wartet auf Pruefung"
            } catch { Write-Log "  -> Fehler bei Quarantaene: $_" "ERROR" }
            continue
        }

        # Interne/Kollegen-Emails: rot markieren und als offene Aufgabe speichern
        if (-not $Config.test_mode -and (Should-Skip $mail $Config)) {
            Write-Log "Kollegen-Email: `"$($mail.Subject)`" | $($mail.SenderName)"
            try { $mail.FlagStatus = 2; $mail.Save() } catch { Write-Log "  -> Markierung fehlgeschlagen: $_" "WARN" }
            $taskEntry = [ordered]@{
                id          = [System.Guid]::NewGuid().ToString()
                entry_id    = $mail.EntryID
                subject     = if ($mail.Subject.Length -gt 80) { $mail.Subject.Substring(0,80) } else { $mail.Subject }
                sender      = $mail.SenderEmailAddress
                sender_name = $mail.SenderName
                received    = $mail.ReceivedTime.ToString("dd.MM.yyyy HH:mm")
                created     = (Get-Date -Format "dd.MM. HH:mm")
                category    = "task"
                status      = "open_task"
                proposed_action = "Manuelle Bearbeitung"
            }
            Add-To-Queue $taskEntry
            $processedKeys[$dedupKey] = $true
            Write-Log "  -> Rot markiert + als Aufgabe gespeichert"
            continue
        }

        # Verdächtige Spam-Emails: Ollama überspringen, direkt als Werbung einreihen
        if (Is-SuspiciousSpam $mail) {
            Write-Log "[$i/$($mails.Count)] Spam-Absender erkannt: `"$($mail.Subject)`" | $($mail.SenderName)"
            try {
                $entry = Build-QueueEntry $mail 'advertisement' '' $Config
                $entry.proposed_action = "Archivieren (verdaechtiger Spam-Absender — kein echter Name gesetzt)"
                $moved = $mail.Move((Resolve-Folder $Inbox "EmailProcessor/Zur Pruefung"))
                $entry.entry_id = $moved.EntryID
                Add-To-Queue $entry
                $processedKeys[$dedupKey] = $true
                Write-Log "  -> Als Werbung/Spam eingereiht (Ollama uebersprungen)"
            } catch { Write-Log "  -> Fehler beim Spam-Handling: $_" "ERROR" }
            continue
        }

        # Ausstehende Genehmigungen VOR dem Ollama-Aufruf verarbeiten
        # (damit genehmigte Emails nicht durch laufende Klassifizierung blockiert werden)
        Process-Queue $Namespace $Inbox $Outlook $Config

        Write-Log "[$i/$($mails.Count)] Klassifiziere: `"$($mail.Subject)`" | $($mail.SenderName)"
        try {
            $cls = Classify-Email $mail $Config

            # KI erkennt Phishing/Betrug ohne technische Indizien (z.B. reine Text-Masche) -> ebenfalls isolieren
            if ($cls.category -eq 'phishing') {
                Write-Log "  -> KI-Verdacht auf Phishing/Betrug: $($cls.reason)" "WARN"
                Quarantine-Suspicious-Email $mail $Inbox @("KI-Einschaetzung: $($cls.reason)") $Config | Out-Null
                $processedKeys[$dedupKey] = $true
                Write-Log "  -> In EmailProcessor/Sicherheitswarnung verschoben, wartet auf Pruefung"
                continue
            }

            $entry = Build-QueueEntry $mail $cls.category $cls.hotel $Config

            # Erst verschieben, dann neue EntryID speichern (EntryID aendert sich beim Verschieben)
            $moved = $mail.Move((Resolve-Folder $Inbox "EmailProcessor/Zur Pruefung"))
            $entry.entry_id = $moved.EntryID
            Add-To-Queue $entry
            $processedKeys[$dedupKey] = $true

            Write-Log "  -> Wartet auf Genehmigung: $($entry.proposed_action)"
        } catch {
            Write-Log "  -> Fehler: $_" "ERROR"
        }
    }
}

# ── Hauptschleife ─────────────────────────────────────────────────────
# Sicherstellen dass keine zweite Instanz läuft
$myPid    = $PID
$otherBot = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.ProcessId -ne $myPid -and $_.CommandLine -like '*bot.ps1*' }
if ($otherBot) {
    Write-Log "FEHLER: Bot läuft bereits (PID $($otherBot.ProcessId)). Diese Instanz wird beendet." "ERROR"
    exit 1
}

Write-Log "Bot gestartet - Sicherheitsmodus aktiv (manuelle Genehmigung erforderlich)"
$app   = Get-OutlookApp
$ns    = $app.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)

$config          = Load-Config
$lastScan        = [DateTime]::MinValue
$scanIntervalSec = $config.check_interval_minutes * 60

Write-Log "Modell: $($config.ollama_model) | Scanintervall: $($config.check_interval_minutes) Min. | Queue-Check: alle 10 Sek."

while ($true) {
    try {
        $config = Load-Config

        # Posteingang neu laden (Outlook-Verbindung frisch halten)
        try { $inbox = $ns.GetDefaultFolder(6) } catch {
            Write-Log "Outlook-Verbindung verloren, versuche Neuverbindung..." "WARN"
            try { $app = Get-OutlookApp; $ns = $app.GetNamespace("MAPI"); $inbox = $ns.GetDefaultFolder(6) }
            catch { Write-Log "Neuverbindung fehlgeschlagen: $_" "ERROR"; Start-Sleep -Seconds 10; continue }
        }

        # Queue alle 10 Sekunden prüfen (schnellere Reaktion auf Genehmigungen)
        Process-Queue $ns $inbox $app $config

        # Posteingang nur alle N Minuten scannen
        $now = [DateTime]::Now
        if (($now - $lastScan).TotalSeconds -ge $scanIntervalSec) {
            Scan-Inbox $config $inbox $ns $app
            $lastScan        = $now
            $scanIntervalSec = $config.check_interval_minutes * 60
        }
    } catch {
        Write-Log "Unerwarteter Fehler im Hauptloop: $_" "ERROR"
    }

    Start-Sleep -Seconds 10
}
