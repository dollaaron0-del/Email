# Email-Processor Bot - PowerShell + Outlook + Ollama
# Emails werden NICHT automatisch versendet - zuerst im Dashboard genehmigen!
$ErrorActionPreference = "Continue"
$scriptDir       = $PSScriptRoot
$correctionsFile = Join-Path $scriptDir "corrections.json"

# Bot-Prozess auf hohe Priorität setzen
try { [System.Diagnostics.Process]::GetCurrentProcess().PriorityClass = [System.Diagnostics.ProcessPriorityClass]::AboveNormal } catch {}

# Langdock-Anbindung (API-Aufruf, Prompt, Antwort-Pruefung)
. (Join-Path $scriptDir "langdock.ps1")

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

function Get-UnsubscribeInfo {
    param($Item)
    $info = @{ raw = ""; mailto = ""; http = ""; one_click = $false }
    try {
        $h = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
        if ($h -match "(?i)List-Unsubscribe:\s*(.+?)(\r?\n\S|\z)") {
            $raw = $Matches[1].Trim()
            $info.raw = $raw
            if ($raw -match '<mailto:([^>]+)>') { $info.mailto = $Matches[1] }
            if ($raw -match '<(https?://[^>]+)>') { $info.http = $Matches[1] }
        }
        if ($h -match '(?i)List-Unsubscribe-Post:\s*List-Unsubscribe=One-Click') {
            $info.one_click = $true
        }
    } catch {}
    return $info
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

function Get-HotelsBlock {
    param($Config)
    return ($Config.hotels | ForEach-Object { "  - $($_.name) (Keywords: $($_.keywords -join ', '))" }) -join "`n"
}

function Get-ExamplesBlock {
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
    return $examplesBlock
}

function Get-AttachmentNames {
    param($Item)
    $names = @()
    try { foreach ($a in $Item.Attachments) { $names += "$($a.FileName)" } } catch {}
    return $names   # Aufrufer immer mit @(...) umschliessen
}

function Get-InternetMessageId {
    # Stabile Kennung der Mail (bleibt beim Verschieben gleich, gibt es auch in Microsoft Graph)
    param($Item)
    try { return "$($Item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F'))" } catch { return "" }
}

# ── Langdock: eine Abfrage pro Mail ──────────────────────────────────
function Get-LlmProvider {
    param($Config)
    if ("$($Config.llm_provider)".ToLower() -eq 'langdock') { return 'langdock' }
    return 'ollama'
}

function Invoke-LangdockAnalysis {
    # Liefert Kategorie, Hotel, Sicherheit, Sprache und Phishing-Einschaetzung - wirft bei Fehlern
    param($Item, $Config)
    $maxChars = if ($Config.langdock.body_chars) { [int]$Config.langdock.body_chars } else { 3000 }
    $bodyText = "$($Item.Body)".Trim()
    if ($bodyText.Length -gt $maxChars) { $bodyText = $bodyText.Substring(0, $maxChars) }
    $att = @(Get-AttachmentNames $Item)

    $prompt = Get-AnalysisPrompt -HotelsBlock (Get-HotelsBlock $Config) -ExamplesBlock (Get-ExamplesBlock) `
        -SenderName "$($Item.SenderName)" -SenderEmail "$($Item.SenderEmailAddress)" -Subject "$($Item.Subject)" `
        -Attachments $(if ($att.Count) { $att -join ', ' } else { 'keine' }) -BodyText $bodyText

    $raw    = Invoke-LangdockJson -Prompt $prompt -Schema (Get-AnalysisSchema) -Config $Config
    $result = ConvertTo-AnalysisResult $raw $Config
    Write-Log "  -> Langdock: $($result.category) | Hotel: $($result.hotel) | $($result.confidence) | $($result.reason)"
    if ($result.is_phishing) { Write-Log "  -> Langdock Phishing-Verdacht ($($result.phishing_confidence)): $($result.phishing_reason)" "WARN" }
    return $result
}

function Merge-LlmPhishing {
    # Uebernimmt die Phishing-Einschaetzung aus der Langdock-Analyse in das Ergebnis der Sicherheitspruefung
    param($SecCheck, $Analysis)
    if ($Analysis -and $Analysis.is_phishing) {
        $SecCheck.reasons += "LLM-Phishing-Erkennung ($($Analysis.phishing_confidence)): $($Analysis.phishing_reason)"
        if ($Analysis.phishing_confidence -in @('high','medium')) { $SecCheck.threat_level = 'danger' }
        elseif ($SecCheck.threat_level -eq 'none') { $SecCheck.threat_level = 'warn' }
        $SecCheck.is_threat = ($SecCheck.threat_level -ne 'none')
    }
    return $SecCheck
}

# ── Automatische Freigabe ────────────────────────────────────────────
$CONF_RANK = @{ low = 1; medium = 2; high = 3 }

function Get-AutoApproveMode {
    param($Config)
    $m = "$($Config.auto_approve.mode)".ToLower()
    if ($m -in @('shadow','on')) { return $m }
    return 'off'
}

function Get-AutoDecision {
    # Darf dieser Eintrag ohne manuelle Pruefung ausgefuehrt werden? Liefert eligible + Begruendung.
    param($Entry, $Prediction, $Config)
    if ((Get-AutoApproveMode $Config) -eq 'off') { return @{ eligible = $false; note = "" } }

    $allowed = if ($Config.auto_approve.categories) { @($Config.auto_approve.categories) } else { @('advertisement','newsletter','application','booking') }
    $minConf = "$($Config.auto_approve.min_confidence)".ToLower()
    if ($minConf -notin @('high','medium')) { $minConf = 'high' }
    $cat  = "$($Entry.category)"
    $conf = if ($Prediction) { "$($Prediction.confidence)".ToLower() } else { "" }

    if ($Config.test_mode)        { return @{ eligible = $false; note = "Test-Modus aktiv" } }
    if (-not $Prediction)         { return @{ eligible = $false; note = "keine KI-Einschaetzung" } }
    if ($Prediction.provider -ne 'langdock' -and -not $Config.auto_approve.allow_ollama) {
                                    return @{ eligible = $false; note = "nur Langdock-Ergebnisse werden automatisch freigegeben" } }
    if ($Entry.is_corrected)      { return @{ eligible = $false; note = "korrigierter Eintrag" } }
    # Diese Faelle bleiben immer manuell, egal was in der Config steht
    if ($cat -in @('sonstiges','unknown','security_threat','task')) { return @{ eligible = $false; note = "Kategorie wird immer manuell geprueft" } }
    if ($cat -notin $allowed)     { return @{ eligible = $false; note = "Kategorie nicht fuer Auto-Freigabe freigeschaltet" } }
    if ($Entry.security_status -ne 'clean') { return @{ eligible = $false; note = "Sicherheitspruefung nicht bestanden" } }
    if (-not ($CONF_RANK[$conf] -ge $CONF_RANK[$minConf])) { return @{ eligible = $false; note = "Sicherheit nur '$conf'" } }

    switch ($cat) {
        'booking' {
            if ($Entry.action_type -ne 'forward' -or -not $Entry.target_email) {
                return @{ eligible = $false; note = "Rueckfrage an Gast wird immer manuell geprueft" }
            }
        }
        'newsletter' {
            if (-not $Entry.unsubscribe_one_click) {
                return @{ eligible = $false; note = "Abmeldung ohne One-Click wird manuell geprueft" }
            }
        }
        'application' {
            if (-not $Entry.target_email) { return @{ eligible = $false; note = "keine Bewerbungs-Adresse konfiguriert" } }
        }
    }
    return @{ eligible = $true; note = "Sicherheit '$conf'" }
}

function Set-AutoDecision {
    # Schreibt die Entscheidung in den Queue-Eintrag. Modus "on": sofort genehmigen, "shadow": nur markieren.
    param($Entry, $Prediction, $Config)
    $mode = Get-AutoApproveMode $Config
    $d = Get-AutoDecision $Entry $Prediction $Config
    $Entry.auto_mode     = $mode
    $Entry.auto_eligible = [bool]$d.eligible
    $Entry.auto_note     = $d.note
    $Entry.auto_approved = $false
    if ($mode -eq 'on' -and $d.eligible) {
        $Entry.status        = 'approved'
        $Entry.auto_approved = $true
    }
}

# ── Trainingsdaten ────────────────────────────────────────────────────
# Eine Zeile pro Ereignis in training_data/emails_JJJJ-MM.jsonl:
#   "prediction" = was die KI vorgeschlagen hat (inkl. Mailtext)
#   "decision"   = was danach passiert ist (genehmigt, korrigiert, abgelehnt, ignoriert, automatisch)
# Zusammenfuehren mit export_training_data.ps1 (Verknuepfung ueber "id").
function Test-TrainingEnabled {
    param($Config)
    return ($Config.training_data -and $Config.training_data.enabled -eq $true)
}

function Write-TrainingEvent {
    param($Config, $Record)
    if (-not (Test-TrainingEnabled $Config)) { return }
    try {
        $dir = if ($Config.training_data.dir) { "$($Config.training_data.dir)" } else { "training_data" }
        if (-not [IO.Path]::IsPathRooted($dir)) { $dir = Join-Path $scriptDir $dir }
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $file = Join-Path $dir ("emails_" + (Get-Date -Format "yyyy-MM") + ".jsonl")
        $line = ConvertTo-Json -InputObject $Record -Depth 8 -Compress
        [IO.File]::AppendAllText($file, $line + "`n", (New-Object Text.UTF8Encoding($false)))
    } catch { Write-Log "  -> Trainingsdaten konnten nicht geschrieben werden: $_" "WARN" }
}

function Get-TrainingInput {
    # Alles, was ein Modell spaeter als Eingabe sehen soll - VOR dem Verschieben erfassen
    param($Item, $Config, [string]$MessageId)
    $maxChars = if ($Config.training_data.body_chars) { [int]$Config.training_data.body_chars } else { 4000 }
    $body = "$($Item.Body)".Trim()
    if ($body.Length -gt $maxChars) { $body = $body.Substring(0, $maxChars) }
    $sender = "$($Item.SenderEmailAddress)".ToLower()
    $u = Get-UnsubscribeInfo $Item
    return [ordered]@{
        message_id            = $MessageId
        received              = $Item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
        sender_domain         = if ($sender -match '@(.+)$') { $Matches[1] } else { "" }
        subject               = "$($Item.Subject)"
        body                  = $body
        attachments           = @(Get-AttachmentNames $Item)
        has_list_unsubscribe  = [bool]$u.raw
        one_click_unsubscribe = [bool]$u.one_click
    }
}

function Write-TrainingPrediction {
    param($Config, $Entry, $TrainInput, $Prediction, $Extra)
    if (-not $TrainInput) { return }
    Write-TrainingEvent $Config ([ordered]@{
        event           = "prediction"
        id              = $Entry.id
        time            = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        prompt_version  = $LANGDOCK_PROMPT_VERSION
        input           = $TrainInput
        prediction      = $Prediction
        queued_category = $Entry.category
        queued_hotel    = $Entry.hotel
        proposed_action = $Entry.proposed_action
        auto            = [ordered]@{ mode = $Entry.auto_mode; eligible = [bool]$Entry.auto_eligible; approved = [bool]$Entry.auto_approved; note = $Entry.auto_note }
        extra           = $Extra
    })
}

function Write-TrainingDecision {
    param($Config, $Entry, [string]$Decision, [string]$FinalCategory, [string]$FinalHotel, [string]$Result = "")
    Write-TrainingEvent $Config ([ordered]@{
        event          = "decision"
        id             = $Entry.id
        message_id     = $Entry.message_id
        time           = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
        decision       = $Decision
        final_category = $FinalCategory
        final_hotel    = $FinalHotel
        was_corrected  = [bool]$Entry.is_corrected
        result         = $Result
    })
}

# ── Ollama ────────────────────────────────────────────────────────────
function Classify-Email {
    param($Item, $Config)
    $hotels = Get-HotelsBlock $Config
    $examplesBlock = Get-ExamplesBlock

    function Invoke-Classify($bodyText) {
        $prompt = @"
Klassifiziere diese E-Mail im Kontext Hotelmanagement.

Kategorien (waehle EXAKT eine):
- booking: Buchungsanfragen, Zimmerverfuegbarkeit, Reservierungen, Gruppenanfragen - AUCH von privaten Adressen wie gmail.com, hotmail.com etc.
- application: Bewerbungen, Lebenslauf-Einsendungen, Jobanfragen - erkennbar an Karrieredomains (stepstone, linkedin, hotelpersonal, dki-lehrlinge) ODER expliziten Bewerbungsformulierungen
- newsletter: Regelmaessige Rundschreiben, Magazine, Branchennews, oft mit Abmelde-Link
- advertisement: Werbung, Verkaufsangebote, B2B-Marketing, Dienstleistungsangebote, Produktpraesentation
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
        if ($cat -notin @('newsletter','application','advertisement','booking','sonstiges','unknown')) { $cat = 'unknown' }
        Write-Log "  -> $cat | Hotel: $($r.hotel_name) | $($r.confidence) | $($r.reason)"
        return @{ category = $cat; hotel = "$($r.hotel_name)"; confidence = "$($r.confidence)".ToLower(); reason = "$($r.reason)" }
    } catch {
        Write-Log "  -> Erster Versuch fehlgeschlagen, retry mit laengerem Text..." "WARN"
        try {
            $body = if ($Item.Body.Length -gt 2000) { $Item.Body.Substring(0,2000) } else { $Item.Body }
            $r = Invoke-Classify $body
            $cat = $r.category
            if ($cat -notin @('newsletter','application','advertisement','booking','sonstiges','unknown')) { $cat = 'unknown' }
            Write-Log "  -> $cat | Hotel: $($r.hotel_name) | $($r.confidence) | $($r.reason)"
            return @{ category = $cat; hotel = "$($r.hotel_name)"; confidence = "$($r.confidence)".ToLower(); reason = "$($r.reason)" }
        } catch {
            Write-Log "  -> Klassifizierungsfehler: $_" "ERROR"
            return @{ category = 'unknown'; hotel = $null; confidence = 'low'; reason = "Klassifizierungsfehler" }
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
    param($Item, $Config, [string]$Language = "")

    $templates = $Config.reply_templates

    # 1. Sprache: aus der Langdock-Analyse uebernehmen, sonst per Ollama erkennen
    if ($Language) {
        $lang = $Language
        if (-not ($templates -and $templates.$lang)) { $lang = if ($templates -and $templates.en) { 'en' } else { 'de' } }
    } else {
        $lang = Detect-Language $Item $Config
    }
    Write-Log "  -> Sprache erkannt: $lang"

    # 2. Passendes Template aus Config laden
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
    param($Item, [string]$Category, [string]$Hotel, $Config, [string]$Language = "")

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
        unsubscribe_header   = ""
        unsubscribe_http     = ""
        unsubscribe_mailto   = ""
        unsubscribe_one_click = $false
        body_preview = if ($Item.Body) { $b = $Item.Body.Trim(); if ($b.Length -gt 2000) { $b.Substring(0,2000) + "…" } else { $b } } else { "" }
    }

    switch ($Category) {
        'newsletter' {
            $u = Get-UnsubscribeInfo $Item
            $entry.unsubscribe_header    = $u.raw
            $entry.unsubscribe_http      = $u.http
            $entry.unsubscribe_mailto    = $u.mailto
            $entry.unsubscribe_one_click = $u.one_click
            if ($u.http) {
                $entry.unsubscribe_method = "http"
                $entry.proposed_action    = if ($u.one_click) { "RFC 8058 One-Click Abmeldung (automatisch)" } else { "Abmelde-Webseite aufrufen + Button klicken" }
                $entry.target_folder      = "EmailProcessor/Abgemeldet"
            } elseif ($u.mailto) {
                $entry.unsubscribe_method = "mailto"
                $entry.proposed_action    = "Abmelde-Email senden"
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
                $draft = Generate-Rueckfrage $Item $Config $Language
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
        'security_threat' {
            $entry.proposed_action = "SICHERHEITSWARNUNG - nicht oeffnen oder ausfuehren!"
            $entry.target_folder   = "EmailProcessor/Sicherheitswarnung"
        }
        default {
            $entry.proposed_action = "Kategorie unbekannt - manuell prüfen"
            $entry.target_folder   = "EmailProcessor/Unbekannt"
        }
    }
    return $entry
}

# ── Aktion ausführen (nach Genehmigung) ──────────────────────────────
function Invoke-WebUnsubscribe {
    param([string]$Url, [bool]$OneClick = $false)

    $ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    $reqHdr = @{
        "User-Agent"      = $ua
        "Accept"          = "text/html,application/xhtml+xml,*/*;q=0.9"
        "Accept-Language" = "de-DE,de;q=0.9,en;q=0.8"
    }

    # RFC 8058 One-Click: einfacher POST, keine Seiteninteraktion nötig
    if ($OneClick) {
        try {
            $r = Invoke-WebRequest -Uri $Url -Method POST -Headers $reqHdr `
                -Body "List-Unsubscribe=One-Click" `
                -ContentType "application/x-www-form-urlencoded" `
                -TimeoutSec 15 -UseBasicParsing
            if ($r.StatusCode -lt 400) {
                return @{ method = "rfc8058"; message = "RFC 8058 One-Click Abmeldung erfolgreich (HTTP $($r.StatusCode))" }
            }
        } catch {
            Write-Log "  -> RFC 8058 POST fehlgeschlagen, versuche normales GET..." "WARN"
        }
    }

    # Phrasen für erkannte Abmeldung (DE/EN/FR/NL/ES)
    $confirmRx = '(?i)(erfolgreich.{0,12}abgemeldet|abmeldung.erfolgreich|erfolgreich.entfernt|abonnement.gek.ndigt|unsubscribed|successfully.removed|opt.?out.success|you.ve been unsubscribed|you have been unsubscribed|you.re.unsubscribed|have been removed|no longer receive|désabonné|vous avez .t. désabonné|afgemeld|avmeldt|dado de baja)'

    try {
        $resp = Invoke-WebRequest -Uri $Url -Method GET -TimeoutSec 20 -UseBasicParsing -SessionVariable sess -Headers $reqHdr
        $html = $resp.Content

        # Direkte Bestätigung per GET?
        if ($html -match $confirmRx) {
            return @{ method = "get"; message = "Seite bestätigt Abmeldung direkt (kein Klick nötig)" }
        }

        $base = [Uri]$Url

        # Abmelde-Formular suchen und abschicken
        $formPattern = "(?is)<form[^>]*?action=""?([^\s"">]+)""?[^>]*?>(.*?)</form>"
        foreach ($fm in [regex]::Matches($html, $formPattern)) {
            $action  = $fm.Groups[1].Value
            $content = $fm.Groups[2].Value
            if ($content -notmatch '(?i)(unsubscribe|abmeld|opt.?out|austragen|confirm|remove|abbestell)') { continue }

            $actionUrl = if ($action -match '^https?://') { $action }
                         elseif ($action.StartsWith('/'))  { "$($base.Scheme)://$($base.Host)$action" }
                         else                              { "$($base.Scheme)://$($base.Host)/$action" }
            $fields = @{}
            [regex]::Matches($content, '(?i)<input([^>]+)>') | ForEach-Object {
                $inp   = $_.Groups[1].Value
                $itype = if ($inp -match '(?i)type="?(\w+)') { $Matches[1] } else { "text" }
                $name  = if ($inp -match '(?i)name="([^">\s]+)"') { $Matches[1] } else { $null }
                $value = if ($inp -match '(?i)value="([^"]*)"')    { $Matches[1] } else { "" }
                if ($name -and $itype -in @('hidden','submit','email','text','checkbox')) { $fields[$name] = $value }
            }
            # <button name="..."> ebenfalls erfassen
            [regex]::Matches($content, '(?i)<button[^>]+name="([^"]+)"[^>]*(?:value="([^"]*)")?') | ForEach-Object {
                $bname = $_.Groups[1].Value; $bval = $_.Groups[2].Value
                if ($bname) { $fields[$bname] = $bval }
            }

            $postResp = Invoke-WebRequest -Uri $actionUrl -Method POST -Body $fields -TimeoutSec 20 -UseBasicParsing -WebSession $sess -Headers $reqHdr
            if ($postResp.Content -match $confirmRx) {
                return @{ method = "post_form"; message = "Formular abgeschickt - Seite bestätigt Abmeldung" }
            }
            return @{ method = "post_form"; message = "Formular abgeschickt (keine explizite Bestätigung - bitte prüfen)" }
        }

        # Abmelde-Link auf der Seite suchen ("Hier klicken zum Abmelden" o.ä.)
        $linkRx = '(?i)<a[^>]+href="(https?://[^"]+)"[^>]*>[^<]{0,80}(?:unsubscribe|abmeld|opt.?out|austragen|abbestell)[^<]{0,40}</a>'
        if ($html -match $linkRx) {
            try {
                $linkUrl  = $Matches[1]
                $linkResp = Invoke-WebRequest -Uri $linkUrl -Method GET -TimeoutSec 15 -UseBasicParsing -Headers $reqHdr
                if ($linkResp.Content -match $confirmRx) {
                    return @{ method = "link_click"; message = "Abmelde-Link geklickt - Seite bestätigt Abmeldung" }
                }
                return @{ method = "link_click"; message = "Abmelde-Link geklickt (Bestätigung unklar - bitte prüfen)" }
            } catch {
                Write-Log "  -> Abmelde-Link fehlgeschlagen: $_" "WARN"
            }
        }

        return @{ method = "get"; message = "Seite aufgerufen, kein Abmelde-Formular/Link gefunden - bitte manuell prüfen" }
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
                # Neue Felder bevorzugen; Fallback für ältere Queue-Einträge
                $unsub_http    = if ($Entry.unsubscribe_http)    { "$($Entry.unsubscribe_http)"   } else { "" }
                $unsub_mailto  = if ($Entry.unsubscribe_mailto)  { "$($Entry.unsubscribe_mailto)" } else { "" }
                $one_click     = [bool]$Entry.unsubscribe_one_click
                if (-not $unsub_http -and -not $unsub_mailto -and $Entry.unsubscribe_header) {
                    $h = $Entry.unsubscribe_header
                    if ($h -match '<mailto:([^>]+)>') { $unsub_mailto = $Matches[1] }
                    if ($h -match '<(https?://[^>]+)>') { $unsub_http = $Matches[1] }
                }

                if ($unsub_http) {
                    $result = Invoke-WebUnsubscribe -Url $unsub_http -OneClick $one_click
                    Write-Log "  -> Web-Abmeldung ($($result.method)): $($result.message)"
                } elseif ($unsub_mailto) {
                    $parts = $unsub_mailto -split '\?', 2
                    $addr  = $parts[0]
                    $subj  = "unsubscribe"
                    if ($parts.Count -gt 1 -and $parts[1] -match '(?i)subject=([^&]+)') {
                        try { $subj = [Uri]::UnescapeDataString($Matches[1]) } catch { $subj = $Matches[1] }
                    }
                    $m = $Outlook.CreateItem(0); $m.To = $addr; $m.Subject = $subj; $m.Body = "unsubscribe"; $m.Send()
                    Write-Log "  -> Abmelde-Email gesendet an $addr (Betreff: $subj)"
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
            'security_threat' {
                # Email liegt bereits in EmailProcessor/Sicherheitswarnung (beim Scannen verschoben)
                Write-Log "  -> Sicherheitswarnung bestaetigt, bleibt in $($Entry.target_folder)"
                return "done"
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
            $how = if ($entry.auto_approved) { "Automatisch freigegeben" } else { "Genehmigt" }
            Write-Log "$how - fuehre aus: `"$($entry.subject)`""
            $result = Execute-Entry $entry $Namespace $Inbox $Outlook
            $entry.status = $result
            $changed = $true
            Record-Email $entry $result
            Write-TrainingDecision $Config $entry $(if ($entry.auto_approved) { 'auto_approved' } else { 'approved' }) $entry.category $entry.hotel $result
        }
        elseif ($entry.status -eq "ignored") {
            Write-Log "Ignoriert: `"$($entry.subject)`""
            Write-TrainingDecision $Config $entry 'ignored' '' ''
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
                Write-TrainingDecision $Config $entry 'corrected' $entry.correct_category $entry.correct_hotel
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
                # Korrigierte Eintraege nie automatisch freigeben
                $entry | Add-Member -NotePropertyName auto_eligible   -NotePropertyValue $false                     -Force
                $entry | Add-Member -NotePropertyName auto_approved   -NotePropertyValue $false                     -Force
                $entry | Add-Member -NotePropertyName auto_note       -NotePropertyValue "korrigierter Eintrag"     -Force
                $entry | Add-Member -NotePropertyName correct_category -NotePropertyValue $null -Force
                $entry | Add-Member -NotePropertyName correct_hotel    -NotePropertyValue $null -Force
                $entry | Add-Member -NotePropertyName rejection_reason -NotePropertyValue $null -Force
                $entry.status = "pending"
                Write-Log "  -> Zurueck in Warteschlange fuer erneute Genehmigung"
            } else {
                # Keine Korrektur - einfach archivieren
                Write-TrainingDecision $Config $entry 'rejected' '' ''
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
              categories=@{newsletter=0;application=0;advertisement=0;booking=0;unknown=0;error=0};
              hotels=@{}; recent=@() }
    if (-not (Test-Path $statsFile)) { return $def }
    try {
        $j = Get-Content $statsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $s = @{ total_processed=[int]$j.total_processed; today_processed=[int]$j.today_processed;
                last_date=[string]$j.last_date; last_checked=[string]$j.last_checked;
                categories=@{}; hotels=@{}; recent=@() }
        foreach ($k in @('newsletter','application','advertisement','booking','sonstiges','unknown','error')) {
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
              corrected=([bool]$Entry.is_corrected); auto=([bool]$Entry.auto_approved) }
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

# ── Sicherheitsprüfung: Anhänge + Links + Phishing ───────────────────
function Test-EmailSecurity {
    # -RulesOnly: nur Anhang-/Link-Regeln, ohne Ollama (bei Langdock kommt die Phishing-Einschaetzung aus der Analyse)
    param($Item, $Config, [switch]$RulesOnly)
    $reasons = @()
    $threatLevel = 'none'

    # 1. Gefährliche Anhänge prüfen
    $dangerExts = @('.exe','.bat','.cmd','.com','.pif','.scr','.vbs','.vbe','.js','.jse','.wsf','.wsh','.ps1','.ps2','.msi','.msp','.jar','.reg','.hta','.lnk','.iso','.img')
    $warnExts   = @('.zip','.rar','.7z','.gz','.tar','.docm','.xlsm','.pptm','.dotm','.xltm','.xlam')
    try {
        foreach ($att in $Item.Attachments) {
            $ext = [IO.Path]::GetExtension($att.FileName).ToLower()
            if ($ext -in $dangerExts) {
                $reasons += "Gefaehrlicher Anhang: $($att.FileName)"
                $threatLevel = 'danger'
            } elseif ($ext -in $warnExts) {
                $reasons += "Verdaechtiger Anhang (Makro/Archiv): $($att.FileName)"
                if ($threatLevel -eq 'none') { $threatLevel = 'warn' }
            }
        }
    } catch {}

    # 2. Verdächtige Links im Body prüfen
    $body = if ($Item.Body) { $Item.Body } else { "" }
    if ($body -match 'https?://\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}') {
        $reasons += "Link zu IP-Adresse (kein Domainname)"
        if ($threatLevel -eq 'none') { $threatLevel = 'warn' }
    }
    $urlMatches = [regex]::Matches($body, 'https?://[^\s<>"]+')
    foreach ($m in $urlMatches) {
        $url = $m.Value
        if ($url.Length -gt 200) {
            $reasons += "Sehr langer/verdaechtiger Link ($($url.Substring(0,60))...)"
            if ($threatLevel -eq 'none') { $threatLevel = 'warn' }
            break
        }
        if ($url -match '(?i)(verify.{0,10}account|confirm.{0,10}password|update.{0,10}billing|paypal.{0,20}\.(?!com)|microsoft.{0,20}\.(?!com)|apple.{0,20}\.(?!com)|secure.{0,10}login\.)') {
            $reasons += "Phishing-Muster in URL: $($url.Substring(0,[Math]::Min(80,$url.Length)))"
            $threatLevel = 'danger'
        }
    }

    # 3. LLM-Phishing-Analyse (nur wenn kein 'danger' durch Regeln und Ollama konfiguriert)
    if (-not $RulesOnly -and $threatLevel -ne 'danger' -and $Config.ollama_model) {
        try {
            $snippet = if ($body.Length -gt 500) { $body.Substring(0,500) } else { $body }
            $prompt = "Analysiere diese E-Mail auf Phishing, Betrug oder Schadsoftware. Antworte NUR mit JSON.`n`nBetreff: $($Item.Subject)`nAbsender: $($Item.SenderEmailAddress)`nInhalt: $snippet`n`nFormat: {`"is_phishing`": false, `"confidence`": `"low`", `"reason`": `"`"}"
            $reqBody = @{
                model    = $Config.ollama_model
                messages = @(@{ role = "user"; content = $prompt })
                stream   = $false; format = "json"; options = @{ temperature = 0.0 }
            } | ConvertTo-Json -Depth 10 -Compress
            $ollamaUrl = if ($Config.ollama_url) { $Config.ollama_url } else { "http://localhost:11434/api/chat" }
            $resp   = Invoke-RestMethod -Uri $ollamaUrl -Method POST `
                -Body ([Text.Encoding]::UTF8.GetBytes($reqBody)) `
                -ContentType "application/json; charset=utf-8" -TimeoutSec 60
            $result = $resp.message.content | ConvertFrom-Json
            if ($result.is_phishing -eq $true) {
                $reasons += "LLM-Phishing-Erkennung ($($result.confidence)): $($result.reason)"
                if ($result.confidence -in @('high','medium')) { $threatLevel = 'danger' }
                elseif ($threatLevel -eq 'none') { $threatLevel = 'warn' }
            }
        } catch { Write-Log "  -> Sicherheitsanalyse (LLM) fehlgeschlagen: $_" "WARN" }
    }

    return @{ is_threat = ($threatLevel -ne 'none'); threat_level = $threatLevel; reasons = $reasons }
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

    $useLangdock    = (Get-LlmProvider $Config) -eq 'langdock'
    $ollamaFallback = if ($Config.langdock -and $null -ne $Config.langdock.fallback_to_ollama) { [bool]$Config.langdock.fallback_to_ollama } else { $true }

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
                body_preview = if ($mail.Body) { $b = $mail.Body.Trim(); if ($b.Length -gt 2000) { $b.Substring(0,2000) + "…" } else { $b } } else { "" }
            }
            Add-To-Queue $taskEntry
            $processedKeys[$dedupKey] = $true
            Write-Log "  -> Rot markiert + als Aufgabe gespeichert"
            continue
        }

        # Ausstehende Genehmigungen VOR dem KI-Aufruf verarbeiten
        # (damit genehmigte Emails nicht durch laufende Klassifizierung blockiert werden)
        Process-Queue $Namespace $Inbox $Outlook $Config

        # Stabile Kennung + Trainingsdaten VOR dem Verschieben erfassen
        $messageId  = Get-InternetMessageId $mail
        $trainInput = if (Test-TrainingEnabled $Config) { Get-TrainingInput $mail $Config $messageId } else { $null }
        $prediction = $null
        $extra      = [ordered]@{}

        # Sicherheitsprüfung: Schadsoftware, Phishing, gefährliche Anhänge
        # Bei Langdock: erst die Regeln, dann EINE Langdock-Abfrage (liefert auch die Phishing-Einschaetzung)
        $secCheck = Test-EmailSecurity $mail $Config -RulesOnly:$useLangdock
        if ($useLangdock -and $secCheck.threat_level -ne 'danger') {
            Write-Log "[$i/$($mails.Count)] Analysiere mit Langdock: `"$($mail.Subject)`" | $($mail.SenderName)"
            try {
                $prediction = Invoke-LangdockAnalysis $mail $Config
            } catch {
                Write-Log "  -> Langdock-Fehler: $(Get-HttpErrorDetail $_)" "ERROR"
            }
            if ($prediction) {
                $secCheck = Merge-LlmPhishing $secCheck $prediction
            } elseif ($ollamaFallback) {
                Write-Log "  -> Weiter mit Ollama (Fallback)" "WARN"
                $secCheck = Test-EmailSecurity $mail $Config
            } else {
                Write-Log "  -> Bleibt im Posteingang, neuer Versuch beim naechsten Scan" "WARN"
                continue
            }
        }

        if ($secCheck.is_threat) {
            $reasonStr = $secCheck.reasons -join " | "
            Write-Log "[$i/$($mails.Count)] SICHERHEITSWARNUNG ($($secCheck.threat_level.ToUpper())): `"$($mail.Subject)`" | $($mail.SenderName)" "WARN"
            Write-Log "  -> Gruende: $reasonStr" "WARN"
            try {
                $entry = Build-QueueEntry $mail 'security_threat' '' $Config
                $entry.proposed_action  = "SICHERHEITSWARNUNG ($($secCheck.threat_level)): $reasonStr"
                $entry.target_folder    = "EmailProcessor/Sicherheitswarnung"
                $entry.security_warning = $reasonStr
                $entry.threat_level     = $secCheck.threat_level
                $entry.message_id       = $messageId
                $moved = $mail.Move((Resolve-Folder $Inbox "EmailProcessor/Sicherheitswarnung"))
                $entry.entry_id = $moved.EntryID
                Add-To-Queue $entry
                $processedKeys[$dedupKey] = $true
                Write-Log "  -> Verschoben nach: EmailProcessor/Sicherheitswarnung"
                $extra.security_reasons = @($secCheck.reasons)
                $extra.threat_level     = $secCheck.threat_level
                Write-TrainingPrediction $Config $entry $trainInput $prediction $extra
            } catch { Write-Log "  -> Fehler beim Sicherheits-Handling: $_" "ERROR" }
            continue
        }

        # Verdächtige Spam-Emails: Ollama überspringen, direkt als Werbung einreihen
        # (bei vorhandener Langdock-Analyse wird diese genutzt - die Regel ist nur ein Merkmal fuer die Trainingsdaten)
        $isSpam = Is-SuspiciousSpam $mail
        $extra.spam_heuristic = [bool]$isSpam
        if ($isSpam -and -not $prediction) {
            Write-Log "[$i/$($mails.Count)] Spam-Absender erkannt: `"$($mail.Subject)`" | $($mail.SenderName)"
            try {
                $entry = Build-QueueEntry $mail 'advertisement' '' $Config
                $entry.proposed_action = "Archivieren (verdaechtiger Spam-Absender — kein echter Name gesetzt)"
                $entry.message_id = $messageId
                Set-AutoDecision $entry $null $Config
                $moved = $mail.Move((Resolve-Folder $Inbox "EmailProcessor/Zur Pruefung"))
                $entry.entry_id = $moved.EntryID
                Add-To-Queue $entry
                $processedKeys[$dedupKey] = $true
                Write-Log "  -> Als Werbung/Spam eingereiht (Ollama uebersprungen)"
                Write-TrainingPrediction $Config $entry $trainInput $null $extra
            } catch { Write-Log "  -> Fehler beim Spam-Handling: $_" "ERROR" }
            continue
        }

        try {
            if (-not $prediction) {
                Write-Log "[$i/$($mails.Count)] Klassifiziere: `"$($mail.Subject)`" | $($mail.SenderName)"
                $cls = Classify-Email $mail $Config
                $prediction = [ordered]@{
                    provider = 'ollama'; model = "$($Config.ollama_model)"
                    category = $cls.category; hotel = "$($cls.hotel)"; confidence = $cls.confidence; reason = $cls.reason
                    language = ""
                }
            } elseif ($Config.training_data.compare_ollama -and $trainInput) {
                # Optional: Ollama parallel mitlaufen lassen, um es spaeter mit Langdock vergleichen zu koennen
                $cls = Classify-Email $mail $Config
                $extra.ollama_prediction = [ordered]@{ model = "$($Config.ollama_model)"; category = $cls.category; hotel = "$($cls.hotel)"; confidence = $cls.confidence }
            }

            $entry = Build-QueueEntry $mail $prediction.category $prediction.hotel $Config $prediction.language
            $entry.security_status = 'clean'
            $entry.message_id      = $messageId
            $entry.llm_provider    = $prediction.provider
            $entry.confidence      = $prediction.confidence
            $entry.llm_reason      = $prediction.reason
            Set-AutoDecision $entry $prediction $Config

            # Erst verschieben, dann neue EntryID speichern (EntryID aendert sich beim Verschieben)
            $moved = $mail.Move((Resolve-Folder $Inbox "EmailProcessor/Zur Pruefung"))
            $entry.entry_id = $moved.EntryID
            Add-To-Queue $entry
            $processedKeys[$dedupKey] = $true
            Write-TrainingPrediction $Config $entry $trainInput $prediction $extra

            if ($entry.auto_approved) {
                Write-Log "  -> Automatisch freigegeben ($($entry.auto_note)): $($entry.proposed_action)"
            } elseif ($entry.auto_mode -eq 'shadow' -and $entry.auto_eligible) {
                Write-Log "  -> Testbetrieb: wuerde automatisch ausgefuehrt - wartet auf Genehmigung: $($entry.proposed_action)"
            } else {
                Write-Log "  -> Wartet auf Genehmigung: $($entry.proposed_action)$(if ($entry.auto_note) { " (manuell: $($entry.auto_note))" })"
            }
        } catch {
            Write-Log "  -> Fehler: $_" "ERROR"
        }
    }
}

# ── Hauptschleife ─────────────────────────────────────────────────────
# Sicherstellen dass keine zweite Instanz läuft
$myPid    = $PID
$otherBot = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.ProcessId -ne $myPid -and $_.CommandLine -like '*-File*bot.ps1*' -and $_.CommandLine -notlike '*-NonInteractive*' -and $_.CommandLine -notlike '*-Command*' }
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

if ((Get-LlmProvider $config) -eq 'langdock') {
    $keyInfo = if (Get-LangdockApiKey $config) { "API-Key gefunden" } else { "KEIN API-Key ($(Get-LangdockKeyName $config))" }
    Write-Log "KI: Langdock ($(Get-LangdockModelLabel $config)) | $keyInfo | Ollama-Fallback: $(if ($config.langdock.fallback_to_ollama -eq $false) { 'aus' } else { $config.ollama_model })"
} else {
    Write-Log "KI: Ollama ($($config.ollama_model))"
}
Write-Log "Auto-Freigabe: $(Get-AutoApproveMode $config) | Trainingsdaten: $(if (Test-TrainingEnabled $config) { 'an' } else { 'aus' })"
Write-Log "Scanintervall: $($config.check_interval_minutes) Min. | Queue-Check: alle 10 Sek."

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
