# Dashboard HTTP-Server — kein Admin noetig (localhost only)
$scriptDir = $PSScriptRoot
$wwwDir    = Join-Path $scriptDir "www"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:5000/")
try {
    $listener.Start()
} catch {
    Write-Host "FEHLER: Port 5000 bereits belegt oder kein Zugriff." -ForegroundColor Red
    Write-Host "Schliesse andere Programme auf Port 5000 und starte erneut."
    pause; exit 1
}

Write-Host ""
Write-Host "  Dashboard laeuft auf http://localhost:5000" -ForegroundColor Green
Write-Host "  Strg+C zum Beenden"
Write-Host ""

function Send-Text {
    param($Ctx, [string]$Body, [string]$Mime = "text/html; charset=utf-8", [int]$Code = 200)
    $r = $Ctx.Response; $r.StatusCode = $Code; $r.ContentType = $Mime
    $b = [Text.Encoding]::UTF8.GetBytes($Body)
    $r.ContentLength64 = $b.Length
    $r.OutputStream.Write($b, 0, $b.Length)
    $r.OutputStream.Close()
}

function Send-File {
    param($Ctx, [string]$Path)
    $ext  = [IO.Path]::GetExtension($Path)
    $mime = switch ($ext) { '.html'{'text/html; charset=utf-8'} '.js'{'application/javascript'} '.css'{'text/css'} default{'text/plain'} }
    $b = [IO.File]::ReadAllBytes($Path)
    $r = $Ctx.Response; $r.StatusCode = 200; $r.ContentType = $mime; $r.ContentLength64 = $b.Length
    $r.OutputStream.Write($b, 0, $b.Length)
    $r.OutputStream.Close()
}

function Save-Queue-File {
    param($Path, $Queue)
    $parts = @($Queue | Where-Object { $null -ne $_ } | ForEach-Object {
        ConvertTo-Json -InputObject $_ -Depth 8 -Compress
    })
    $json = if ($parts.Count -eq 0) { '[]' } else { "[`n" + ($parts -join ",`n") + "`n]" }
    [IO.File]::WriteAllText($Path, $json, [Text.Encoding]::UTF8)
}

function Load-Queue-File {
    param($Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        $parsed = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $q = [System.Collections.Queue]::new(); $q.Enqueue($parsed)
        $all = @()
        while ($q.Count -gt 0) {
            $cur = $q.Dequeue()
            if ($null -eq $cur) { continue }
            if ($cur -is [System.Collections.IEnumerable] -and $cur -isnot [string]) {
                foreach ($i in $cur) { $q.Enqueue($i) }
            } elseif ($cur.PSObject.Properties.Name -contains 'id') { $all += $cur }
            elseif ($cur.PSObject.Properties.Name -contains 'value') { $q.Enqueue($cur.value) }
        }
        return $all
    } catch { return @() }
}

function Is-BotRunning {
    $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
        if ($p.CommandLine -like "*bot.ps1*") { return $true }
    }
    return $false
}

while ($true) {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $url = $req.Url.LocalPath

    try {
        switch -Wildcard ($url) {

            "/api/stats" {
                $f = Join-Path $scriptDir "stats.json"
                $json = if (Test-Path $f) { Get-Content $f -Raw -Encoding UTF8 } else { '{}' }
                $data = $json | ConvertFrom-Json
                $data | Add-Member -NotePropertyName bot_running -NotePropertyValue (Is-BotRunning) -Force
                # Config lesen (test_mode + ollama_url)
                $cfgF = Join-Path $scriptDir "config.json"
                $testMode = $false; $ollamaUrl = "http://localhost:11434/api/tags"
                if (Test-Path $cfgF) {
                    try {
                        $cfg = Get-Content $cfgF -Raw -Encoding UTF8 | ConvertFrom-Json
                        $testMode = [bool]$cfg.test_mode
                        if ($cfg.ollama_url) { $ollamaUrl = $cfg.ollama_url -replace '/api/chat','/api/tags' }
                    } catch {}
                }
                $data | Add-Member -NotePropertyName test_mode -NotePropertyValue $testMode -Force
                # Ollama-Erreichbarkeit pruefen
                $ollamaOk = $false
                try { Invoke-RestMethod -Uri $ollamaUrl -TimeoutSec 2 -ErrorAction Stop | Out-Null; $ollamaOk = $true } catch {}
                $data | Add-Member -NotePropertyName ollama_ok -NotePropertyValue $ollamaOk -Force
                Send-Text $ctx ($data | ConvertTo-Json -Depth 5) "application/json; charset=utf-8"
            }

            "/api/logs" {
                $f = Join-Path $scriptDir "email_processor.log"
                $lines = if (Test-Path $f) { @(Get-Content $f -Tail 40 -Encoding UTF8 -ErrorAction SilentlyContinue) } else { @() }
                Send-Text $ctx ($lines | ConvertTo-Json) "application/json; charset=utf-8"
            }

            "/api/queue" {
                $f = Join-Path $scriptDir "queue.json"
                $all = @(Load-Queue-File $f)
                $pending = @($all | Where-Object { $_.status -in @("pending","approved","rejected") } | Sort-Object {
                    try { [DateTime]::ParseExact($_.received, "dd.MM.yyyy HH:mm", $null) } catch { [DateTime]::MinValue }
                } -Descending)
                Send-Text $ctx (ConvertTo-Json -InputObject $pending -Depth 8) "application/json; charset=utf-8"
            }

            "/api/queue/approve" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) {
                        $e.status = "approved"
                        if ($data.reply_text) { $e | Add-Member -NotePropertyName reply_text -NotePropertyValue $data.reply_text -Force }
                        break
                    }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/queue/reject" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) {
                        $e.status = "rejected"
                        if ($data.rejection_reason)  { $e | Add-Member -NotePropertyName rejection_reason  -NotePropertyValue $data.rejection_reason  -Force }
                        if ($data.correct_category) { $e | Add-Member -NotePropertyName correct_category -NotePropertyValue $data.correct_category -Force }
                        if ($data.correct_hotel)    { $e | Add-Member -NotePropertyName correct_hotel    -NotePropertyValue $data.correct_hotel    -Force }
                        break
                    }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/queue/ignore" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) { $e.status = "ignored"; break }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/errors" {
                $f = Join-Path $scriptDir "queue.json"
                $all = @(Load-Queue-File $f)
                $errs = @($all | Where-Object { $_.status -eq "error" } | Sort-Object {
                    try { [DateTime]::ParseExact($_.received, "dd.MM.yyyy HH:mm", $null) } catch { [DateTime]::MinValue }
                } -Descending)
                Send-Text $ctx (ConvertTo-Json -InputObject $errs -Depth 8) "application/json; charset=utf-8"
            }

            "/api/errors/dismiss" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) { $e.status = "dismissed"; break }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/errors/retry" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) { $e.status = "approved"; break }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/tasks" {
                $f = Join-Path $scriptDir "queue.json"
                $all = @(Load-Queue-File $f)
                $tasks = @($all | Where-Object { $_.category -eq "task" -and $_.status -eq "open_task" } | Sort-Object {
                    try { [DateTime]::ParseExact($_.received, "dd.MM.yyyy HH:mm", $null) } catch { [DateTime]::MinValue }
                } -Descending)
                Send-Text $ctx (ConvertTo-Json -InputObject $tasks -Depth 8) "application/json; charset=utf-8"
            }

            "/api/tasks/done" {
                $body  = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data  = $body | ConvertFrom-Json
                $f     = Join-Path $scriptDir "queue.json"
                $queue = @(Load-Queue-File $f)
                foreach ($e in $queue) {
                    if ($e.id -eq $data.id) { $e.status = "task_done"; break }
                }
                Save-Queue-File $f $queue
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/corrections" {
                $f = Join-Path $scriptDir "corrections.json"
                if (Test-Path $f) {
                    Send-Text $ctx (Get-Content $f -Raw -Encoding UTF8) "application/json; charset=utf-8"
                } else {
                    Send-Text $ctx "[]" "application/json; charset=utf-8"
                }
            }

            "/api/email/open" {
                $body = (New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)).ReadToEnd()
                $data = $body | ConvertFrom-Json
                try {
                    $ol  = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
                    $ns  = $ol.GetNamespace("MAPI")
                    $item = $ns.GetItemFromID($data.entry_id)
                    $item.Display()
                    Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
                } catch {
                    Send-Text $ctx "{`"success`":false,`"error`":`"$($_ -replace '"',"'")`"}" "application/json; charset=utf-8"
                }
            }

            "/api/bot/start" {
                $botScript = Join-Path $scriptDir "bot.ps1"
                Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$botScript`""
                Start-Sleep -Seconds 2
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/bot/stop" {
                $procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*bot.ps1*" }
                foreach ($p in $procs) { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
                Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
            }

            "/api/config" {
                $f = Join-Path $scriptDir "config.json"
                if ($req.HttpMethod -eq "GET") {
                    Send-Text $ctx (Get-Content $f -Raw -Encoding UTF8) "application/json; charset=utf-8"
                } else {
                    $body = (New-Object IO.StreamReader($req.InputStream)).ReadToEnd()
                    # Validierung: muss gueltiges JSON sein
                    $body | ConvertFrom-Json | Out-Null
                    $body | Set-Content $f -Encoding UTF8
                    Send-Text $ctx '{"success":true}' "application/json; charset=utf-8"
                }
            }

            "/" { Send-File $ctx (Join-Path $wwwDir "index.html") }

            default {
                $file = Join-Path $wwwDir $url.TrimStart('/')
                if (Test-Path $file) { Send-File $ctx $file }
                else { Send-Text $ctx "404 Not Found" "text/plain" 404 }
            }
        }
    } catch {
        $errMsg = "$_" -replace '\\', '/' -replace '"', "'"
        try { Send-Text $ctx "{`"error`":`"$errMsg`"}" "application/json; charset=utf-8" 500 } catch {}
    }
}
