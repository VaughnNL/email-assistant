param(
    [int]    $Hours     = 24,
    [string] $BriefFile = "",
    [string] $Model     = "gemma4:latest",
    [string] $OllamaUrl = "http://localhost:11434",
    [switch] $Force,
    [switch] $Help
)

$scriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$briefsDir  = Join-Path (Split-Path $scriptDir -Parent) "briefs"
New-Item -ItemType Directory -Path $briefsDir -Force | Out-Null
if (-not $BriefFile) { $BriefFile = Join-Path $briefsDir ((Get-Date -Format "MMdd") + "_Brief.txt") }

if ($Help) {
    Write-Host ""
    Write-Host "Daily Brief Agent" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    Write-Host "Reviews recent inbox and sent emails, then uses Claude AI to produce"
    Write-Host "a daily brief saved to your Desktop as 'Daily Brief.txt'."
    Write-Host ""
    Write-Host "USAGE" -ForegroundColor Yellow
    Write-Host "  .\email_task_agent.ps1 [parameters]"
    Write-Host ""
    Write-Host "PARAMETERS" -ForegroundColor Yellow
    Write-Host "  -Hours N         Review emails from the past N hours. Default: 24"
    Write-Host "  -BriefFile PATH  Output file path. Default: Desktop\Daily Brief.txt"
    Write-Host "  -Model NAME      Ollama model to use. Default: gemma4:latest"
    Write-Host "  -OllamaUrl URL   Ollama base URL. Default: http://localhost:11434"
    Write-Host "  -Force           Run even if a brief was already generated today"
    Write-Host "  -Help            Show this help message."
    Write-Host ""
    exit 0
}

# --- Skip if already run today (unless -Force) ---
if (-not $Force -and (Test-Path $BriefFile)) {
    $lastWrite = (Get-Item $BriefFile).LastWriteTime.Date
    if ($lastWrite -eq (Get-Date).Date) {
        Write-Host "Brief already generated today. Use -Force to regenerate."
        exit 0
    }
}

# --- Verify Ollama is reachable ---
try {
    $null = Invoke-RestMethod -Uri "$OllamaUrl/api/tags" -Method GET -ErrorAction Stop
} catch {
    Write-Host "ERROR: Ollama not reachable at $OllamaUrl. Make sure Ollama is running." -ForegroundColor Red
    exit 1
}

# --- Logging ---
$logDir = Join-Path $scriptDir "logs"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir ("agent_" + (Get-Date -Format "yyyy-MM-dd") + ".log")

function Log([string]$msg) {
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "--- Daily Brief Agent starting. Reviewing past $Hours hour(s). ---"

# --- Connect to Outlook ---
Log "Connecting to Outlook..."
try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    # Skip Logon — if Outlook is already running, COM reuses the existing session
} catch {
    Log "ERROR: Could not connect to Outlook: $_"
    exit 1
}

# --- Collect emails ---
function Get-EmailSummaries {
    param([int]$FolderID, [DateTime]$Cutoff, [bool]$IsSent = $false)
    try {
        $folder    = $namespace.GetDefaultFolder($FolderID)
        $items     = $folder.Items
        $sortField = if ($IsSent) { "[SentOn]" } else { "[ReceivedTime]" }
        $items.Sort($sortField, $true)

        $summaries      = [System.Collections.ArrayList]::new()
        $consecutiveOld = 0

        foreach ($msg in $items) {
            try { $msgTime = if ($IsSent) { $msg.SentOn } else { $msg.ReceivedTime } } catch { continue }
            if ($msgTime -lt $Cutoff) {
                $consecutiveOld++
                if ($consecutiveOld -ge 50) { break }
                continue
            }
            $consecutiveOld = 0
            try {
                $subject = if ($msg.Subject) { $msg.Subject } else { "(no subject)" }
                $fromTo  = if ($IsSent) { $msg.To } else { $msg.SenderName }
                $body    = $msg.Body
                if ($body -and $body.Length -gt 800) { $body = $body.Substring(0, 800) + "..." }
                $null = $summaries.Add([PSCustomObject]@{
                    Time    = $msgTime.ToString("yyyy-MM-dd HH:mm")
                    FromTo  = $fromTo
                    Subject = $subject
                    Body    = $body
                })
            } catch { continue }
        }
        return ,$summaries
    } catch {
        Log "Warning: Could not read folder (ID $FolderID): $_"
        return ,([System.Collections.ArrayList]::new())
    }
}

$cutoff = (Get-Date).AddHours(-$Hours)
Log "Cutoff: $($cutoff.ToString('yyyy-MM-dd HH:mm'))"

Log "Fetching inbox emails..."
$inboxEmails = Get-EmailSummaries -FolderID 6 -Cutoff $cutoff -IsSent $false

Log "Fetching sent emails..."
$sentEmails = Get-EmailSummaries -FolderID 5 -Cutoff $cutoff -IsSent $true

Log "Found $($inboxEmails.Count) inbox and $($sentEmails.Count) sent email(s)."

if ($inboxEmails.Count -eq 0 -and $sentEmails.Count -eq 0) {
    Log "No emails in this period. No brief to generate."
    exit 0
}

# --- Format email list ---
function Format-Emails([System.Collections.ArrayList]$Emails, [bool]$IsSent) {
    if ($Emails.Count -eq 0) { return "(none)" }
    $sb    = [System.Text.StringBuilder]::new()
    $label = if ($IsSent) { "To" } else { "From" }
    foreach ($e in $Emails) {
        $null = $sb.AppendLine("---")
        $null = $sb.AppendLine("Time:    $($e.Time)")
        $null = $sb.AppendLine("${label}: $($e.FromTo)")
        $null = $sb.AppendLine("Subject: $($e.Subject)")
        $null = $sb.AppendLine("Body:")
        $null = $sb.AppendLine($e.Body)
        $null = $sb.AppendLine("")
    }
    return $sb.ToString()
}

$inboxText = Format-Emails -Emails $inboxEmails -IsSent $false
$sentText  = Format-Emails -Emails $sentEmails  -IsSent $true

# --- Prompts ---
$systemPrompt = @"
You are a personal executive assistant. You read email summaries and produce a concise, well-structured daily brief as a plain text document.

OUTPUT FORMAT — output ONLY the brief content using this exact structure. No markdown symbols, no asterisks, no hash signs, no code fences. Plain readable text only.

DAILY BRIEF — YYYY-MM-DD
=========================

OVERVIEW
--------
2-3 sentences summarising the overall shape of the day's email activity — volume, dominant themes, anything urgent.

KEY THREADS
-----------
One entry per meaningful email thread. Include who it's from/to and one sentence on what's happening or what's needed.

  [Subject] (From: Name)
  What's happening / what's at stake.

ACTION ITEMS
------------
Specific things that need doing, ordered by urgency. Start each with a verb.

  > Action (from: email subject)

FYI — NO ACTION NEEDED
-----------------------
Emails that are informational only. Omit this section if there are none.

  [Subject] — one-line summary

RULES:
1. Be specific and concrete — no vague summaries
2. Only list action items that genuinely require a response or decision
3. If there are no action items, write "(none)" under that section
4. Output ONLY the brief content — no explanation, no preamble
"@

$userPrompt = @"
Today: $(Get-Date -Format 'yyyy-MM-dd HH:mm')

INBOX (past $Hours hours):
$inboxText

SENT (past $Hours hours):
$sentText

Write the daily brief now.
"@

# --- Call Ollama API ---
Log "Calling Ollama ($Model)..."

$requestBody = @{
    model    = $Model
    messages = @(
        @{ role = "system"; content = $systemPrompt }
        @{ role = "user";   content = $userPrompt }
    )
    stream   = $false
} | ConvertTo-Json -Depth 10

try {
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($requestBody)
    $response = Invoke-RestMethod `
        -Uri         "$OllamaUrl/api/chat" `
        -Method      POST `
        -ContentType "application/json; charset=utf-8" `
        -Body        $bodyBytes `
        -TimeoutSec  300
} catch {
    Log "ERROR: Ollama API call failed: $_"
    exit 1
}

$briefText = $response.message.content.Trim()

if (-not $briefText) {
    Log "ERROR: No text content in Ollama response."
    exit 1
}

# --- Save brief ---
Set-Content -Path $BriefFile -Value $briefText -Encoding UTF8
Log "Brief saved: '$BriefFile'"

$evalCount   = if ($response.eval_count)   { $response.eval_count }   else { "?" }
$promptCount = if ($response.prompt_eval_count) { $response.prompt_eval_count } else { "?" }
Log "Tokens: $promptCount prompt, $evalCount generated."

# --- Windows notification ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    $notify        = New-Object System.Windows.Forms.NotifyIcon
    $notify.Icon   = [System.Drawing.SystemIcons]::Information
    $notify.Visible = $true
    $notify.ShowBalloonTip(8000, "Daily Brief Ready", "Your brief for $(Get-Date -Format 'dd MMM') is on your Desktop: Daily Brief.txt", [System.Windows.Forms.ToolTipIcon]::Info)
    Start-Sleep -Seconds 2
    $notify.Dispose()
} catch {
    Log "Note: Could not show notification: $_"
}

Log "--- Done. ---"

