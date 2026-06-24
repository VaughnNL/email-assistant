param(
    [int]    $Hours  = 0,
    [int]    $Count  = 0,
    [switch] $Unread,
    [string] $Output = "emails",
    [string] $Folder = "Inbox",
    [switch] $Help
)

if ($Help) {
    Write-Host ""
    Write-Host "Outlook Email Downloader" -ForegroundColor Cyan
    Write-Host "========================" -ForegroundColor Cyan
    Write-Host "Downloads emails from your local Outlook and saves them as .msg files"
    Write-Host "organised into subfolders by date (e.g. emails\2026-06-19\)."
    Write-Host ""
    Write-Host "USAGE" -ForegroundColor Yellow
    Write-Host "  .\outlook_downloader.ps1 [parameters]"
    Write-Host ""
    Write-Host "PARAMETERS" -ForegroundColor Yellow
    Write-Host "  -Hours  N      Download emails received in the past N hours."
    Write-Host "  -Count  N      Download the last N emails (newest first)."
    Write-Host "  -Unread        Only include unread emails."
    Write-Host "  -Output PATH   Folder to save emails into. Default: .\emails"
    Write-Host "  -Folder NAME   Outlook folder to pull from. Default: Inbox"
    Write-Host "                 Other options: 'Sent Items', 'Drafts', 'Deleted', 'Junk'"
    Write-Host "  -Help          Show this help message."
    Write-Host ""
    Write-Host "NOTES" -ForegroundColor Yellow
    Write-Host "  - At least one of -Hours, -Count, or -Unread is required."
    Write-Host "  - Filters combine: -Hours 24 -Unread finds unread emails in the past 24 hours."
    Write-Host "  - -Count always picks the newest emails first."
    Write-Host "  - Emails are saved in native .msg format - double-click to open in Outlook."
    Write-Host "  - Microsoft Outlook must be installed and signed in."
    Write-Host ""
    Write-Host "EXAMPLES" -ForegroundColor Yellow
    Write-Host "  .\outlook_downloader.ps1 -Hours 18"
    Write-Host "  .\outlook_downloader.ps1 -Count 20"
    Write-Host "  .\outlook_downloader.ps1 -Unread"
    Write-Host "  .\outlook_downloader.ps1 -Hours 24 -Unread"
    Write-Host "  .\outlook_downloader.ps1 -Count 20 -Unread"
    Write-Host "  .\outlook_downloader.ps1 -Count 20 -Unread -Output C:\EmailBackup"
    Write-Host "  .\outlook_downloader.ps1 -Hours 12 -Folder 'Sent Items'"
    Write-Host ""
    exit 0
}

if (-not $Hours -and -not $Count -and -not $Unread) {
    Write-Host "ERROR: Specify at least one of -Hours, -Count, or -Unread" -ForegroundColor Red
    Write-Host "Run with -Help for full usage information."
    exit 1
}

function Sanitize-Filename([string]$Text, [int]$MaxLen = 80) {
    $clean = $Text -replace '[<>:"/\\|?*\r\n\t]', '_'
    $clean = $clean.Trim('. ')
    if (-not $clean) { $clean = "untitled" }
    if ($clean.Length -gt $MaxLen) { $clean = $clean.Substring(0, $MaxLen) }
    return $clean
}

Write-Host "Connecting to Outlook..." -ForegroundColor Cyan
try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $namespace.Logon()
} catch {
    Write-Host "ERROR: Could not connect to Outlook: $_" -ForegroundColor Red
    Write-Host "Make sure Microsoft Outlook is installed and you are signed in."
    exit 1
}

$folderMap = @{
    "inbox"      = 6
    "sent"       = 5
    "sent items" = 5
    "drafts"     = 16
    "deleted"    = 3
    "junk"       = 23
    "junk email" = 23
}

$folderKey = $Folder.Trim().ToLower()

if ($folderMap.ContainsKey($folderKey)) {
    $mailFolder = $namespace.GetDefaultFolder($folderMap[$folderKey])
} else {
    $mailFolder = $null
    foreach ($store in $namespace.Stores) {
        try {
            foreach ($sub in $store.GetRootFolder().Folders) {
                if ($sub.Name.Trim().ToLower() -eq $folderKey) {
                    $mailFolder = $sub
                    break
                }
            }
        } catch {}
        if ($mailFolder) { break }
    }
    if (-not $mailFolder) {
        Write-Host "Folder '$Folder' not found - falling back to Inbox." -ForegroundColor Yellow
        $mailFolder = $namespace.GetDefaultFolder(6)
    }
}

$items = $mailFolder.Items
$items.Sort("[ReceivedTime]", $true)

if ($Unread) {
    $items = $items.Restrict("[UnRead] = True")
    $items.Sort("[ReceivedTime]", $true)
}

$cutoff = $null
if ($Hours -gt 0) {
    $cutoff = (Get-Date).AddHours(-$Hours)
}

$outputPath = [System.IO.Path]::GetFullPath($Output)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

Write-Host ""
Write-Host ("Folder  : " + $mailFolder.Name)
if ($cutoff) { Write-Host ("Time    : past $Hours hour(s) since " + $cutoff.ToString("yyyy-MM-dd HH:mm")) }
if ($Count)  { Write-Host ("Limit   : $Count email(s)") }
if ($Unread) { Write-Host "Filter  : unread only" }
Write-Host ("Output  : $outputPath")
Write-Host ""

$downloaded    = 0
$scanned       = 0
$seen          = @{}
$consecutiveOld = 0

foreach ($msg in $items) {
    # -Count limit: only active when -Count was explicitly passed
    if ($Count -gt 0 -and $downloaded -ge $Count) { break }

    # Hard safety cap to prevent runaway scans on huge inboxes
    if ($scanned -ge 10000) {
        Write-Host "Scan limit reached (10,000 items) - stopping." -ForegroundColor Yellow
        break
    }
    $scanned++

    try {
        $received = $msg.ReceivedTime
    } catch {
        continue
    }

    # -Hours limit: only active when -Hours was explicitly passed.
    # Uses continue (not break) to handle any minor sort inconsistencies in the
    # COM collection after Restrict. Exits after 100 consecutive out-of-window
    # items so we don't scan the entire inbox needlessly.
    if ($cutoff -and $received -lt $cutoff) {
        $consecutiveOld++
        if ($consecutiveOld -ge 100) { break }
        continue
    }
    $consecutiveOld = 0

    try {
        $subject   = Sanitize-Filename $msg.Subject
        $sender    = Sanitize-Filename $msg.SenderName
        $ts        = $received.ToString("HHmmss")
        $dayLabel  = $received.ToString("yyyy-MM-dd")
        $dayFolder = Join-Path $outputPath $dayLabel

        New-Item -ItemType Directory -Path $dayFolder -Force | Out-Null

        if (-not $seen.ContainsKey($dayFolder)) {
            $seen[$dayFolder] = [System.Collections.Generic.HashSet[string]]::new()
        }

        $base     = "${ts}_${sender}_${subject}"
        $filename = "${base}.msg"
        $n = 1
        while ($seen[$dayFolder].Contains($filename)) {
            $filename = "${base}_${n}.msg"
            $n++
        }
        $seen[$dayFolder].Add($filename) | Out-Null

        $filepath = Join-Path $dayFolder $filename
        $msg.SaveAs($filepath, 3)

        $downloaded++

        $readFlag = ""
        if (-not $msg.UnRead) { $readFlag = " [read]" }

        $subj = if ($msg.Subject) { $msg.Subject } else { "(no subject)" }
        if ($subj.Length -gt 60) { $subj = $subj.Substring(0, 60) }

        Write-Host ("  [{0,4}] {1}  |  {2}{3}" -f $downloaded, $received.ToString("yyyy-MM-dd HH:mm"), $subj, $readFlag)

    } catch {
        Write-Host ("  Warning: could not save item - " + $_) -ForegroundColor Yellow
        continue
    }
}

Write-Host ""
Write-Host ("------------------------------------------------------------")
Write-Host ("Done. $downloaded email(s) saved to '$outputPath'")
if ($downloaded -eq 0) {
    Write-Host "No emails matched the given criteria." -ForegroundColor Yellow
}
