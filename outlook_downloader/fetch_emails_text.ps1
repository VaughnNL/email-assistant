param([int]$Hours = 24)

try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $namespace.Logon()
} catch {
    Write-Error "Could not connect to Outlook: $_"
    exit 1
}

$cutoff = (Get-Date).AddHours(-$Hours)

function Get-FolderEmails {
    param([int]$FolderID, [DateTime]$Cutoff, [bool]$IsSent = $false)
    try {
        $folder     = $namespace.GetDefaultFolder($FolderID)
        $items      = $folder.Items
        $sortField  = if ($IsSent) { "[SentOn]" } else { "[ReceivedTime]" }
        $items.Sort($sortField, $true)

        $results        = [System.Collections.ArrayList]::new()
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
                if ($body -and $body.Length -gt 1000) { $body = $body.Substring(0, 1000) + "..." }
                $null = $results.Add([PSCustomObject]@{
                    Time   = $msgTime.ToString("yyyy-MM-dd HH:mm")
                    FromTo = $fromTo
                    Subject = $subject
                    Body   = $body
                })
            } catch { continue }
        }
        return ,$results
    } catch {
        return ,([System.Collections.ArrayList]::new())
    }
}

$inboxEmails = Get-FolderEmails -FolderID 6 -Cutoff $cutoff -IsSent $false
$sentEmails  = Get-FolderEmails -FolderID 5 -Cutoff $cutoff -IsSent $true

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("=== LOCAL EMAIL DATA (past $Hours hours as of $(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("INBOX ($($inboxEmails.Count) emails):")

foreach ($e in $inboxEmails) {
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("From:    $($e.FromTo)")
    $null = $sb.AppendLine("Time:    $($e.Time)")
    $null = $sb.AppendLine("Subject: $($e.Subject)")
    $null = $sb.AppendLine("Body:")
    $null = $sb.AppendLine($e.Body)
    $null = $sb.AppendLine("---")
}

$null = $sb.AppendLine("")
$null = $sb.AppendLine("SENT ($($sentEmails.Count) emails):")

foreach ($e in $sentEmails) {
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("To:      $($e.FromTo)")
    $null = $sb.AppendLine("Time:    $($e.Time)")
    $null = $sb.AppendLine("Subject: $($e.Subject)")
    $null = $sb.AppendLine("Body:")
    $null = $sb.AppendLine($e.Body)
    $null = $sb.AppendLine("---")
}

$null = $sb.AppendLine("")
$null = $sb.AppendLine("=== END EMAIL DATA ===")

Write-Output $sb.ToString()
