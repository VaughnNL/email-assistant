# Outlook Email Downloader - convenience wrapper
# Usage examples:
#   .\run.ps1 --hours 18
#   .\run.ps1 --count 20
#   .\run.ps1 --unread
#   .\run.ps1 --hours 24 --unread
#   .\run.ps1 --count 20 --unread --output C:\EmailBackup
#   .\run.ps1 --hours 12 --folder "Sent Items"

param(
    [int]$hours,
    [int]$count,
    [switch]$unread,
    [string]$output = "emails",
    [string]$folder = "Inbox"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = "$scriptDir\outlook_downloader.py"

$args_list = @()
if ($hours)  { $args_list += "--hours",  $hours  }
if ($count)  { $args_list += "--count",  $count  }
if ($unread) { $args_list += "--unread"           }
if ($output) { $args_list += "--output", $output  }
if ($folder -ne "Inbox") { $args_list += "--folder", $folder }

if ($args_list.Count -eq 0 -and -not $unread) {
    Write-Host "Usage: .\run.ps1 [--hours N] [--count N] [--unread] [--output PATH] [--folder NAME]"
    exit 1
}

python $py @args_list
