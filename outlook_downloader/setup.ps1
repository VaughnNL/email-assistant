Write-Host ""
Write-Host "Email Task Agent - Setup" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan
Write-Host ""

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$agentPath = Join-Path $scriptDir "email_task_agent.ps1"

if (-not (Test-Path $agentPath)) {
    Write-Host "ERROR: email_task_agent.ps1 not found at '$agentPath'" -ForegroundColor Red
    exit 1
}

# ----------------------------------------------------------------
# 1. Ollama
# ----------------------------------------------------------------
Write-Host "Step 1: Ollama" -ForegroundColor Yellow
Write-Host ""

$ollamaUrl   = "http://localhost:11434"
$defaultModel = "gemma4:latest"

# Check if Ollama is running
try {
    $tags = Invoke-RestMethod -Uri "$ollamaUrl/api/tags" -Method GET -ErrorAction Stop
    Write-Host "  Ollama is running at $ollamaUrl." -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Ollama is not running or not reachable at $ollamaUrl." -ForegroundColor Red
    Write-Host "  Start Ollama and re-run this setup, or install it from https://ollama.com" -ForegroundColor Yellow
    exit 1
}

# Check if default model is available
$availableModels = $tags.models | ForEach-Object { $_.name }
if ($availableModels -contains $defaultModel) {
    Write-Host "  Model '$defaultModel' is available." -ForegroundColor Green
} else {
    Write-Host "  Available models: $($availableModels -join ', ')"
    Write-Host "  Model '$defaultModel' not found." -ForegroundColor Yellow
    $pull = Read-Host "  Pull '$defaultModel' now? (Y/n)"
    if ($pull -ne 'n' -and $pull -ne 'N') {
        Write-Host "  Pulling '$defaultModel' (this may take a while)..." -ForegroundColor Cyan
        & ollama pull $defaultModel
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ERROR: Failed to pull model. You can do it manually: ollama pull $defaultModel" -ForegroundColor Red
            exit 1
        }
        Write-Host "  Model pulled successfully." -ForegroundColor Green
    } else {
        Write-Host "  Skipped. Edit -Model parameter in the scheduled task action if you use a different model." -ForegroundColor Yellow
    }
}

Write-Host ""

# ----------------------------------------------------------------
# 2. Task Scheduler
# ----------------------------------------------------------------
Write-Host "Step 2: Schedule daily 7am run" -ForegroundColor Yellow
Write-Host ""

$taskName  = "EmailTaskAgent"
$psExe     = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$arguments = "-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File `"$agentPath`""

$action   = New-ScheduledTaskAction -Execute $psExe -Argument $arguments
$trigger  = New-ScheduledTaskTrigger -Daily -At "07:00"
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

# Interactive logon is required so Outlook COM can access the mail profile
$principal = New-ScheduledTaskPrincipal `
    -UserId    ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel  Limited

# Remove old registration if it exists
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

try {
    Register-ScheduledTask `
        -TaskName    $taskName `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description "Daily 7am email review and task list update via Claude AI" | Out-Null

    Write-Host "  Task '$taskName' registered - runs daily at 7:00 AM." -ForegroundColor Green
} catch {
    Write-Host "  ERROR registering task: $_" -ForegroundColor Red
    Write-Host "  You may need to run this setup script as Administrator." -ForegroundColor Yellow
    exit 1
}

Write-Host ""

# ----------------------------------------------------------------
# Done
# ----------------------------------------------------------------
Write-Host "Setup complete!" -ForegroundColor Cyan
Write-Host ""
Write-Host "To test right now:"
Write-Host "  .\email_task_agent.ps1"
Write-Host ""
Write-Host "To check scheduled task status:"
Write-Host "  Get-ScheduledTask -TaskName EmailTaskAgent | Get-ScheduledTaskInfo"
Write-Host ""
Write-Host "To open Task Scheduler GUI:"
Write-Host "  taskschd.msc"
Write-Host ""
Write-Host "Briefs are saved to:"
Write-Host "  $env:USERPROFILE\Desktop\briefs\"
Write-Host ""
Write-Host "Logs are saved to:"
Write-Host "  $(Join-Path $scriptDir 'logs\')"
Write-Host ""
