#Requires -RunAsAdministrator
#Requires -Version 5.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Reports

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "$PSScriptRoot\config\scheduler-config.json",
    [switch]$Register,
    [switch]$Unregister
)

# スケジューラー設定の読み込み
function Initialize-SchedulerConfig {
    try {
        if (-not (Test-Path $ConfigPath)) {
            $defaultConfig = @{
                Tasks = @(
                    @{
                        Name = "OneDriveMonitoring"
                        Description = "OneDrive監視タスク"
                        Script = "Monitor-OneDriveStatus.ps1"
                        Schedule = @{
                            Frequency = "Daily"
                            Time = "09:00"
                            Interval = 60 # minutes
                        }
                        Enabled = $true
                    },
                    @{
                        Name = "OneDriveNotification"
                        Description = "OneDrive通知タスク"
                        Script = "Send-OneDriveNotification.ps1"
                        Schedule = @{
                            Frequency = "Hourly"
                            Interval = 60 # minutes
                        }
                        Enabled = $true
                    },
                    @{
                        Name = "OneDriveDiagnostics"
                        Description = "OneDrive診断タスク"
                        Script = "Troubleshoot-OneDriveKFM.ps1"
                        Schedule = @{
                            Frequency = "Daily"
                            Time = "00:00"
                        }
                        Enabled = $true
                    }
                )
                Logging = @{
                    Path = "$PSScriptRoot\Logs\scheduler"
                    RetentionDays = 30
                }
                ErrorHandling = @{
                    MaxRetries = 3
                    RetryIntervalMinutes = 5
                    NotifyOnError = $true
                }
            }
            
            $configDir = Split-Path $ConfigPath -Parent
            if (-not (Test-Path $configDir)) {
                New-Item -ItemType Directory -Path $configDir -Force | Out-Null
            }
            
            $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content $ConfigPath -Encoding UTF8
        }
        
        $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Error "Failed to initialize scheduler config: $_"
        throw
    }
}

# タスクスケジューラーへの登録
function Register-ScheduledTasks {
    try {
        foreach ($task in $Config.Tasks) {
            if (-not $task.Enabled) { continue }

            $scriptPath = Join-Path $PSScriptRoot $task.Script
            if (-not (Test-Path $scriptPath)) {
                Write-Warning "Script not found: $scriptPath"
                continue
            }

            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

            $trigger = switch ($task.Schedule.Frequency) {
                "Daily" {
                    $time = [DateTime]::Parse($task.Schedule.Time)
                    New-ScheduledTaskTrigger -Daily -At $time
                }
                "Hourly" {
                    New-ScheduledTaskTrigger -Once -At (Get-Date) `
                        -RepetitionInterval (New-TimeSpan -Minutes $task.Schedule.Interval) `
                        -RepetitionDuration (New-TimeSpan -Days 1)
                }
            }

            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)

            Register-ScheduledTask -TaskName "OneDriveManager_$($task.Name)" `
                -Description $task.Description `
                -Action $action `
                -Trigger $trigger `
                -Principal $principal `
                -Settings $settings `
                -Force

            Write-Host "Registered task: $($task.Name)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to register scheduled tasks: $_"
    }
}

# タスクスケジューラーからの削除
function Unregister-ScheduledTasks {
    try {
        foreach ($task in $Config.Tasks) {
            $taskName = "OneDriveManager_$($task.Name)"
            if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                Write-Host "Unregistered task: $($task.Name)" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Error "Failed to unregister scheduled tasks: $_"
    }
}

# エラー発生時のリトライ処理
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$TaskName
    )

    $retryCount = 0
    $success = $false

    while (-not $success -and $retryCount -lt $Config.ErrorHandling.MaxRetries) {
        try {
            & $ScriptBlock
            $success = $true
        }
        catch {
            $retryCount++
            $error = $_
            Write-Warning "Task '$TaskName' failed (Attempt $retryCount of $($Config.ErrorHandling.MaxRetries)): $error"

            if ($retryCount -lt $Config.ErrorHandling.MaxRetries) {
                Start-Sleep -Seconds ($Config.ErrorHandling.RetryIntervalMinutes * 60)
            }
            else {
                if ($Config.ErrorHandling.NotifyOnError) {
                    $errorMessage = "Task '$TaskName' failed after $retryCount attempts: $error"
                    Send-ErrorNotification -Message $errorMessage
                }
                throw $error
            }
        }
    }
}

# エラー通知の送信
function Send-ErrorNotification {
    param([string]$Message)

    try {
        $notificationParams = @{
            Subject = "OneDriveManager Task Error"
            Body = $Message
        }
        & "$PSScriptRoot\Send-OneDriveNotification.ps1" @notificationParams
    }
    catch {
        Write-Error "Failed to send error notification: $_"
    }
}

# メイン処理
try {
    Initialize-SchedulerConfig

    if ($Register) {
        Write-Host "Registering scheduled tasks..." -ForegroundColor Cyan
        Register-ScheduledTasks
    }
    elseif ($Unregister) {
        Write-Host "Unregistering scheduled tasks..." -ForegroundColor Cyan
        Unregister-ScheduledTasks
    }
    else {
        Write-Host "Please specify -Register or -Unregister switch" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Scheduler configuration failed: $_"
    exit 1
}