#Requires -RunAsAdministrator
#Requires -Version 5.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Teams

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "$PSScriptRoot\config\notification-config.json",
    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl,
    [Parameter(Mandatory = $false)]
    [string]$SmtpServer,
    [Parameter(Mandatory = $false)]
    [string]$FromAddress,
    [Parameter(Mandatory = $false)]
    [string[]]$ToAddress
)

# 設定ファイルの読み込み
function Initialize-NotificationConfig {
    try {
        if (-not (Test-Path $ConfigPath)) {
            $defaultConfig = @{
                Alerts = @{
                    StorageThreshold = 90
                    ErrorThreshold = 5
                    WarningThreshold = 10
                }
                Notification = @{
                    Email = @{
                        Enabled = $false
                        SmtpServer = $SmtpServer
                        FromAddress = $FromAddress
                        ToAddress = $ToAddress
                    }
                    Teams = @{
                        Enabled = $false
                        WebhookUrl = $TeamsWebhookUrl
                    }
                }
                AuditLog = @{
                    Enabled = $true
                    Path = "$PSScriptRoot\Logs\audit"
                    RetentionDays = 90
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
        Write-Error "Failed to initialize notification config: $_"
        throw
    }
}

# メール通知の送信
function Send-EmailAlert {
    param(
        [string]$Subject,
        [string]$Body
    )
    
    try {
        if (-not $Config.Notification.Email.Enabled) { return }
        
        $emailParams = @{
            SmtpServer = $Config.Notification.Email.SmtpServer
            From = $Config.Notification.Email.FromAddress
            To = $Config.Notification.Email.ToAddress
            Subject = $Subject
            Body = $Body
            BodyAsHtml = $true
            Encoding = [System.Text.Encoding]::UTF8
        }
        
        Send-MailMessage @emailParams
        Write-Log "Email alert sent successfully"
    }
    catch {
        Write-Error "Failed to send email alert: $_"
    }
}

# Teams通知の送信
function Send-TeamsAlert {
    param(
        [string]$Title,
        [string]$Message,
        [ValidateSet('Normal', 'Warning', 'Error')]
        [string]$Severity = 'Normal'
    )
    
    try {
        if (-not $Config.Notification.Teams.Enabled) { return }
        
        $color = switch ($Severity) {
            'Normal' { '00ff00' }
            'Warning' { 'ffff00' }
            'Error' { 'ff0000' }
        }
        
        $body = @{
            "@type" = "MessageCard"
            "@context" = "http://schema.org/extensions"
            "themeColor" = $color
            "summary" = $Title
            "sections" = @(
                @{
                    "activityTitle" = $Title
                    "activitySubtitle" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "text" = $Message
                }
            )
        }
        
        $params = @{
            Uri = $Config.Notification.Teams.WebhookUrl
            Method = 'POST'
            Body = ($body | ConvertTo-Json -Depth 10)
            ContentType = 'application/json'
        }
        
        Invoke-RestMethod @params
        Write-Log "Teams alert sent successfully"
    }
    catch {
        Write-Error "Failed to send Teams alert: $_"
    }
}

# 監査ログの記録
function Write-AuditLog {
    param(
        [string]$Action,
        [string]$Details,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    try {
        if (-not $Config.AuditLog.Enabled) { return }
        
        $auditDir = $Config.AuditLog.Path
        if (-not (Test-Path $auditDir)) {
            New-Item -ItemType Directory -Path $auditDir -Force | Out-Null
        }
        
        $logFile = Join-Path $auditDir "OneDriveAudit_$(Get-Date -Format 'yyyyMMdd').log"
        $logEntry = @{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Action = $Action
            Details = $Details
            Level = $Level
            User = $env:USERNAME
            Computer = $env:COMPUTERNAME
        }
        
        $logEntry | ConvertTo-Json | Add-Content -Path $logFile -Encoding UTF8
        
        # 古いログファイルの削除
        $retentionDate = (Get-Date).AddDays(-$Config.AuditLog.RetentionDays)
        Get-ChildItem $auditDir -Filter "OneDriveAudit_*.log" | 
            Where-Object { $_.CreationTime -lt $retentionDate } |
            Remove-Item -Force
    }
    catch {
        Write-Error "Failed to write audit log: $_"
    }
}

# アラート条件のチェック
function Test-AlertConditions {
    param(
        [Array]$UsageData,
        [Array]$ErrorData
    )
    
    try {
        $alerts = @()
        
        # ストレージ使用率のチェック
        foreach ($user in $UsageData) {
            if ($user.UsagePercentage -ge $Config.Alerts.StorageThreshold) {
                $alerts += @{
                    Type = 'Storage'
                    Severity = 'Warning'
                    Message = "ストレージ使用率が閾値を超えています: $($user.Name) ($($user.UsagePercentage)%)"
                }
            }
        }
        
        # エラー数のチェック
        $errorCount = ($ErrorData | Where-Object { $_.LevelDisplayName -eq 'Error' }).Count
        $warningCount = ($ErrorData | Where-Object { $_.LevelDisplayName -eq 'Warning' }).Count
        
        if ($errorCount -ge $Config.Alerts.ErrorThreshold) {
            $alerts += @{
                Type = 'Error'
                Severity = 'Error'
                Message = "エラー数が閾値を超えています: $errorCount件"
            }
        }
        
        if ($warningCount -ge $Config.Alerts.WarningThreshold) {
            $alerts += @{
                Type = 'Warning'
                Severity = 'Warning'
                Message = "警告数が閾値を超えています: $warningCount件"
            }
        }
        
        return $alerts
    }
    catch {
        Write-Error "Failed to check alert conditions: $_"
        return @()
    }
}

# メイン処理
try {
    Initialize-NotificationConfig
    
    # 監視データの取得
    $usageData = Get-OneDriveUsageReport
    $errorData = Get-OneDriveSyncErrors
    
    # アラート条件のチェック
    $alerts = Test-AlertConditions -UsageData $usageData -ErrorData $errorData
    
    foreach ($alert in $alerts) {
        # 監査ログの記録
        Write-AuditLog -Action "Alert" -Details $alert.Message -Level $alert.Severity
        
        # メール通知
        if ($Config.Notification.Email.Enabled) {
            Send-EmailAlert -Subject "OneDrive Alert: $($alert.Type)" -Body $alert.Message
        }
        
        # Teams通知
        if ($Config.Notification.Teams.Enabled) {
            Send-TeamsAlert -Title "OneDrive Alert: $($alert.Type)" -Message $alert.Message -Severity $alert.Severity
        }
    }
}
catch {
    Write-Error "Notification processing failed: $_"
    exit 1
}