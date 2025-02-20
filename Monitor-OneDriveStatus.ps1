#Requires -RunAsAdministrator
#Requires -Version 5.0
#Requires -Modules MSOnline, Microsoft.Graph.Authentication, Microsoft.Graph.Reports, Microsoft.Graph.Users

[CmdletBinding()]
param(
    [string]$OutputPath = "$PSScriptRoot\Reports",
    [string]$LogPath = "$PSScriptRoot\Logs"
)

# Ensure output directories exist
$null = New-Item -ItemType Directory -Force -Path $OutputPath
$null = New-Item -ItemType Directory -Force -Path $LogPath

$LogFile = Join-Path $LogPath "OneDriveMonitor_$(Get-Date -Format 'yyyyMMdd').log"

function Write-Log {
    param($Message)
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage
}

# Install required modules if not present
function Install-RequiredModules {
    $requiredModules = @('MSOnline', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports', 'Microsoft.Graph.Users')
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Log "Installing $module module..."
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -Force
    }
}

# Check if the current user is a Global Administrator
function Test-GlobalAdminRole {
    try {
        Connect-MgGraph -Scopes "Directory.Read.All", "Reports.Read.All"
        $currentUser = Get-MgContext
        if (-not $currentUser) {
            throw "Failed to authenticate with Microsoft Graph"
        }

        $roles = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" }
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $roles.Id
        $isGlobalAdmin = $members.AdditionalProperties.userPrincipalName -contains $currentUser.Account

        if (-not $isGlobalAdmin) {
            throw "This script requires Global Administrator privileges."
        }
        return $true
    }
    catch {
        Write-Log "Global Administrator check failed: $_"
        return $false
    }
}

function Convert-ToHTML {
    param(
        [Parameter(Mandatory = $true)]
        [Array]$Data,
        [string]$Title,
        [string]$Description,
        [string]$TableId = "dataTable"
    )
    
    $htmlHeader = @"
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <title>$Title</title>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="description" content="OneDrive for Business monitoring report">
        <meta name="format-detection" content="telephone=no">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <style>
            body { 
                font-family: Arial, sans-serif; 
                margin: 20px;
                background-color: #f5f6fa;
            }
            .container {
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            h1 { 
                color: #2c3e50;
                display: flex;
                align-items: center;
                gap: 10px;
            }
            h1 i {
                font-size: 0.8em;
            }
            table.data-table { 
                border-collapse: collapse; 
                width: 100%; 
                margin-top: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            th { 
                background-color: #3498db; 
                color: white;
                cursor: pointer;
                user-select: none;
                position: relative;
                padding-right: 20px;
            }
            th:hover {
                background-color: #2980b9;
            }
            th::after {
                content: '↕';
                position: absolute;
                right: 8px;
                opacity: 0.5;
            }
            th.sort-asc::after {
                content: '↑';
                opacity: 1;
            }
            th.sort-desc::after {
                content: '↓';
                opacity: 1;
            }
            td, th { 
                border: 1px solid #ddd; 
                padding: 12px 8px; 
                text-align: left; 
            }
            tr:nth-child(even) { background-color: #f8f9fa; }
            tr:hover { background-color: #f1f2f6; }
            .description { 
                color: #7f8c8d; 
                margin-bottom: 20px;
                padding: 10px;
                background-color: #f8f9fa;
                border-left: 4px solid #3498db;
            }
            .timestamp { 
                color: #95a5a6; 
                margin-top: 20px; 
                font-size: 0.9em;
                text-align: right;
            }
            .error { color: #e74c3c; }
            .warning { color: #f39c12; }
            .success { color: #2ecc71; }
            .controls {
                margin: 20px 0;
                display: flex;
                gap: 10px;
                align-items: center;
            }
            .csv-download {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            .csv-download:hover {
                background-color: #219a52;
            }
            .search-container {
                margin: 20px 0;
                display: flex;
                gap: 10px;
                align-items: center;
            }
            .table-search {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                width: 300px;
                font-size: 14px;
            }
            .print-report {
                background-color: #34495e;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            .print-report:hover {
                background-color: #2c3e50;
            }
            .charts-container {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
                gap: 20px;
                margin: 20px 0;
            }
            .chart-wrapper {
                background: white;
                padding: 15px;
                border-radius: 8px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            .chart-title {
                font-size: 16px;
                color: #2c3e50;
                margin-bottom: 10px;
                text-align: center;
            }
            @media (max-width: 768px) {
                .container { padding: 10px; }
                table { display: block; overflow-x: auto; }
                td, th { white-space: nowrap; }
                .controls { flex-direction: column; }
            }
            .auto-refresh {
                display: flex;
                align-items: center;
                gap: 10px;
                margin: 10px 0;
            }
            .auto-refresh select {
                padding: 5px;
                border-radius: 4px;
                border: 1px solid #ddd;
            }
            .export-options {
                display: flex;
                gap: 10px;
            }
            .pdf-export {
                background-color: #c0392b;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            .pdf-export:hover {
                background-color: #a93226;
            }
            @media screen and (min-width: 2048px) {
                .container {
                    max-width: 1600px;
                }
                .charts-container {
                    grid-template-columns: repeat(3, 1fr);
                }
            }
            @media screen and (max-width: 768px) {
                .charts-container {
                    grid-template-columns: 1fr;
                }
                .export-options {
                    flex-direction: column;
                }
            }
            .loading {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(255, 255, 255, 0.8);
                display: none;
                justify-content: center;
                align-items: center;
                z-index: 1000;
            }
            .loading.active {
                display: flex;
            }
            .loading-spinner {
                width: 50px;
                height: 50px;
                border: 5px solid #f3f3f3;
                border-top: 5px solid #3498db;
                border-radius: 50%;
                animation: spin 1s linear infinite;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <div class="loading">
            <div class="loading-spinner"></div>
        </div>
        <div class="container">
            <h1><i class="fas fa-cloud"></i>$Title</h1>
            <div class="description"><i class="fas fa-info-circle"></i> $Description</div>
            
            <div class="auto-refresh">
                <i class="fas fa-sync-alt"></i>
                <label>自動更新間隔：</label>
                <select id="refreshInterval">
                    <option value="0">無効</option>
                    <option value="60">1分</option>
                    <option value="300">5分</option>
                    <option value="600">10分</option>
                    <option value="1800">30分</option>
                    <option value="3600">1時間</option>
                </select>
            </div>

            <div class="charts-container">
                <div class="chart-wrapper">
                    <div class="chart-title">使用率グラフ</div>
                    <canvas id="usageChart"></canvas>
                </div>
                <div class="chart-wrapper">
                    <div class="chart-title">エラー統計</div>
                    <canvas id="errorChart"></canvas>
                </div>
            </div>

            <div class="controls">
                <div class="search-container">
                    <i class="fas fa-search"></i>
                    <input type="text" class="table-search" data-table="$TableId" placeholder="テーブルを検索...">
                </div>
                <div class="export-options">
                    <button class="csv-download" data-table="$TableId">
                        <i class="fas fa-download"></i> CSVダウンロード
                    </button>
                    <button class="pdf-export" data-table="$TableId">
                        <i class="fas fa-file-pdf"></i> PDFダウンロード
                    </button>
                    <button class="print-report">
                        <i class="fas fa-print"></i> 印刷
                    </button>
                </div>
            </div>
"@

    $htmlFooter = @"
            <div class="timestamp">
                <i class="fas fa-clock"></i> レポート生成日時: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            </div>
        </div>
        <script src="report-utils.js"></script>
    </body>
    </html>
"@

    # テーブルにIDを追加し、データ行にステータスに応じたクラスを適用
    $tableHtml = "<table id='$TableId' class='data-table'>"
    $headers = $Data[0].PSObject.Properties.Name
    $tableHtml += "<thead><tr>"
    foreach ($header in $headers) {
        $tableHtml += "<th>$header</th>"
    }
    $tableHtml += "</tr></thead><tbody>"

    foreach ($row in $Data) {
        $rowHtml = "<tr>"
        foreach ($header in $headers) {
            $value = $row.$header
            $class = ""
            
            # ステータスに基づいて色分けクラスを適用
            if ($value -match "error|failed|disabled|false" -or $value -eq $false) {
                $class = "error"
            }
            elseif ($value -match "warning|pending|unknown" -or ($header -eq "UsagePercentage" -and [double]$value -gt 90)) {
                $class = "warning"
            }
            elseif ($value -match "success|enabled|true|completed" -or $value -eq $true) {
                $class = "success"
            }
            
            $rowHtml += "<td$(if($class){" class='$class'"})>$value</td>"
        }
        $rowHtml += "</tr>"
        $tableHtml += $rowHtml
    }
    $tableHtml += "</tbody></table>"

    return $htmlHeader + $tableHtml + $htmlFooter
}

function Get-OneDriveSyncErrors {
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'Microsoft-Windows-OneDrive'
            Level = @(2,3) # Error and Warning levels
            StartTime = (Get-Date).AddDays(-1)
        } -ErrorAction SilentlyContinue

        if ($events) {
            $events = $events | Select-Object TimeCreated, Id, LevelDisplayName, Message
            
            # CSV出力
            $events | Export-Csv -Path "$OutputPath\OneDriveSyncErrors_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation -Encoding UTF8
            
            # HTML出力
            $htmlContent = Convert-ToHTML -Data $events `
                -Title "OneDrive同期エラーレポート" `
                -Description "過去24時間のOneDrive同期エラーと警告"
            $htmlContent | Out-File "$OutputPath\OneDriveSyncErrors_$(Get-Date -Format 'yyyyMMdd').html" -Encoding UTF8
        }
        Write-Log "Sync errors collected successfully"
    }
    catch {
        Write-Log "Error collecting sync errors: $_"
    }
}

function Get-OneDriveClientVersion {
    try {
        $oneDrivePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
        if (Test-Path $oneDrivePath) {
            $version = (Get-Item $oneDrivePath).VersionInfo.FileVersion
            return $version
        }
    }
    catch {
        Write-Log "Error getting OneDrive client version: $_"
        return "Unknown"
    }
}

function Get-OneDriveUsageReport {
    try {
        # Connect to Microsoft Online Service
        if (-not (Get-Module -ListAvailable -Name MSOnline)) {
            Write-Log "MSOnline module not found. Installing..."
            Install-Module -Name MSOnline -Force -AllowClobber
        }
        
        # Get OneDrive usage data
        $users = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true }
        $report = @()
        
        foreach ($user in $users) {
            $userData = [PSCustomObject]@{
                Name = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Email = $user.UserPrincipalName
                Status = $user.BlockCredential ? "Disabled" : "Enabled"
                AllocatedStorageGB = 1024 # Default OneDrive storage, adjust as needed
                UsedStorageGB = $null
                RemainingStorageGB = $null
                UsagePercentage = $null
                LastModified = $null
                OneDriveVersion = Get-OneDriveClientVersion
            }
            
            $report += $userData
        }
        
        # CSV出力
        $report | Export-Csv -Path "$OutputPath\OneDriveUsage_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation -Encoding UTF8
        
        # HTML出力
        $htmlContent = Convert-ToHTML -Data $report `
            -Title "OneDrive使用状況レポート" `
            -Description "ユーザーごとのOneDrive使用状況詳細"
        $htmlContent | Out-File "$OutputPath\OneDriveUsage_$(Get-Date -Format 'yyyyMMdd').html" -Encoding UTF8
        
        Write-Log "Usage report generated successfully"
    }
    catch {
        Write-Log "Error generating usage report: $_"
    }
}

function Get-OneDrivePolicyCompliance {
    try {
        $requiredSettings = @{
            "KFMSilentOptIn" = "Should exist"
            "KFMSilentOptInWithNotification" = 1
            "FilesOnDemandEnabled" = 1
            "DisableExternalSharing" = 1
        }
        
        $compliance = @()
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
        
        foreach ($setting in $requiredSettings.GetEnumerator()) {
            $actual = Get-ItemProperty -Path $regPath -Name $setting.Key -ErrorAction SilentlyContinue
            $compliant = if ($setting.Value -eq "Should exist") {
                $null -ne $actual
            } else {
                $actual.$($setting.Key) -eq $setting.Value
            }
            
            $compliance += [PSCustomObject]@{
                Setting = $setting.Key
                ExpectedValue = $setting.Value
                ActualValue = $actual.$($setting.Key)
                Compliant = $compliant
            }
        }
        
        # CSV出力
        $compliance | Export-Csv -Path "$OutputPath\OneDriveCompliance_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation -Encoding UTF8
        
        # HTML出力
        $htmlContent = Convert-ToHTML -Data $compliance `
            -Title "OneDriveポリシー準拠状況レポート" `
            -Description "OneDriveポリシー設定の準拠状況確認"
        $htmlContent | Out-File "$OutputPath\OneDriveCompliance_$(Get-Date -Format 'yyyyMMdd').html" -Encoding UTF8
        
        Write-Log "Policy compliance report generated successfully"
    }
    catch {
        Write-Log "Error checking policy compliance: $_"
    }
}

# Main execution
Write-Log "Starting OneDrive monitoring..."

Write-Log "Checking and installing required modules..."
Install-RequiredModules

Write-Log "Verifying Global Administrator privileges..."
if (-not (Test-GlobalAdminRole)) {
    Write-Log "Error: This script must be run as a Global Administrator."
    exit 1
}

Get-OneDriveSyncErrors
Get-OneDriveUsageReport
Get-OneDrivePolicyCompliance
Write-Log "OneDrive monitoring completed"