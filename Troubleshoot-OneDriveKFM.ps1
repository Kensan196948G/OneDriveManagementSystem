#Requires -RunAsAdministrator
#Requires -Version 5.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

[CmdletBinding()]
param()

# Install required modules if not present
function Install-RequiredModules {
    $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users')
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing $module module..." -ForegroundColor Yellow
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -Force
    }
}

# Check if the current user is a Global Administrator
function Test-GlobalAdminRole {
    try {
        Connect-MgGraph -Scopes "Directory.Read.All"
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
        Write-Error "Global Administrator check failed: $_"
        return $false
    }
}

function Test-OneDriveConnection {
    $odProcess = Get-Process OneDrive -ErrorAction SilentlyContinue
    if (-not $odProcess) {
        Write-Warning "OneDrive process is not running!"
        return $false
    }
    
    $odConnected = Test-Path "$env:OneDriveCommercial"
    if (-not $odConnected) {
        Write-Warning "OneDrive Business is not connected!"
        return $false
    }
    
    Write-Host "OneDrive connection status: OK" -ForegroundColor Green
    return $true
}

function Test-KFMConfiguration {
    $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
    $requiredKeys = @{
        "KFMSilentOptIn" = "String"
        "KFMSilentOptInWithNotification" = "DWord"
        "FilesOnDemandEnabled" = "DWord"
        "DisableExternalSharing" = "DWord"
    }
    
    $issues = @()
    
    if (-not (Test-Path $regPath)) {
        Write-Warning "OneDrive policy registry key not found!"
        return $false
    }
    
    foreach ($key in $requiredKeys.GetEnumerator()) {
        $value = Get-ItemProperty -Path $regPath -Name $key.Key -ErrorAction SilentlyContinue
        if (-not $value) {
            $issues += "Missing registry key: $($key.Key)"
        }
    }
    
    if ($issues.Count -gt 0) {
        Write-Warning "KFM configuration issues found:"
        $issues | ForEach-Object { Write-Warning "- $_" }
        return $false
    }
    
    Write-Host "KFM configuration status: OK" -ForegroundColor Green
    return $true
}

function Repair-OneDriveSetup {
    try {
        # Reset OneDrive
        Write-Host "Attempting to reset OneDrive..."
        Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Clear OneDrive cache
        Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive\settings\*" -Recurse -Force -ErrorAction SilentlyContinue
        
        # Reinstall OneDrive if needed
        $oneDrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
        if (Test-Path $oneDrivePath) {
            Write-Host "Reinstalling OneDrive..."
            Start-Process $oneDrivePath -ArgumentList "/reset" -Wait
            Start-Sleep -Seconds 5
            Start-Process $oneDrivePath
        }
        
        Write-Host "OneDrive reset completed. Please sign in again." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to repair OneDrive: $_"
    }
}

function Get-KFMMigrationStatus {
    $knownFolders = @{
        Desktop = [Environment]::GetFolderPath("Desktop")
        Documents = [Environment]::GetFolderPath("MyDocuments")
        Pictures = [Environment]::GetFolderPath("MyPictures")
    }
    
    foreach ($folder in $knownFolders.GetEnumerator()) {
        $path = $folder.Value
        $isOnOneDrive = $path -like "*OneDrive*"
        
        Write-Host "$($folder.Key) folder status:" -NoNewline
        if ($isOnOneDrive) {
            Write-Host " Migrated to OneDrive" -ForegroundColor Green
        } else {
            Write-Host " Not migrated" -ForegroundColor Yellow
        }
    }
}

function Convert-ToHTML {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,
        [string]$Title = "OneDrive診断レポート"
    )
    
    $htmlContent = @"
    <!DOCTYPE html>
    <html>
    <head>
        <title>$Title</title>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>
            body { 
                font-family: Arial, sans-serif; 
                margin: 20px; 
                background-color: #f5f6fa; 
            }
            .container { 
                max-width: 1200px;
                margin: 0 auto;
                background-color: white; 
                padding: 20px; 
                border-radius: 8px; 
                box-shadow: 0 2px 4px rgba(0,0,0,0.1); 
            }
            h1 { 
                color: #2c3e50; 
                border-bottom: 2px solid #3498db; 
                padding-bottom: 10px;
                display: flex;
                align-items: center;
                gap: 10px;
            }
            h2 { 
                color: #34495e; 
                margin-top: 20px;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            .section { 
                margin-bottom: 30px;
                background-color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            .section:hover {
                box-shadow: 0 2px 6px rgba(0,0,0,0.15);
            }
            pre { 
                background-color: #f8f9fa; 
                padding: 15px; 
                border-radius: 5px; 
                overflow-x: auto;
                border-left: 4px solid #3498db;
            }
            .error { 
                color: #e74c3c;
                border-left-color: #e74c3c;
            }
            .warning { 
                color: #f39c12;
                border-left-color: #f39c12;
            }
            .success { 
                color: #27ae60;
                border-left-color: #27ae60;
            }
            .timestamp { 
                color: #95a5a6; 
                margin-top: 20px; 
                font-size: 0.9em;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            table.data-table { 
                border-collapse: collapse; 
                width: 100%; 
                margin-top: 10px;
            }
            th { 
                background-color: #3498db; 
                color: white;
                cursor: pointer;
            }
            td, th { 
                border: 1px solid #ddd; 
                padding: 8px; 
                text-align: left; 
            }
            tr:nth-child(even) { background-color: #f8f9fa; }
            tr:hover { background-color: #f1f2f6; }
            .status-icon {
                margin-right: 8px;
            }
            .controls {
                margin: 20px 0;
                display: flex;
                gap: 10px;
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
            @media (max-width: 768px) {
                .container { 
                    padding: 10px;
                    margin: 10px;
                }
                pre { 
                    font-size: 14px;
                    padding: 10px;
                }
                .section {
                    padding: 15px;
                }
                table { 
                    display: block;
                    overflow-x: auto;
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1><i class="fas fa-diagnoses"></i>$Title</h1>
            <div class="content">
                $Content
            </div>
            <div class="timestamp">
                <i class="fas fa-clock"></i>
                レポート生成日時: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            </div>
        </div>
        <script src="report-utils.js"></script>
    </body>
    </html>
"@

    return $htmlContent
}

function Export-DiagnosticInfo {
    $diagFile = "OneDriveDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $htmlContent = ""
    
    # OneDriveバージョン情報
    $htmlContent += "<div class='section' id='version-info'>"
    $htmlContent += "<h2><i class='fas fa-info-circle'></i>OneDriveバージョン情報</h2>"
    try {
        $versionInfo = Get-ItemProperty "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" | Select-Object VersionInfo
        $htmlContent += "<pre class='success'>$($versionInfo | ConvertTo-Html -Fragment)</pre>"
    }
    catch {
        $htmlContent += "<pre class='error'>バージョン情報の取得に失敗しました: $_</pre>"
    }
    $htmlContent += "</div>"
    
    # OneDriveプロセス状態
    $htmlContent += "<div class='section' id='process-status'>"
    $htmlContent += "<h2><i class='fas fa-tasks'></i>OneDriveプロセス状態</h2>"
    $process = Get-Process OneDrive -ErrorAction SilentlyContinue
    if ($process) {
        $processInfo = $process | Select-Object Id, CPU, WorkingSet, StartTime
        $htmlContent += "<pre class='success'>$($processInfo | ConvertTo-Html -Fragment)</pre>"
        $htmlContent += "<div class='controls'>"
        $htmlContent += "<button class='csv-download' data-table='process-status'><i class='fas fa-download'></i> CSVダウンロード</button>"
        $htmlContent += "</div>"
    }
    else {
        $htmlContent += "<pre class='warning'>OneDriveプロセスが実行されていません。</pre>"
    }
    $htmlContent += "</div>"
    
    # レジストリ設定
    $htmlContent += "<div class='section' id='registry-settings'>"
    $htmlContent += "<h2><i class='fas fa-cogs'></i>レジストリ設定</h2>"
    try {
        $regSettings = Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" | 
            Select-Object KFMSilentOptIn, KFMSilentOptInWithNotification, FilesOnDemandEnabled, DisableExternalSharing
        $htmlContent += "<pre class='success'>$($regSettings | ConvertTo-Html -Fragment)</pre>"
        $htmlContent += "<div class='controls'>"
        $htmlContent += "<button class='csv-download' data-table='registry-settings'><i class='fas fa-download'></i> CSVダウンロード</button>"
        $htmlContent += "</div>"
    }
    catch {
        $htmlContent += "<pre class='error'>レジストリ設定の取得に失敗しました: $_</pre>"
    }
    $htmlContent += "</div>"
    
    # イベントログ
    $htmlContent += "<div class='section' id='event-logs'>"
    $htmlContent += "<h2><i class='fas fa-clipboard-list'></i>最近のOneDriveイベントログ</h2>"
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName='Microsoft-Windows-OneDrive'
            StartTime=(Get-Date).AddDays(-1)
        } -ErrorAction SilentlyContinue |
            Select-Object TimeCreated, Id, LevelDisplayName, Message
        
        if ($events) {
            $eventsHtml = $events | ConvertTo-Html -Fragment
            $eventsHtml = $eventsHtml -replace '<td>Error</td>', '<td class="error">Error</td>' `
                                     -replace '<td>Warning</td>', '<td class="warning">Warning</td>' `
                                     -replace '<td>Information</td>', '<td class="success">Information</td>'
            $htmlContent += "<pre>$eventsHtml</pre>"
            $htmlContent += "<div class='controls'>"
            $htmlContent += "<button class='csv-download' data-table='event-logs'><i class='fas fa-download'></i> CSVダウンロード</button>"
            $htmlContent += "</div>"
        }
        else {
            $htmlContent += "<pre class='warning'>イベントログが存在しません。</pre>"
        }
    }
    catch {
        $htmlContent += "<pre class='warning'>イベントログの取得に失敗したか、ログが存在しません。</pre>"
    }
    $htmlContent += "</div>"
    
    # 既知のフォルダー状態
    $htmlContent += "<div class='section' id='known-folders'>"
    $htmlContent += "<h2><i class='fas fa-folder'></i>既知のフォルダーの状態</h2><pre>"
    $knownFolders = @{
        Desktop = [Environment]::GetFolderPath("Desktop")
        Documents = [Environment]::GetFolderPath("MyDocuments")
        Pictures = [Environment]::GetFolderPath("MyPictures")
    }
    foreach ($folder in $knownFolders.GetEnumerator()) {
        $isOnOneDrive = $folder.Value -like "*OneDrive*"
        $icon = if ($isOnOneDrive) {
            "<i class='fas fa-check-circle success'></i>"
        } else {
            "<i class='fas fa-exclamation-triangle warning'></i>"
        }
        $status = if ($isOnOneDrive) {
            "<span class='success'>OneDriveに移行済み</span>"
        } else {
            "<span class='warning'>未移行</span>"
        }
        $htmlContent += "$icon $($folder.Key): $status - パス: $($folder.Value)<br>"
    }
    $htmlContent += "</pre></div>"
    
    # HTML形式で保存
    $finalHtml = Convert-ToHTML -Content $htmlContent -Title "OneDrive診断レポート"
    $finalHtml | Out-File "$diagFile.html" -Encoding UTF8
    
    # テキスト形式も維持（従来の出力）
    "=== OneDrive Diagnostic Information ===" | Out-File "$diagFile.txt"
    "Generated: $(Get-Date)" | Out-File "$diagFile.txt" -Append
    
    "--- OneDrive Version ---" | Out-File "$diagFile.txt" -Append
    $versionInfo | Out-File "$diagFile.txt" -Append
    
    "--- OneDrive Process Status ---" | Out-File "$diagFile.txt" -Append
    $process | Out-File "$diagFile.txt" -Append
    
    "--- Registry Configuration ---" | Out-File "$diagFile.txt" -Append
    $regSettings | Out-File "$diagFile.txt" -Append
    
    Write-Host "Diagnostic information exported to: $diagFile.html and $diagFile.txt" -ForegroundColor Green
}

# Main execution
Write-Host "OneDrive KFM Troubleshooter" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan

Write-Host "`nChecking and installing required modules..." -ForegroundColor Cyan
Install-RequiredModules

Write-Host "`nVerifying Global Administrator privileges..." -ForegroundColor Cyan
if (-not (Test-GlobalAdminRole)) {
    Write-Error "This script must be run as a Global Administrator."
    exit 1
}

Write-Host "`nChecking OneDrive connection..."
$connected = Test-OneDriveConnection

Write-Host "`nChecking KFM configuration..."
$configured = Test-KFMConfiguration

Write-Host "`nChecking Known Folder Migration status..."
Get-KFMMigrationStatus

Write-Host "`nExporting diagnostic information..."
Export-DiagnosticInfo

if (-not ($connected -and $configured)) {
    Write-Host "`nWould you like to attempt automatic repair? (Y/N)" -ForegroundColor Yellow
    $response = Read-Host
    if ($response -eq 'Y') {
        Repair-OneDriveSetup
    }
}