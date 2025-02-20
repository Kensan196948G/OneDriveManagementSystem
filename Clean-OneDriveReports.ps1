#Requires -RunAsAdministrator
#Requires -Version 5.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$RetentionDays = 30,
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "$PSScriptRoot\Reports",
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$PSScriptRoot\Logs"
)

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $logFile = Join-Path $LogPath "CleanupLog_$(Get-Date -Format 'yyyyMMdd').log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    
    switch ($Level) {
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage }
    }
}

function Remove-OldReports {
    param(
        [string]$Path,
        [int]$Days
    )
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$Days)
        
        # 各種レポートファイルの削除
        $extensions = @('*.csv', '*.html', '*.txt', '*.pdf')
        foreach ($ext in $extensions) {
            $files = Get-ChildItem -Path $Path -Filter $ext -Recurse |
                Where-Object { $_.LastWriteTime -lt $cutoffDate }
            
            foreach ($file in $files) {
                Remove-Item $file.FullName -Force
                Write-Log "Removed old file: $($file.FullName)"
            }
        }
        
        # 空のフォルダの削除
        Get-ChildItem -Path $Path -Directory -Recurse |
            Where-Object { (Get-ChildItem $_.FullName -Force).Count -eq 0 } |
            ForEach-Object {
                Remove-Item $_.FullName -Force
                Write-Log "Removed empty directory: $($_.FullName)"
            }
    }
    catch {
        Write-Log "Error removing old reports: $_" -Level Error
    }
}

function Remove-OldLogs {
    param(
        [string]$Path,
        [int]$Days
    )
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$Days)
        
        # ログファイルの削除
        $logFiles = Get-ChildItem -Path $Path -Filter "*.log" -Recurse |
            Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        foreach ($file in $logFiles) {
            Remove-Item $file.FullName -Force
            Write-Log "Removed old log: $($file.FullName)"
        }
        
        # 監査ログの処理
        $auditPath = Join-Path $Path "audit"
        if (Test-Path $auditPath) {
            $auditFiles = Get-ChildItem -Path $auditPath -Filter "OneDriveAudit_*.log" -Recurse |
                Where-Object { $_.LastWriteTime -lt $cutoffDate }
            
            foreach ($file in $auditFiles) {
                # 監査ログはアーカイブに移動
                $archivePath = Join-Path $auditPath "archive"
                if (-not (Test-Path $archivePath)) {
                    New-Item -ItemType Directory -Path $archivePath -Force | Out-Null
                }
                
                $archiveFile = Join-Path $archivePath $file.Name
                Move-Item $file.FullName $archiveFile -Force
                Write-Log "Archived audit log: $($file.Name)"
            }
        }
    }
    catch {
        Write-Log "Error removing old logs: $_" -Level Error
    }
}

function Compress-OldFiles {
    param(
        [string]$Path,
        [int]$Days
    )
    
    try {
        $cutoffDate = (Get-Date).AddDays(-$Days)
        $archivePath = Join-Path $Path "archive"
        
        if (-not (Test-Path $archivePath)) {
            New-Item -ItemType Directory -Path $archivePath -Force | Out-Null
        }
        
        # 圧縮対象のファイルを収集
        $filesToArchive = Get-ChildItem -Path $archivePath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($filesToArchive.Count -gt 0) {
            $archiveName = "Archive_$(Get-Date -Format 'yyyyMMdd').zip"
            $archiveFile = Join-Path $archivePath $archiveName
            
            Compress-Archive -Path $filesToArchive.FullName -DestinationPath $archiveFile -Force
            
            # 圧縮後の元ファイルを削除
            $filesToArchive | ForEach-Object {
                Remove-Item $_.FullName -Force
                Write-Log "Archived and removed: $($_.Name)"
            }
        }
    }
    catch {
        Write-Log "Error compressing old files: $_" -Level Error
    }
}

# メイン処理
try {
    Write-Log "Starting cleanup process..."
    
    # レポートの削除
    Write-Log "Cleaning up old reports..."
    Remove-OldReports -Path $ReportPath -Days $RetentionDays
    
    # ログの削除
    Write-Log "Cleaning up old logs..."
    Remove-OldLogs -Path $LogPath -Days $RetentionDays
    
    # 古いファイルの圧縮
    Write-Log "Compressing old files..."
    Compress-OldFiles -Path $LogPath -Days ($RetentionDays * 2)
    
    Write-Log "Cleanup completed successfully."
}
catch {
    Write-Log "Cleanup process failed: $_" -Level Error
    exit 1
}