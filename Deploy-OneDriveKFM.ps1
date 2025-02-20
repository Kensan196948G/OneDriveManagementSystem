#Requires -RunAsAdministrator
#Requires -Version 5.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

# OneDrive KFM deployment script
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantID,
    
    [Parameter()]
    [switch]$WhatIf,

    [Parameter()]
    [switch]$EnableKFM,

    [Parameter()]
    [switch]$EnableNotification
)

$script:ShowNotification = $true  # デフォルトで通知を有効化

$ScriptVersion = "1.1.0"

# Registry paths
$OneDriveRegKey = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
$TenantRegKey = "$OneDriveRegKey\AllowTenants"

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

function Test-AdminPrivileges {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Initialize-OneDriveRegistryKeys {
    try {
        if (-not (Test-Path $OneDriveRegKey)) {
            New-Item -Path $OneDriveRegKey -Force | Out-Null
        }
        if (-not (Test-Path $TenantRegKey)) {
            New-Item -Path $TenantRegKey -Force | Out-Null
        }

        # ベース設定
        Set-ItemProperty -Path $OneDriveRegKey -Name "FilesOnDemandEnabled" -Value 1 -Type DWord
        Set-ItemProperty -Path $OneDriveRegKey -Name "DisablePersonalSync" -Value 1 -Type DWord
        Set-ItemProperty -Path $OneDriveRegKey -Name "SilentAccountConfig" -Value 1 -Type DWord

        # KFM設定（オプション）
        if ($EnableKFM) {
            Write-Host "Configuring KFM settings with user notification..." -ForegroundColor Green
            Set-ItemProperty -Path $OneDriveRegKey -Name "KFMOptIn" -Value $TenantID -Type String
            if ($EnableNotification -or $script:ShowNotification) {
                Set-ItemProperty -Path $OneDriveRegKey -Name "KFMOptInWithNotification" -Value 1 -Type DWord
            }
        }

        # Add tenant to allowed list
        Set-ItemProperty -Path $TenantRegKey -Name $TenantID -Value $TenantID -Type String

        Write-Host "OneDrive registry keys configured successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to configure registry keys: $_"
        throw
    }
}

function Restart-OneDriveProcess {
    try {
        Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 5
        Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
        Write-Host "OneDrive process restarted successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to restart OneDrive process: $_"
    }
}

function Update-OneDriveClient {
    try {
        # OneDrive update URI
        $uri = "https://go.microsoft.com/fwlink/?linkid=844652"
        $setupPath = "$env:TEMP\OneDriveSetup.exe"
        
        # Download and install latest OneDrive
        Invoke-WebRequest -Uri $uri -OutFile $setupPath
        Start-Process -FilePath $setupPath -ArgumentList "/silent" -Wait
        Remove-Item -Path $setupPath -Force
        
        Write-Host "OneDrive client updated successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to update OneDrive client: $_"
    }
}

# Main execution
try {
    Write-Host "Checking and installing required modules..." -ForegroundColor Cyan
    Install-RequiredModules

    Write-Host "Verifying Global Administrator privileges..." -ForegroundColor Cyan
    if (-not (Test-GlobalAdminRole)) {
        throw "This script must be run as a Global Administrator."
    }

    if (-not (Test-AdminPrivileges)) {
        throw "This script requires local administrator privileges."
    }

    Write-Host "Starting OneDrive configuration deployment (Version: $ScriptVersion)" -ForegroundColor Cyan
    
    if ($WhatIf) {
        Write-Host "`nWhatIf: Would perform the following actions:" -ForegroundColor Yellow
        Write-Host "- Update OneDrive client"
        Write-Host "- Configure base OneDrive settings"
        if ($EnableKFM) {
            Write-Host "- Enable KFM with notification: $EnableNotification"
        } else {
            Write-Host "- KFM will not be configured (use -EnableKFM to enable)"
        }
        exit 0
    }

    # Update OneDrive client first
    Update-OneDriveClient
    
    # Configure registry settings
    Initialize-OneDriveRegistryKeys
    
    # Restart OneDrive to apply changes
    Restart-OneDriveProcess
    
    Write-Host "OneDrive configuration deployment completed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Deployment failed: $_"
    exit 1
}