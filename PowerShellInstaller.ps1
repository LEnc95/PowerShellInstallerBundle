<#
.SYNOPSIS
    Installs and updates required PowerShell modules for Microsoft 365 management.
    
.DESCRIPTION
    This script installs or updates essential Microsoft PowerShell modules including:
    - Microsoft.Graph
    - ExchangeOnlineManagement
    - Microsoft.Online.SharePoint.PowerShell
    - AzureAD
    - MSOnline
    
    The script includes comprehensive logging, error handling, and version validation.
    
.PARAMETER UninstallExisting
    If specified, uninstalls existing versions of modules before installing fresh copies.
    
.PARAMETER SkipExecutionPolicy
    If specified, skips setting the execution policy (useful when run via LaunchInstaller.bat).
    
.EXAMPLE
    .\PowerShellInstaller.ps1
    Runs the installer with default settings, prompting for uninstall option.
    
.EXAMPLE
    .\PowerShellInstaller.ps1 -UninstallExisting
    Uninstalls existing modules before installing fresh versions.
    
.NOTES
    Requires PowerShell 5.1 or higher.
    Run as Administrator for system-wide installation, or CurrentUser scope will be used.
#>

[CmdletBinding()]
param(
    [switch]$UninstallExisting,
    [switch]$SkipExecutionPolicy
)

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "This script requires PowerShell 5.1 or higher. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}

# Define log file path with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = "$env:USERPROFILE\Documents\PowerShellModuleInstallLog_$timestamp.txt"

# Start transcript logging
try {
    Start-Transcript -Path $logPath -Append -ErrorAction Stop
    Write-Host "`n=== PowerShell Module Installer ===" -ForegroundColor Cyan
    Write-Host "Log file: $logPath`n" -ForegroundColor Gray
} catch {
    Write-Warning "Could not start transcript. Continuing without file logging..."
    $logPath = $null
}

# Define required modules and minimum versions
$requiredModules = @{
    "Microsoft.Graph"                        = "1.27.0"
    "ExchangeOnlineManagement"              = "3.4.0"
    "Microsoft.Online.SharePoint.PowerShell"= "16.0.24310.12000"
    "AzureAD"                               = "2.0.2.140"
    "MSOnline"                              = "1.1.183.57"
}

# Track installation results
$script:InstallResults = @{
    Success = @()
    Failed = @()
    Skipped = @()
    Updated = @()
}

function Uninstall-ModuleSafe {
    param (
        [string]$ModuleName
    )
    try {
        $installed = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
        if ($installed) {
            Write-Host "  Uninstalling $ModuleName (Version: $($installed.Version))..." -ForegroundColor Yellow
            Uninstall-Module -Name $ModuleName -AllVersions -Force -ErrorAction Stop
            Write-Host "  [OK] $ModuleName uninstalled successfully." -ForegroundColor Green
            return $true
        }
        return $false
    } catch {
        Write-Warning "  Error uninstalling $ModuleName`: $_"
        return $false
    }
}

function Test-ModuleInstalled {
    param (
        [string]$ModuleName,
        [string]$MinVersion
    )
    try {
        $installed = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
        if ($null -eq $installed) {
            return @{ Status = "NotInstalled"; Version = $null }
        }
        
        $installedVersion = [version]$installed.Version
        $requiredVersion = [version]$MinVersion
        
        if ($installedVersion -ge $requiredVersion) {
            return @{ Status = "UpToDate"; Version = $installed.Version }
        } else {
            return @{ Status = "OutOfDate"; Version = $installed.Version }
        }
    } catch {
        return @{ Status = "Error"; Version = $null; Error = $_ }
    }
}

function Install-Or-UpdateModule {
    param (
        [string]$ModuleName,
        [string]$MinVersion
    )

    Write-Host "`n[$ModuleName]" -ForegroundColor Cyan
    Write-Host "  Required version: $MinVersion" -ForegroundColor Gray

    try {
        # Check current installation status
        $status = Test-ModuleInstalled -ModuleName $ModuleName -MinVersion $MinVersion

        # Handle uninstall if requested
        if ($UninstallExisting) {
            $uninstalled = Uninstall-ModuleSafe -ModuleName $ModuleName
            if ($uninstalled) {
                $status.Status = "NotInstalled"
            }
        }

        # Install or update based on status
        switch ($status.Status) {
            "NotInstalled" {
                Write-Host "  Installing $ModuleName..." -ForegroundColor Yellow
                try {
                    Install-Module -Name $ModuleName -MinimumVersion $MinVersion -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                    $installed = Get-InstalledModule -Name $ModuleName -ErrorAction Stop
                    Write-Host "  [OK] $ModuleName installed successfully (Version: $($installed.Version))" -ForegroundColor Green
                    $script:InstallResults.Success += "$ModuleName ($($installed.Version))"
                } catch {
                    Write-Host "  [FAILED] Failed to install $ModuleName`: $_" -ForegroundColor Red
                    $script:InstallResults.Failed += "$ModuleName - $_"
                }
            }
            "OutOfDate" {
                Write-Host "  Installed version ($($status.Version)) is older than required. Updating..." -ForegroundColor Yellow
                try {
                    Update-Module -Name $ModuleName -Force -ErrorAction Stop
                    $updated = Get-InstalledModule -Name $ModuleName -ErrorAction Stop
                    Write-Host "  [OK] $ModuleName updated successfully (Version: $($updated.Version))" -ForegroundColor Green
                    $oldVersion = $status.Version
                    $newVersion = $updated.Version
                    $script:InstallResults.Updated += "$ModuleName ($oldVersion to $newVersion)"
                } catch {
                    Write-Host "  [FAILED] Failed to update $ModuleName`: $_" -ForegroundColor Red
                    $script:InstallResults.Failed += "$ModuleName - $_"
                }
            }
            "UpToDate" {
                Write-Host "  [OK] $ModuleName is up-to-date (Version: $($status.Version))" -ForegroundColor Green
                $script:InstallResults.Skipped += "$ModuleName ($($status.Version))"
            }
            default {
                Write-Host "  [WARNING] Error checking $ModuleName status: $($status.Error)" -ForegroundColor Yellow
                # Try to install anyway
                try {
                    Write-Host "  Attempting to install $ModuleName..." -ForegroundColor Yellow
                    Install-Module -Name $ModuleName -MinimumVersion $MinVersion -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                    $installed = Get-InstalledModule -Name $ModuleName -ErrorAction Stop
                    Write-Host "  [OK] $ModuleName installed successfully (Version: $($installed.Version))" -ForegroundColor Green
                    $script:InstallResults.Success += "$ModuleName ($($installed.Version))"
                } catch {
                    Write-Host "  [FAILED] Failed to install $ModuleName`: $_" -ForegroundColor Red
                    $script:InstallResults.Failed += "$ModuleName - $_"
                }
            }
        }
    } catch {
        Write-Host "  [FAILED] Unexpected error processing $ModuleName`: $_" -ForegroundColor Red
        $script:InstallResults.Failed += "$ModuleName - $_"
    }
}

try {
    # Set Execution Policy (unless skipped)
    if (-not $SkipExecutionPolicy) {
        Write-Host "`n[Prerequisites]" -ForegroundColor Cyan
        Write-Host "Setting execution policy to RemoteSigned..." -ForegroundColor Gray
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "  [OK] Execution policy set." -ForegroundColor Green
        } catch {
            Write-Warning "  Could not set execution policy: $_"
        }
    }

    # Ensure NuGet Provider is available
    Write-Host "`n[Prerequisites]" -ForegroundColor Cyan
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Host "  Installing NuGet provider..." -ForegroundColor Yellow
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion "2.8.5.201" -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "  [OK] NuGet provider installed." -ForegroundColor Green
        } catch {
            Write-Warning "  Could not install NuGet provider: $_"
        }
    } else {
        Write-Host "  [OK] NuGet provider already available." -ForegroundColor Green
    }

    # Ensure PowerShellGet is available and updated
    if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
        Write-Host "  Installing PowerShellGet..." -ForegroundColor Yellow
        try {
            Install-Module -Name PowerShellGet -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
            Write-Host "  [OK] PowerShellGet installed." -ForegroundColor Green
        } catch {
            Write-Warning "  Could not install PowerShellGet: $_"
        }
    } else {
        Write-Host "  [OK] PowerShellGet already available." -ForegroundColor Green
        # Try to update PowerShellGet to latest version
        try {
            Update-Module -Name PowerShellGet -Force -ErrorAction SilentlyContinue
        } catch {
            # Non-critical, continue
        }
    }

    # Install or update each required module
    Write-Host "`n[Module Installation]" -ForegroundColor Cyan
    $moduleCount = 0
    foreach ($module in $requiredModules.Keys) {
        $moduleCount++
        Write-Progress -Activity "Installing Modules" -Status "Processing $module ($moduleCount of $($requiredModules.Count))" -PercentComplete (($moduleCount / $requiredModules.Count) * 100)
        Install-Or-UpdateModule -ModuleName $module -MinVersion $requiredModules[$module]
    }
    Write-Progress -Activity "Installing Modules" -Completed

    # Display summary
    Write-Host "`n=== Installation Summary ===" -ForegroundColor Cyan
    if ($script:InstallResults.Success.Count -gt 0) {
        Write-Host "`n[SUCCESS] Successfully Installed:" -ForegroundColor Green
        $script:InstallResults.Success | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }
    if ($script:InstallResults.Updated.Count -gt 0) {
        Write-Host "`n[UPDATED] Updated:" -ForegroundColor Yellow
        $script:InstallResults.Updated | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }
    if ($script:InstallResults.Skipped.Count -gt 0) {
        Write-Host "`n[SKIPPED] Already Up-to-Date:" -ForegroundColor Cyan
        $script:InstallResults.Skipped | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }
    if ($script:InstallResults.Failed.Count -gt 0) {
        Write-Host "`n[FAILED] Failed:" -ForegroundColor Red
        $script:InstallResults.Failed | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }

    Write-Host "`n=== Setup Complete ===" -ForegroundColor Cyan
    Write-Host "You may need to restart PowerShell for changes to take effect." -ForegroundColor Yellow
    Write-Host "If you encounter issues, check the log file for details.`n" -ForegroundColor Gray

} catch {
    Write-Error "Fatal error during setup: $_"
    Write-Host "`nPlease check the log file for more details." -ForegroundColor Red
} finally {
    if ($null -ne $logPath) {
        try {
            Stop-Transcript
            Write-Host "Log saved to: $logPath" -ForegroundColor Gray
        } catch {
            # Transcript may not have been started
        }
    }
}