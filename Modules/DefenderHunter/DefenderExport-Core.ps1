# IoC Analyzer and Detection Export Tool
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
# Description: A PowerShell script to analyze IoCs from a CSV file and generate hunting queries for Microsoft Defender for Endpoint. 
# It also provides options to configure automated exports and run manual exports of detection results.


# Run-DefenderExport.ps1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\export-config.json",
    
    [Parameter(Mandatory=$false)]
    [string]$WorkspaceId,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".\DefenderExports"
)

# Import required modules
Import-Module Az.OperationalInsights -ErrorAction Stop

# Initialize logging
$LogPath = ".\DefenderExport.log"
$Script:ErrorCount = 0
$Script:WarningCount = 0

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with appropriate color
    switch ($Level) {
        'Warning' { 
            Write-Host $logMessage -ForegroundColor Yellow
            $Script:WarningCount++
        }
        'Error' { 
            Write-Host $logMessage -ForegroundColor Red
            $Script:ErrorCount++
        }
        default { Write-Host $logMessage }
    }
    
    # Write to log file
    Add-Content -Path $LogPath -Value $logMessage
}

function Initialize-ExportEnvironment {
    param(
        [string]$ExportPath
    )
    
    try {
        # Create export directory if it doesn't exist
        if (-not (Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
            Write-Log "Created export directory: $ExportPath"
        }
        
        # Create log file if it doesn't exist
        if (-not (Test-Path $LogPath)) {
            New-Item -ItemType File -Path $LogPath -Force | Out-Null
            Write-Log "Created log file: $LogPath"
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to initialize export environment: $_" -Level Error
        return $false
    }
}

function Test-Configuration {
    param(
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log "Configuration file not found at: $ConfigPath" -Level Error
            return $null
        }
        
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        
        # Validate required properties
        $requiredProps = @('WorkspaceId', 'ExportPath')
        $missingProps = $requiredProps | Where-Object { -not $config.$_ }
        
        if ($missingProps) {
            Write-Log "Missing required configuration properties: $($missingProps -join ', ')" -Level Error
            return $null
        }
        
        return $config
    }
    catch {
        Write-Log "Failed to load configuration: $_" -Level Error
        return $null
    }
}