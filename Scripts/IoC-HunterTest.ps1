# IoC Analyzer and Detection Export Tool
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
# Description: A PowerShell script to analyze IoCs from a CSV file and generate hunting queries for Microsoft Defender for Endpoint. 
# It also provides options to configure automated exports and run manual exports of detection results.

# Create Main.ps1
$mainScript = @"
# Import component scripts
. '.\Scripts\DefenderExport-Core.ps1'
. '.\Scripts\DefenderExport-Queries.ps1'
. '.\Scripts\DefenderExport-Functions.ps1'
. '.\Scripts\DefenderExport-Main.ps1'

# Execution parameters
param(
    [Parameter(Mandatory=`$false)]
    [string]`$ConfigPath = ".\export-config.json",
    
    [Parameter(Mandatory=`$false)]
    [string]`$IocsFile,
    
    [Parameter(Mandatory=`$false)]
    [switch]`$ExportOnly
)

# Main execution
try {
    # Initialize
    Write-Host "Initializing..."
    Initialize-ExportEnvironment

    # Process IoCs if provided
    if (`$IocsFile) {
        Write-Host "Processing IoCs from: `$IocsFile"
        Analyze-IoCs -CsvPath `$IocsFile
    }

    # Run export
    if (-not `$ExportOnly) {
        Write-Host "Running detection export..."
        Start-DefenderExport -ConfigPath `$ConfigPath
    }
}
catch {
    Write-Error "Execution failed: `$_"
    exit 1
}
"@

$mainScript | Out-File "C:\DefenderHunting\Main.ps1"