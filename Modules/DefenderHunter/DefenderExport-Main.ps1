# IoC Analyzer and Detection Export Tool
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
# Description: A PowerShell script to analyze IoCs from a CSV file and generate hunting queries for Microsoft Defender for Endpoint. 
# It also provides options to configure automated exports and run manual exports of detection results.

# Main execution block
try {
    Write-Log "Starting Defender detection export"
    
    # Initialize environment
    if (-not (Initialize-ExportEnvironment -ExportPath $ExportPath)) {
        throw "Failed to initialize export environment"
    }
    
    # Load configuration
    $config = Test-Configuration -ConfigPath $ConfigPath
    if (-not $config) {
        throw "Failed to load configuration"
    }
    
    # Override config with parameters if provided
    if ($WorkspaceId) { $config.WorkspaceId = $WorkspaceId }
    if ($ExportPath) { $config.ExportPath = $ExportPath }
    
    # Get queries
    $queries = Get-DefenderQueries
    
    # Track export results
    $exportResults = @{}
    $summary = @{
        ExportTime = Get-Date
        ExportPath = $config.ExportPath
        UniqueDevices = 0
        FirstDetection = $null
        LastDetection = $null
    }
    
    # Execute queries and export results
    foreach ($queryType in $queries.Keys) {
        Write-Log "Processing $queryType query"
        $results = Invoke-DefenderQuery -WorkspaceId $config.WorkspaceId -Query $queries[$queryType]
        
        if ($results) {
            $exportPath = Export-QueryResults -Results $results -ExportPath $config.ExportPath -QueryType $queryType
            if ($exportPath) {
                $exportResults[$queryType] = @{
                    Path = $exportPath
                    Count = $results.Count
                }
                $summary[$queryType] = $results.Count
            }
        }
    }
    
    # Update summary statistics
    if ($exportResults.Count -gt 0) {
        $allDevices = @()
        $allTimes = @()
        
        foreach ($export in $exportResults.Values) {
            $data = Import-Csv $export.Path
            $allDevices += $data.DeviceName
            $allTimes += $data.TimeGenerated
        }
        
        $summary.UniqueDevices = ($allDevices | Select-Object -Unique).Count
        $summary.FirstDetection = ($allTimes | Measure-Object -Minimum).Minimum
        $summary.LastDetection = ($allTimes | Measure-Object -Maximum).Maximum
    }
    
    # Generate summary report
    $summaryPath = Join-Path $config.ExportPath "Summary-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    $summary | ConvertTo-Json | Out-File $summaryPath
    Write-Log "Generated summary report: $summaryPath"
    
    # Send notification
    if ($config.EmailRecipient) {
        Send-ExportNotification -Summary $summary -EmailRecipient $config.EmailRecipient
    }
    
    # Cleanup old exports
    if ($config.RetentionDays) {
        Remove-OldExports -ExportPath $config.ExportPath -RetentionDays $config.RetentionDays
    }
    
    Write-Log "Export completed successfully"
}
catch {
    Write-Log "Export failed: $_" -Level Error
    exit 1
}