# IoC Analyzer and Detection Export Tool
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
# Description: A PowerShell script to analyze IoCs from a CSV file and generate hunting queries for Microsoft Defender for Endpoint. 
# It also provides options to configure automated exports and run manual exports of detection results.

function Export-QueryResults {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Results,
        
        [Parameter(Mandatory=$true)]
        [string]$ExportPath,
        
        [Parameter(Mandatory=$true)]
        [string]$QueryType
    )
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $fileName = "$QueryType-$timestamp.csv"
        $filePath = Join-Path $ExportPath $fileName
        
        $Results | Export-Csv -Path $filePath -NoTypeInformation
        Write-Log "Exported $($Results.Count) records to $filePath"
        
        return $filePath
    }
    catch {
        Write-Log "Failed to export results: $_" -Level Error
        return $null
    }
}

function Send-ExportNotification {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Summary,
        
        [Parameter(Mandatory=$false)]
        [string]$EmailRecipient
    )
    
    if (-not $EmailRecipient) { return }
    
    try {
        $subject = "Defender Detection Export Summary - $(Get-Date -Format 'yyyy-MM-dd')"
        $body = @"
Defender Detection Export Summary
================================

Export Time: $($Summary.ExportTime)

Detection Counts:
----------------
File Events: $($Summary.FileEvents)
Process Events: $($Summary.ProcessEvents)
Network Events: $($Summary.NetworkEvents)
Registry Events: $($Summary.RegistryEvents)

Details:
--------
Total Devices Affected: $($Summary.UniqueDevices)
First Detection: $($Summary.FirstDetection)
Last Detection: $($Summary.LastDetection)
Export Location: $($Summary.ExportPath)

Errors: $($Script:ErrorCount)
Warnings: $($Script:WarningCount)
"@
        
        Send-MailMessage -To $EmailRecipient `
            -Subject $subject `
            -Body $body `
            -From "defender-export@yourdomain.com" `
            -SmtpServer "smtp.yourdomain.com"
        
        Write-Log "Notification sent to $EmailRecipient"
    }
    catch {
        Write-Log "Failed to send notification: $_" -Level Warning
    }
}

function Remove-OldExports {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExportPath,
        
        [Parameter(Mandatory=$true)]
        [int]$RetentionDays
    )
    
    try {
        $oldFiles = Get-ChildItem -Path $ExportPath -File |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) }
        
        foreach ($file in $oldFiles) {
            Remove-Item -Path $file.FullName -Force
            Write-Log "Removed old export: $($file.Name)"
        }
    }
    catch {
        Write-Log "Failed to clean up old exports: $_" -Level Warning
    }
}