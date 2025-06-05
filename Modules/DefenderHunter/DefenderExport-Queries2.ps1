# IoC Analyzer and Detection Export Tool
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
# Description: A PowerShell script to analyze IoCs from a CSV file and generate hunting queries for Microsoft Defender for Endpoint. 
# It also provides options to configure automated exports and run manual exports of detection results.

function Get-DefenderAdvancedHuntingQueriesries {
    # Returns a hashtable of predefined queries
    return @{
        FileEvents = @"
// Advanced Hunting Query - File Events
DeviceFileEvents
| where Timestamp > ago(24h)
| where SHA256 in (RansomwareHashes) or 
        SHA1 in (RansomwareHashes) or 
        MD5 in (RansomwareHashes)
| project
    Timestamp,
    DeviceName,
    DeviceId,
    FileName,
    FolderPath,
    SHA256,
    SHA1,
    MD5,
    InitiatingProcessFileName,
    InitiatingProcessCommandLine,
    InitiatingProcessAccountName,
    InitiatingProcessAccountDomain,
    FileSize,
    ActionType
"@

        ProcessEvents = @"
// Advanced Hunting Query - Process Events
DeviceProcessEvents
| where Timestamp > ago(24h)
| where SHA256 in (RansomwareHashes) or 
        InitiatingProcessSHA256 in (RansomwareHashes)
| project
    Timestamp,
    DeviceName,
    DeviceId,
    ActionType,
    FileName,
    ProcessCommandLine,
    InitiatingProcessFileName,
    InitiatingProcessCommandLine,
    AccountName,
    AccountDomain,
    SHA256,
    ProcessId,
    InitiatingProcessParentId,
    ProcessVersionInfoProductName,
    ProcessVersionInfoOriginalFileName
"@

        NetworkConnections = @"
// Advanced Hunting Query - Network Connections
DeviceNetworkEvents
| where Timestamp > ago(24h)
| where InitiatingProcessSHA256 in (RansomwareHashes)
| project
    Timestamp,
    DeviceName,
    DeviceId,
    ActionType,
    InitiatingProcessFileName,
    InitiatingProcessCommandLine,
    RemoteIP,
    RemotePort,
    RemoteUrl,
    LocalIP,
    LocalPort,
    Protocol,
    InitiatingProcessAccountName,
    InitiatingProcessAccountDomain
"@

        RegistryEvents = @"
// Advanced Hunting Query - Registry Events
DeviceRegistryEvents
| where Timestamp > ago(24h)
| where InitiatingProcessSHA256 in (RansomwareHashes)
| project
    Timestamp,
    DeviceName,
    DeviceId,
    ActionType,
    RegistryKey,
    RegistryValueName,
    RegistryValueData,
    InitiatingProcessFileName,
    InitiatingProcessCommandLine,
    InitiatingProcessAccountName
"@

        RansomwareIndicators = @"
// Advanced Hunting Query - Ransomware Indicators
let SuspiciousExtensions = dynamic([".encrypted", ".locked", ".crypto", ".crypt", ".криптор", ".cry", ".raid"]);
let SuspiciousProcesses = dynamic(["vssadmin.exe", "wevtutil.exe", "bcdedit.exe", "wbadmin.exe"]);
let SuspiciousCommands = dynamic([
    "delete shadows",
    "delete catalog",
    "delete backup",
    "recoveryenabled no",
    "shadowcopy delete"
]);
// File Extension Changes
let FileExtEvents = 
    DeviceFileEvents
    | where Timestamp > ago(24h)
    | where FileName has_any (SuspiciousExtensions)
    | project
        Timestamp,
        DeviceName,
        FileName,
        FolderPath,
        InitiatingProcessFileName,
        InitiatingProcessCommandLine;
// Suspicious Process Executions
let SuspiciousProcs = 
    DeviceProcessEvents
    | where Timestamp > ago(24h)
    | where FileName in~ (SuspiciousProcesses)
        or ProcessCommandLine has_any (SuspiciousCommands)
    | project
        Timestamp,
        DeviceName,
        FileName,
        ProcessCommandLine,
        AccountName;
// Combine Results
union FileExtEvents, SuspiciousProcs
| order by Timestamp desc
"@

        SummaryStats = @"
// Advanced Hunting Query - Detection Summary
let timeframe = 24h;
let FileDetections = 
    DeviceFileEvents
    | where Timestamp > ago(timeframe)
    | where SHA256 in (RansomwareHashes)
    | summarize
        FileCount = count(),
        UniqueDevices = dcount(DeviceId),
        FirstSeen = min(Timestamp),
        LastSeen = max(Timestamp)
    | extend DetectionType = "File";
let ProcessDetections =
    DeviceProcessEvents
    | where Timestamp > ago(timeframe)
    | where SHA256 in (RansomwareHashes)
    | summarize
        ProcessCount = count(),
        UniqueDevices = dcount(DeviceId),
        FirstSeen = min(Timestamp),
        LastSeen = max(Timestamp)
    | extend DetectionType = "Process";
let NetworkDetections =
    DeviceNetworkEvents
    | where Timestamp > ago(timeframe)
    | where InitiatingProcessSHA256 in (RansomwareHashes)
    | summarize
        ConnectionCount = count(),
        UniqueDevices = dcount(DeviceId),
        FirstSeen = min(Timestamp),
        LastSeen = max(Timestamp)
    | extend DetectionType = "Network";
// Combine all detections
union FileDetections, ProcessDetections, NetworkDetections
| project
    DetectionType,
    EventCount = coalesce(FileCount, ProcessCount, ConnectionCount),
    UniqueDevices,
    FirstSeen,
    LastSeen,
    Duration = LastSeen - FirstSeen
"@

        ImpactedDevices = @"
// Advanced Hunting Query - Impacted Devices Detail
let timeframe = 24h;
let ImpactedDeviceIds =
    DeviceFileEvents
    | where Timestamp > ago(timeframe)
    | where SHA256 in (RansomwareHashes)
    | distinct DeviceId;
// Get device details
DeviceInfo
| where DeviceId in (ImpactedDeviceIds)
| project
    DeviceName,
    DeviceId,
    OSPlatform,
    OSVersion,
    OSVersionInfo,
    OSBuild,
    JoinType,
    ClientVersion,
    PublicIP,
    Model,
    LoggedOnUsers,
    MachineGroup,
    OnboardingStatus
| join kind=leftouter (
    DeviceNetworkInfo
    | where DeviceId in (ImpactedDeviceIds)
    | summarize arg_max(Timestamp, *) by DeviceId
    | project DeviceId, MacAddress, IPAddresses
) on DeviceId
"@
    }
}

function Invoke-DefenderQuery {
    param(
        [Parameter(Mandatory=$true)]
        [string]$WorkspaceId,
        
        [Parameter(Mandatory=$true)]
        [string]$Query
    )
    
    try {
        Write-Log "Executing query against workspace: $WorkspaceId"
        $results = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $Query
        
        if ($results.Error) {
            Write-Log "Query execution failed: $($results.Error)" -Level Error
            return $null
        }
        
        Write-Log "Query executed successfully. Retrieved $($results.Results.Count) results."
        return $results.Results
    }
    catch {
        Write-Log "Error executing query: $_" -Level Error
        return $null
    }
}