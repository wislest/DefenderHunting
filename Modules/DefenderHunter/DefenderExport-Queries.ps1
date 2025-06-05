# DefenderExport-Queries.ps1
# Requêtes Kusto pour Advanced Hunting basées sur les IoCs fournis

function Get-KustoQueries {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Hashes,  # Hashes par famille de malware
        [int]$LookbackDays = 30
    )

    # Requête par famille de malware
    $familyQueries = @{}
    foreach ($family in $Hashes.Keys) {
        $familyQueries[$family] = @"
// Détection pour la famille $family
let ${family}_SHA256 = dynamic([$($Hashes[$family].SHA256 -join ',')]);
let ${family}_SHA1 = dynamic([$($Hashes[$family].SHA1 -join ',')]);
let ${family}_MD5 = dynamic([$($Hashes[$family].MD5 -join ',')]);

// 1. Détection de fichiers
let ${family}_Files =
DeviceFileEvents
| where Timestamp > ago($LookbackDays d)
| where SHA256 in (${family}_SHA256)
    or SHA1 in (${family}_SHA1)
    or MD5 in (${family}_MD5)
| extend 
    DetectionType = "File",
    MalwareFamily = "$family";

// 2. Détection de processus
let ${family}_Processes =
DeviceProcessEvents
| where Timestamp > ago($LookbackDays d)
| where SHA256 in (${family}_SHA256)
    or InitiatingProcessSHA256 in (${family}_SHA256)
| extend 
    DetectionType = "Process",
    MalwareFamily = "$family";

// 3. Détection réseau
let ${family}_Network =
DeviceNetworkEvents
| where Timestamp > ago($LookbackDays d)
| where InitiatingProcessSHA256 in (${family}_SHA256)
| extend 
    DetectionType = "Network",
    MalwareFamily = "$family";

// Combiner les détections
union ${family}_Files, ${family}_Processes, ${family}_Network
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    MalwareFamily,
    FileName,
    SHA256,
    SHA1,
    MD5,
    ProcessCommandLine = coalesce(ProcessCommandLine, InitiatingProcessCommandLine),
    RemoteIP,
    RemoteUrl,
    AccountName
| order by TimeDetected desc
"@
    }

    # Requête combinée pour toutes les familles
    $allFamiliesQuery = @"
// Requête combinée pour toutes les familles de malware
let AllDetections =
$(foreach ($family in $Hashes.Keys) {
    "union (DeviceFileEvents
    | where Timestamp > ago($LookbackDays d)
    | where SHA256 in (${family}_SHA256)
        or SHA1 in (${family}_SHA1)
        or MD5 in (${family}_MD5)
    | extend MalwareFamily = `"$family`", DetectionType = `"File`")"
})
| project
    TimeDetected = Timestamp,
    DeviceName,
    MalwareFamily,
    DetectionType,
    FileName,
    SHA256,
    ProcessCommandLine = InitiatingProcessCommandLine;

// Résumé des détections
AllDetections
| summarize
    DetectionCount = count(),
    AffectedDevices = dcount(DeviceName),
    FirstSeen = min(TimeDetected),
    LastSeen = max(TimeDetected)
    by MalwareFamily
| extend DetectionTimespan = LastSeen - FirstSeen
| order by DetectionCount desc
"@

    # Requêtes spécifiques aux comportements de ransomware
    $ransomwareQuery = @"
// Détection de comportements typiques de ransomware
let SuspiciousExtensions = dynamic([".encrypted", ".locked", ".crypto", ".crypt", ".wcry", ".wncry", ".akira"]);
let SuspiciousProcesses = dynamic([
    "vssadmin.exe", 
    "wevtutil.exe", 
    "bcdedit.exe",
    "wbadmin.exe",
    "psexec.exe",
    "mimikatz.exe"
]);
let SuspiciousCommands = dynamic([
    "delete shadows",
    "delete catalog",
    "delete backup",
    "recoveryenabled no",
    "shadowcopy delete"
]);

// 1. Détection d'extensions suspectes
DeviceFileEvents
| where Timestamp > ago($LookbackDays d)
| where FileName has_any (SuspiciousExtensions)
| extend DetectionType = "Suspicious Extension"
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    FolderPath,
    SHA256;

// 2. Détection de processus suspects
DeviceProcessEvents
| where Timestamp > ago($LookbackDays d)
| where FileName in~ (SuspiciousProcesses)
    or ProcessCommandLine has_any (SuspiciousCommands)
| extend DetectionType = "Suspicious Process"
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    ProcessCommandLine,
    AccountName
"@

    return @{
        FamilyQueries = $familyQueries
        AllFamilies = $allFamiliesQuery
        RansomwareBehaviors = $ransomwareQuery
    }
}