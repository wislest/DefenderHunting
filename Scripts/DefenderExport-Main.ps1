# C:\DefenderHunting\Scripts\DefenderExport-Main.ps1
function Start-DefenderHunting {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$IocsFile,
        [int]$LookbackDays = 30
    )
    
    Write-Host "Starting IoC hunting with file: $IocsFile"
    
    try {
        # 1. Importer les IoCs
        if (!(Test-Path $IocsFile)) {
            throw "Fichier IoC non trouvé: $IocsFile"
        }
        
        $iocs = Import-Csv $IocsFile
        Write-Host "Importé $($iocs.Count) IoCs"
        
        # 2. Traiter les IoCs
        $hashesByFamily = @{}
        foreach ($ioc in $iocs) {
            if ($ioc.FileHashes -and $ioc.FamilyName) {
                $hashes = $ioc.FileHashes | ConvertFrom-Json
                if (!$hashesByFamily[$ioc.FamilyName]) {
                    $hashesByFamily[$ioc.FamilyName] = @{
                        SHA256 = @()
                        SHA1 = @()
                        MD5 = @()
                    }
                }
                if ($hashes.sha256) { $hashesByFamily[$ioc.FamilyName].SHA256 += "'$($hashes.sha256)'" }
                if ($hashes.sha1) { $hashesByFamily[$ioc.FamilyName].SHA1 += "'$($hashes.sha1)'" }
                if ($hashes.md5) { $hashesByFamily[$ioc.FamilyName].MD5 += "'$($hashes.md5)'" }
            }
        }
        
        # 3. Générer les requêtes Kusto
        Write-Host "Génération des requêtes Kusto..."
        foreach ($family in $hashesByFamily.Keys) {
            $query = @"
// Détections pour la famille: $family
let timerange = $LookbackDays d;
let SHA256_Hashes = dynamic([$(($hashesByFamily[$family].SHA256 | Select-Object -Unique) -join ',')]);
let SHA1_Hashes = dynamic([$(($hashesByFamily[$family].SHA1 | Select-Object -Unique) -join ',')]);
let MD5_Hashes = dynamic([$(($hashesByFamily[$family].MD5 | Select-Object -Unique) -join ',')]);

// Fichiers
let FileEvents =
DeviceFileEvents
| where Timestamp > ago(timerange)
| where SHA256 in (SHA256_Hashes)
    or SHA1 in (SHA1_Hashes)
    or MD5 in (MD5_Hashes)
| extend DetectionType = "File";

// Processus
let ProcessEvents =
DeviceProcessEvents
| where Timestamp > ago(timerange)
| where SHA256 in (SHA256_Hashes)
    or InitiatingProcessSHA256 in (SHA256_Hashes)
| extend DetectionType = "Process";

// Réseau
let NetworkEvents =
DeviceNetworkEvents
| where Timestamp > ago(timerange)
| where InitiatingProcessSHA256 in (SHA256_Hashes)
| extend DetectionType = "Network";

// Combiner les résultats
union FileEvents, ProcessEvents, NetworkEvents
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    SHA256,
    CommandLine = coalesce(ProcessCommandLine, InitiatingProcessCommandLine),
    AccountName,
    RemoteIP,
    RemoteUrl
| order by TimeDetected desc
"@
            
            # Sauvegarder la requête
            $queryPath = Join-Path $PWD "Queries\$family.kql"
            $query | Out-File -Path $queryPath -Encoding UTF8
            Write-Host "Requête générée pour $family : $queryPath"
        }
        
        Write-Host "Opération terminée avec succès !"
    }
    catch {
        Write-Error "Erreur lors du traitement : $_"
    }
}

# Exporter la fonction
Export-ModuleMember -Function Start-DefenderHunting