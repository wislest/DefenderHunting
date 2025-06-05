# IoC-Hunter.ps1
# Script pour la recherche automatisée d'IoCs dans Microsoft Defender Advanced Hunting
# Version: 1.0
# Created: 2024-11-22
# Author: Wiston Lestin
# Copyright: Copyright (c) 2024, Wiston Lestin
#
# Prérequis:
# - Module ImportExcel
# - Accès à Microsoft Defender Advanced Hunting
# - Fichier de configuration avec les identifiants d'accès

param(
    [Parameter(Mandatory=$true)]
    [string]$IocsFile,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "***REMOVED***",
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = "C:\DefenderHunting\config.json",
    
    [Parameter(Mandatory=$false)]
    [int]$LookbackDays = 30
)

# Fonction de logging
function Write-Log {
    param($Message)
    
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    $logPath = "C:\DefenderHunting\Logs\hunting.log"
    
    Write-Host $logMessage
    Add-Content -Path $logPath -Value $logMessage
}

# Fonction pour valider le fichier d'IoCs
function Test-IocsFile {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        throw "Fichier IoCs non trouvé: $FilePath"
    }
    
    $extension = [System.IO.Path]::GetExtension($FilePath)
    if ($extension -notin @('.csv', '.xlsx')) {
        throw "Format de fichier non supporté. Utilisez CSV ou XLSX."
    }
}

# Fonction pour importer les IoCs
function Import-IoCs {
    param([string]$FilePath)
    
    Write-Log "Importation des IoCs depuis $FilePath"
    
    if ($FilePath -like "*.csv") {
        $iocs = Import-Csv -Path $FilePath
    }
    elseif ($FilePath -like "*.xlsx") {
        $iocs = Import-Excel -Path $FilePath
    }
    
    # Valider la structure des IoCs
    $requiredColumns = @('FileHashes', 'FamilyName')
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $iocs[0].PSObject.Properties.Name }
    
    if ($missingColumns) {
        throw "Colonnes manquantes dans le fichier IoCs: $($missingColumns -join ', ')"
    }
    
    Write-Log "Importation réussie: $($iocs.Count) IoCs trouvés"
    return $iocs
}

# Fonction pour extraire les hashes des IoCs
function Get-HashesFromIoCs {
    param([array]$IoCs)
    
    $hashes = @{
        SHA256 = [System.Collections.Generic.HashSet[string]]::new()
        SHA1 = [System.Collections.Generic.HashSet[string]]::new()
        MD5 = [System.Collections.Generic.HashSet[string]]::new()
    }
    
    foreach ($ioc in $IoCs) {
        if ($ioc.FileHashes) {
            try {
                $hashData = $ioc.FileHashes | ConvertFrom-Json
                if ($hashData.sha256) { [void]$hashes.SHA256.Add($hashData.sha256.ToLower()) }
                if ($hashData.sha1) { [void]$hashes.SHA1.Add($hashData.sha1.ToLower()) }
                if ($hashData.md5) { [void]$hashes.MD5.Add($hashData.md5.ToLower()) }
            }
            catch {
                Write-Log "Erreur lors du traitement des hashes pour l'IoC: $($ioc.FamilyName)"
            }
        }
    }
    
    return $hashes
}

# Fonction pour générer la requête Advanced Hunting
function New-HuntingQuery {
    param(
        [hashtable]$Hashes,
        [int]$Days
    )
    
    Write-Log "Génération de la requête Advanced Hunting"
    
    $query = @"
// Requête de recherche d'IoCs dans Microsoft Defender
// Générée le: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
// Période de recherche: $Days jours

let sha256_hashes = dynamic(["$(($Hashes.SHA256 | ForEach-Object { "'" + $_ + "'" }) -join '","')"]);
let sha1_hashes = dynamic(["$(($Hashes.SHA1 | ForEach-Object { "'" + $_ + "'" }) -join '","')"]);
let md5_hashes = dynamic(["$(($Hashes.MD5 | ForEach-Object { "'" + $_ + "'" }) -join '","')"]);

// 1. Détections dans les fichiers
let FileDetections =
DeviceFileEvents
| where Timestamp > ago($($Days)d)
| where SHA256 in (sha256_hashes)
    or SHA1 in (sha1_hashes)
    or MD5 in (md5_hashes)
| extend DetectionType = "File"
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    FolderPath,
    SHA256,
    SHA1,
    MD5,
    InitiatingProcessCommandLine,
    InitiatingProcessFileName;

// 2. Détections dans les processus
let ProcessDetections =
DeviceProcessEvents
| where Timestamp > ago($($Days)d)
| where SHA256 in (sha256_hashes)
    or SHA1 in (sha1_hashes)
    or MD5 in (md5_hashes)
| extend DetectionType = "Process"
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    ProcessCommandLine,
    SHA256,
    SHA1,
    MD5,
    AccountName,
    AccountDomain;

// 3. Détections dans les connexions réseau
let NetworkDetections =
DeviceNetworkEvents
| where Timestamp > ago($($Days)d)
| where InitiatingProcessSHA256 in (sha256_hashes)
| extend DetectionType = "Network"
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    InitiatingProcessFileName,
    RemoteIP,
    RemotePort,
    RemoteUrl,
    InitiatingProcessCommandLine;

// Combiner tous les résultats
union FileDetections, ProcessDetections, NetworkDetections
| order by TimeDetected desc
"@
    
    Write-Log "Requête générée avec succès"
    return $query
}

# Fonction pour exécuter la requête via l'API
function Invoke-DefenderHunting {
    param(
        [string]$Query,
        [hashtable]$Config
    )
    
    Write-Log "Exécution de la requête Advanced Hunting"
    
    try {
        # Ici, vous devez implémenter l'appel à l'API Microsoft Defender
        # Utilisez les informations d'authentification du fichier de configuration
        
        # Exemple de structure de l'appel API:
        $token = Get-DefenderToken -Config $Config
        $url = "https://api.securitycenter.microsoft.com/api/advancedqueries/run"
        $headers = @{
            'Content-Type' = 'application/json'
            'Accept' = 'application/json'
            'Authorization' = "Bearer $token"
        }
        
        $body = @{
            'Query' = $Query
        } | ConvertTo-Json
        
        $response = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        return $response.Results
    }
    catch {
        Write-Log "Erreur lors de l'exécution de la requête: $_"
        throw
    }
}

# Fonction pour exporter les résultats
function Export-HuntingResults {
    param(
        [array]$Results,
        [string]$ExportPath
    )
    
    Write-Log "Export des résultats de la recherche"
    
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $resultsPath = Join-Path $ExportPath "HuntingResults-$timestamp.csv"
    $summaryPath = Join-Path $ExportPath "Summary-$timestamp.txt"
    
    # Export des résultats détaillés
    $Results | Export-Csv -Path $resultsPath -NoTypeInformation
    
    # Création du résumé
    $summary = @"
Résumé de la Recherche d'IoCs
============================
Date d'exécution : $(Get-Date)
Période analysée : $LookbackDays jours

Statistiques :
-------------
Total des détections : $($Results.Count)
Appareils uniques   : $(($Results.DeviceName | Select-Object -Unique).Count)

Détections par type :
-------------------
$(($Results | Group-Object DetectionType | ForEach-Object {
    "- $($_.Name): $($_.Count) détections"
}) -join "`n")

Premiers appareils affectés :
---------------------------
$(($Results | Group-Object DeviceName | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
    "- $($_.Name): $($_.Count) détections"
}) -join "`n")

Fichiers exportés :
-----------------
- Résultats détaillés : $resultsPath
- Ce résumé          : $summaryPath
"@
    
    $summary | Out-File -FilePath $summaryPath
    
    Write-Log "Résultats exportés vers: $resultsPath"
    Write-Log "Résumé exporté vers: $summaryPath"
    
    return @{
        ResultsPath = $resultsPath
        SummaryPath = $summaryPath
    }
}

# Bloc d'exécution principal
try {
    Write-Log "Démarrage de la recherche d'IoCs"
    
    # 1. Vérifier et créer les dossiers nécessaires
    $ExportPath = [System.IO.Path]::GetFullPath($ExportPath)
    if (-not (Test-Path $ExportPath)) {
        New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    }
    
    # 2. Charger la configuration
    $config = Get-Content $ConfigPath | ConvertFrom-Json -AsHashtable
    
    # 3. Valider et importer les IoCs
    Test-IocsFile -FilePath $IocsFile
    $iocs = Import-IoCs -FilePath $IocsFile
    
    # 4. Extraire les hashes
    $hashes = Get-HashesFromIoCs -IoCs $iocs
    
    # 5. Générer la requête
    $query = New-HuntingQuery -Hashes $hashes -Days $LookbackDays
    
    # 6. Exécuter la recherche
    $results = Invoke-DefenderHunting -Query $query -Config $config
    
    # 7. Exporter les résultats
    $exportPaths = Export-HuntingResults -Results $results -ExportPath $ExportPath
    
    Write-Log "Recherche terminée avec succès"
    Write-Host "`nRésultats disponibles dans:"
    Write-Host "- $($exportPaths.ResultsPath)"
    Write-Host "- $($exportPaths.SummaryPath)"
}
catch {
    Write-Log "ERREUR: $_" -Level Error
    throw
}