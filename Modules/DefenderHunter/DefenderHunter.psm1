
# DefenderHunter.psm1
# Module pour l'analyse des IoCs via Microsoft Defender Advanced Hunting
# Version: 2.0
# Auteur: Wiston Lestin
# Définition de l'encodage UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8


# Installation du module ImportExcel si nécessaire
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installation du module ImportExcel..."
    Install-Module ImportExcel -Force -Scope CurrentUser
}

# Fonction pour la gestion des noms de fichiers sécurisés
function Get-SafeFileName {
    param(
        [string]$Name
    )
    # Remplacer tous les caractères non-alphanumériques par un underscore
    $safeName = $Name -replace '[^a-zA-Z0-9]', '_'
    return $safeName
}

# Fonction pour valider la requête Kusto
function Test-KustoQuery {
    param(
        [string]$Query
    )
    
    try {
        # Vérifier la longueur des tableaux de hashes
        if ($Query -match "dynamic\(\[\s*\]\)") {
            Write-Warning "La requête contient des tableaux vides"
            return $false
        }
        
        # Vérifier la syntaxe de base
        if ($Query -notmatch "where|project|extend") {
            Write-Warning "La requête pourrait être invalide (manque de clauses essentielles)"
            return $false
        }
        return $true
    }
    catch {
        Write-Warning "Erreur lors de la validation de la requête: $_"
        return $false
    }
}

# Fonction pour la connexion à l'API Defender
function Connect-DefenderAPI {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantId,
        [Parameter(Mandatory=$true)]
        [string]$ClientId,
        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )
    
    try {
        Write-Host "Tentative de connexion à l'API Microsoft Defender..."
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        $body = @{
            client_id = $ClientId
            client_secret = $ClientSecret
            scope = "https://api.securitycenter.microsoft.com/.default"
            grant_type = "client_credentials"
        }
        
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
        Write-Host "Connexion réussie."
        return $response.access_token
    }
    catch {
        throw "Erreur de connexion à l'API Defender: $_"
    }
}

# Fonction pour exécuter une requête via l'API
function Invoke-DefenderQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Query,
        [Parameter(Mandatory=$true)]
        [string]$Token,
        [string]$FamilyName = "Unknown"
    )
    
    try {
        # Valider la requête d'abord
        if (-not (Test-KustoQuery -Query $Query)) {
            Write-Warning "La requête pour $FamilyName n'est pas valide"
            return $null
        }

        $headers = @{
            'Content-Type' = 'application/json'
            'Accept' = 'application/json'
            'Authorization' = "Bearer $Token"
        }
        
        $cleanQuery = $Query.Trim()
        $body = @{
            'Query' = $cleanQuery
        } | ConvertTo-Json -Depth 10 -Compress

        Write-Verbose "Envoi de la requête pour $FamilyName..."
        Write-Verbose "Requête: $cleanQuery"
        
        $response = Invoke-RestMethod -Uri "https://api.securitycenter.microsoft.com/api/advancedqueries/run" `
                                    -Method Post `
                                    -Headers $headers `
                                    -Body $body `
                                    -ContentType "application/json; charset=utf-8" `
                                    -ErrorAction Stop

        Write-Verbose "Requête exécutée avec succès"

        # Vérifier la propriété Results (même si c'est un tableau vide)
        if ($response.PSObject.Properties.Name -contains 'Results') {
            $resultCount = if ($response.Results) { $response.Results.Count } else { 0 }
            Write-Verbose "Nombre de résultats retournés par l'API : $resultCount"

            if ($resultCount -gt 0) {
                return $response.Results
            }
            else {
                Write-Verbose "Tableau Results vide - aucune détection"
                return $null
            }
        }
        else {
            Write-Verbose "Aucune propriété 'Results' dans la réponse de l'API"
            Write-Verbose "Contenu de la réponse : $($response | ConvertTo-Json -Depth 2)"
            return $null
        }
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            
            Write-Verbose "Requête qui a échoué: $cleanQuery"
            Write-Verbose "Corps de la requête envoyée: $body"
            Write-Verbose "Réponse complète: $responseBody"
            
            throw "Erreur API ($statusCode): $responseBody"
        }
        catch {
            throw "Erreur lors de l'exécution de la requête: $_"
        }
    }
}

# Fonction principale pour l'analyse des IoCs
function Start-DefenderHunting {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$IocsFile,
        
        [Parameter(Mandatory=$false)]
        [string]$TenantId,
        
        [Parameter(Mandatory=$false)]
        [string]$ClientId,
        
        [Parameter(Mandatory=$false)]
        [string]$ClientSecret,
        
        [Parameter(Mandatory=$false)]
        [string]$ConfigPath = "C:\DefenderHunting\config.json",
        
        [int]$LookbackDays = 30,
        [int]$MaxHashesPerQuery = 50
    )
    
    try {
        # Chargement de la configuration
        if (Test-Path $ConfigPath) {
            Write-Verbose "Chargement de la configuration depuis $ConfigPath"
            $config = Get-Content $ConfigPath | ConvertFrom-Json
            
            # Utilisation des valeurs du fichier config si non fournies en paramètre
            if (-not $IocsFile -and $config.IocsFile) {
                $IocsFile = $config.IocsFile
                Write-Verbose "Utilisation du fichier IoCs depuis config: $IocsFile"
            }
            
            if (-not $TenantId -and $config.TenantId) { $TenantId = $config.TenantId }
            if (-not $ClientId -and $config.ClientId) { $ClientId = $config.ClientId }
            if (-not $ClientSecret -and $config.ClientSecret) { $ClientSecret = $config.ClientSecret }
            
            if ($config.LookbackDays) { $LookbackDays = $config.LookbackDays }
            if ($config.MaxHashesPerQuery) { $MaxHashesPerQuery = $config.MaxHashesPerQuery }
        }
        
        # Vérification des paramètres requis
        if (-not $IocsFile) {
            throw "Le fichier IoCs n'est pas spécifié. Fournissez-le en paramètre ou dans config.json"
        }
        if (-not ($TenantId -and $ClientId -and $ClientSecret)) {
            throw "Informations d'authentification manquantes. Fournissez-les en paramètre ou dans config.json"
        }

        Write-Host "Démarrage de l'analyse Threat Hunting..."
        Write-Host "Traitement du fichier IoCs: $IocsFile"

        # Création des dossiers nécessaires
        $basePath = "C:\DefenderHunting"
        $paths = @{
            Queries = Join-Path $basePath "Queries"
            Exports = Join-Path $basePath "Exports"
            Logs = Join-Path $basePath "Logs"
        }

        foreach ($dir in $paths.Values) {
            if (!(Test-Path $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                Write-Host "Dossier créé: $dir"
            }
        }

        # Connexion à l'API
        $token = Connect-DefenderAPI -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

        # Importation et vérification des IoCs
        if (!(Test-Path $IocsFile)) {
            throw "Fichier IoCs introuvable: $IocsFile"
        }
        $iocs = Import-Csv $IocsFile
        Write-Host "Importé $($iocs.Count) IoCs"

        # Création du dossier pour cette session d'export
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $sessionExportPath = Join-Path $paths.Exports $timestamp
        New-Item -ItemType Directory -Path $sessionExportPath -Force | Out-Null

        # Préparation des variables Excel (création à la demande)
        $excelPath = Join-Path $sessionExportPath "DefenderHunting_Results_$timestamp.xlsx"
        $excelCreated = $false

        # Traitement et organisation des IoCs par famille
        $hashesByFamily = @{}
        $detectionResults = @{} # Tracker les résultats de détection par hash

        foreach ($ioc in $iocs) {
            if (-not $ioc.FileHashes) {
                Write-Warning "Ligne ignorée - Pas de hashes : $($ioc | ConvertTo-Json)"
                continue
            }
        
            # Définition du nom de famille
            $familyName = if ([string]::IsNullOrWhiteSpace($ioc.FamilyName)) {
                "UnknownFamily_$(Get-Date -Format 'yyyyMMdd')"
            } else {
                $ioc.FamilyName
            }
            
            $safeName = Get-SafeFileName -Name $familyName
            
            # Initialisation de la structure pour la famille si nécessaire
            if (!$hashesByFamily.ContainsKey($safeName)) {
                $hashesByFamily[$safeName] = @{
                    OriginalName = $familyName
                    SHA256 = [System.Collections.Generic.HashSet[string]]::new()
                    SHA1 = [System.Collections.Generic.HashSet[string]]::new()
                    MD5 = [System.Collections.Generic.HashSet[string]]::new()
                    Paths = [System.Collections.Generic.HashSet[string]]::new()
                }
            }
        
            try {
                # Traitement des hashes avec validation JSON
                $hashesJson = $ioc.FileHashes.Trim()

                # Vérifier si c'est un JSON valide ou un simple hash
                if ($hashesJson -match '^\{.*\}$') {
                    # Format JSON structuré
                    try {
                        $hashes = $hashesJson | ConvertFrom-Json -ErrorAction Stop
                        if ($hashes.sha256 -and $hashes.sha256 -match '^[a-fA-F0-9]{64}$') {
                            [void]$hashesByFamily[$safeName].SHA256.Add("'$($hashes.sha256)'")
                        }
                        if ($hashes.sha1 -and $hashes.sha1 -match '^[a-fA-F0-9]{40}$') {
                            [void]$hashesByFamily[$safeName].SHA1.Add("'$($hashes.sha1)'")
                        }
                        if ($hashes.md5 -and $hashes.md5 -match '^[a-fA-F0-9]{32}$') {
                            [void]$hashesByFamily[$safeName].MD5.Add("'$($hashes.md5)'")
                        }
                    }
                    catch {
                        Write-Warning "Erreur de parsing JSON pour $familyName : $_. Hash brut: $hashesJson"
                        continue
                    }
                }
                elseif ($hashesJson -match '^[a-fA-F0-9]+$') {
                    # Hash simple sans structure JSON
                    $hashLength = $hashesJson.Length
                    if ($hashLength -eq 32) {
                        [void]$hashesByFamily[$safeName].MD5.Add("'$hashesJson'")
                    }
                    elseif ($hashLength -eq 40) {
                        [void]$hashesByFamily[$safeName].SHA1.Add("'$hashesJson'")
                    }
                    elseif ($hashLength -eq 64) {
                        [void]$hashesByFamily[$safeName].SHA256.Add("'$hashesJson'")
                    }
                    else {
                        Write-Warning "Hash avec longueur invalide ($hashLength) pour $familyName : $hashesJson"
                        continue
                    }
                }
                else {
                    Write-Warning "Format de hash non reconnu pour $familyName : $hashesJson"
                    continue
                }

                # Traitement du chemin complet
                if (-not [string]::IsNullOrWhiteSpace($ioc.'FullPath')) {
                    [void]$hashesByFamily[$safeName].Paths.Add($ioc.'FullPath')
                }
            }
            catch {
                Write-Warning "Erreur inattendue lors du traitement pour $familyName : $_"
                continue
            }
        }

        # Traitement de chaque famille de malware
        foreach ($safeName in $hashesByFamily.Keys) {
            $originalName = $hashesByFamily[$safeName].OriginalName
            Write-Host "`nTraitement de la famille: $originalName"

            $allHashes = @{
                SHA256 = $hashesByFamily[$safeName].SHA256
                SHA1 = $hashesByFamily[$safeName].SHA1
                MD5 = $hashesByFamily[$safeName].MD5
            }

            # Template de la requête KQL
            $baseQuery = @"
// Requête de détection pour: $originalName
// Générée le: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

let timerange = $($LookbackDays)d;
{0}
// Détection des événements fichiers
let FileEvents =
DeviceFileEvents
| where Timestamp > ago(timerange)
| where {1}
| extend DetectionType = "File"
| summarize arg_max(Timestamp, *) by DeviceName, FileName, SHA256
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    FolderPath,
    SHA256,
    SHA1,
    MD5,
    InitiatingProcessCommandLine;
// Détection des événements processus
let ProcessEvents =
DeviceProcessEvents
| where Timestamp > ago(timerange)
| where {2}
| extend DetectionType = "Process"
| summarize arg_max(Timestamp, *) by DeviceName, FileName, SHA256
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    FileName,
    ProcessCommandLine,
    AccountName,
    SHA256;
// Détection des événements réseau
let NetworkEvents =
DeviceNetworkEvents
| where Timestamp > ago(timerange)
| where {3}
| extend DetectionType = "Network"
| summarize arg_max(Timestamp, *) by DeviceName, InitiatingProcessFileName, InitiatingProcessSHA256
| project
    TimeDetected = Timestamp,
    DeviceName,
    DetectionType,
    InitiatingProcessFileName,
    RemoteIP,
    RemoteUrl,
    RemotePort;
// Consolidation des résultats
union FileEvents, ProcessEvents, NetworkEvents
| order by TimeDetected desc
| extend Entity = case(
    DetectionType == "File", FileName,
    DetectionType == "Network", RemoteIP,
    DetectionType == "Process", AccountName,
    "Unknown")
"@

            # Construction des conditions de recherche
            $hashDeclarations = ''
            $fileConditions = ''
            $processConditions = ''
            $networkConditions = ''

            # Génération des déclarations et conditions pour chaque type de hash
            foreach ($hashType in @('SHA256', 'SHA1', 'MD5')) {
                if ($allHashes[$hashType].Count -gt 0) {
                    # Déclaration des variables de hashes
                    $hashDeclarations += "let ${hashType}_Hashes = dynamic([" + ($allHashes[$hashType] -join ',') + "]);"
                    
                    # Construction des conditions selon le type d'événement
                    if ($hashType -eq 'SHA256') {
                        $fileConditions += "$hashType in (${hashType}_Hashes) or "
                        $processConditions += "$hashType in (${hashType}_Hashes) or InitiatingProcess$hashType in (${hashType}_Hashes) or "
                        $networkConditions += "InitiatingProcess$hashType in (${hashType}_Hashes) or "
                    } else {
                        $fileConditions += "$hashType in (${hashType}_Hashes) or "
                    }
                }
            }

            # Nettoyage des conditions
            $fileConditions = $fileConditions.TrimEnd(' or ')
            $processConditions = $processConditions.TrimEnd(' or ')
            $networkConditions = $networkConditions.TrimEnd(' or ')

            # Formation de la requête finale
            $query = $baseQuery -f $hashDeclarations, $fileConditions, $processConditions, $networkConditions

            # Sauvegarde de la requête
            $queryPath = Join-Path $paths.Queries "$safeName.kql"
            $query | Out-File -FilePath $queryPath -Force
            Write-Host "Requête générée: $queryPath"

            # Exécution de la requête et traitement des résultats
            try {
                Write-Host "Exécution de la requête pour $originalName..."
                $results = Invoke-DefenderQuery -Query $query -Token $token -FamilyName $originalName
            
                # Préparation du nom de la feuille Excel (limite de 31 caractères)
                $worksheetName = $safeName
                if ($worksheetName.Length -gt 31) {
                    $worksheetName = $worksheetName.Substring(0, 28) + "..."
                }
            
                if ($results -and $results.Count -gt 0) {
                    try {
                        # Marquer qu'on a trouvé des résultats
                        $excelCreated = $true

                        # Tracker les détections par hash
                        $machinesAffected = @()
                        foreach ($result in $results) {
                            $sha = if ($result.SHA256) { $result.SHA256 } else { "" }
                            $device = if ($result.DeviceName) { $result.DeviceName } else { "" }

                            if ($sha -and $device) {
                                if (-not $detectionResults.ContainsKey($sha)) {
                                    $detectionResults[$sha] = @{
                                        Count = 0
                                        Machines = [System.Collections.Generic.HashSet[string]]::new()
                                        Family = $originalName
                                    }
                                }
                                $detectionResults[$sha].Count++
                                [void]$detectionResults[$sha].Machines.Add($device)
                            }

                            if ($device -and $device -notin $machinesAffected) {
                                $machinesAffected += $device
                            }
                        }

                        # Nettoyer les résultats - remplacer les valeurs null par des chaînes vides
                        $cleanResults = @()
                        foreach ($result in $results) {
                            $obj = [PSCustomObject]@{}
                            foreach ($prop in $result.PSObject.Properties) {
                                $value = if ($null -eq $prop.Value) { "" } else { $prop.Value }
                                $obj | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $value
                            }
                            $cleanResults += $obj
                        }

                        # Export simple vers Excel (compatible avec toutes les versions d'ImportExcel)
                        $cleanResults | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Show:$false

                        Write-Host "  $($results.Count) détections trouvées et exportées vers Excel"
                        if ($machinesAffected.Count -gt 0) {
                            Write-Host "    Machines affectées: $($machinesAffected -join ', ')" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Warning "Erreur lors de l'export Excel pour $originalName : $_"
                        Write-Verbose "Détails de l'erreur : $($_.Exception.Message)"
                        Write-Verbose "Stack trace : $($_.ScriptStackTrace)"
                    }
                }
                else {
                    Write-Host "  Aucune détection trouvée"
                }
            }
            catch {
                Write-Warning "Erreur lors de l'exécution de la requête pour $originalName : $_"
                continue
            }
        } # Fin de la boucle foreach

        # Création de la feuille de résumé des IoCs
        if ($excelCreated) {
            Write-Host "`nCréation de la feuille de résumé des IoCs..."

            try {
                # Importer à nouveau le CSV source pour créer le résumé
                $iocsSummary = @()
                $iocsSource = Import-Csv $IocsFile

                foreach ($ioc in $iocsSource) {
                    # Parser les hashes
                    $sha256 = ""
                    $sha1 = ""
                    $md5 = ""

                    try {
                        $hashesJson = $ioc.FileHashes.Trim()
                        if ($hashesJson -match '^\{.*\}$') {
                            $hashes = $hashesJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                            $sha256 = if ($hashes.sha256) { $hashes.sha256 } else { "" }
                            $sha1 = if ($hashes.sha1) { $hashes.sha1 } else { "" }
                            $md5 = if ($hashes.md5) { $hashes.md5 } else { "" }
                        }
                    }
                    catch {
                        # Ignorer les erreurs de parsing
                    }

                    # Vérifier si ce hash a été détecté
                    $detected = $false
                    $detectionCount = 0
                    $affectedMachines = ""

                    if ($sha256 -and $detectionResults.ContainsKey($sha256)) {
                        $detected = $true
                        $detectionCount = $detectionResults[$sha256].Count
                        $affectedMachines = ($detectionResults[$sha256].Machines | Sort-Object) -join '; '
                    }

                    $summaryObj = [PSCustomObject]@{
                        'Famille' = if ($ioc.FamilyName) { $ioc.FamilyName } else { "Unknown" }
                        'Chemin Fichier' = if ($ioc.FullPath) { $ioc.FullPath } else { "N/A" }
                        'SHA256' = $sha256
                        'SHA1' = $sha1
                        'MD5' = $md5
                        'Statut' = if ($detected) { "DETECTE" } else { "Non detecte" }
                        'Detections' = $detectionCount
                        'Machines Affectees' = $affectedMachines
                    }

                    $iocsSummary += $summaryObj
                }

                # Exporter le resume comme premiere feuille
                $iocsSummary | Export-Excel -Path $excelPath -WorksheetName "0_Resume_IoCs" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -MoveToStart

                Write-Host "Feuille de resume creee avec succes ($($iocsSummary.Count) IoCs)"

                # Statistiques du resume
                $detectedItems = @($iocsSummary | Where-Object { $_.Statut -eq "DETECTE" })
                $notDetectedItems = @($iocsSummary | Where-Object { $_.Statut -eq "Non detecte" })
                $detected = $detectedItems.Count
                $notDetected = $notDetectedItems.Count
                Write-Host "  - IoCs detectes : $detected" -ForegroundColor $(if ($detected -gt 0) { "Red" } else { "Gray" })
                Write-Host "  - IoCs non detectes : $notDetected" -ForegroundColor Green
            }
            catch {
                Write-Warning "Erreur lors de la création de la feuille de résumé : $_"
            }
        }

        # Finalisation
        if ($excelCreated) {
            Write-Host "`nRapport Excel généré : $excelPath"
            Write-Host "Vous pouvez ouvrir le fichier pour voir les détections."
        }
        else {
            Write-Host "`nAucune détection trouvée - Aucun fichier Excel généré"
        }
    }
    catch {
        Write-Error "Erreur lors de l'analyse : $_"
    }

} # Fin de la fonction Start-DefenderHunting

# Export de la fonction principale
Export-ModuleMember -Function Start-DefenderHunting