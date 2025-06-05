
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
        return $response.Results
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

        # Création du fichier Excel
        $excelPath = Join-Path $sessionExportPath "DefenderHunting_Results_$timestamp.xlsx"
        $excel = Export-Excel -Path $excelPath -PassThru

        # Traitement et organisation des IoCs par famille
        $hashesByFamily = @{}
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
                # Traitement des hashes
                $hashes = $ioc.FileHashes | ConvertFrom-Json
                if ($hashes.sha256) { [void]$hashesByFamily[$safeName].SHA256.Add("'$($hashes.sha256)'") }
                if ($hashes.sha1) { [void]$hashesByFamily[$safeName].SHA1.Add("'$($hashes.sha1)'") }
                if ($hashes.md5) { [void]$hashesByFamily[$safeName].MD5.Add("'$($hashes.md5)'") }
                
                # Traitement du chemin complet
                if (-not [string]::IsNullOrWhiteSpace($ioc.'FullPath')) { 
                    [void]$hashesByFamily[$safeName].Paths.Add($ioc.'FullPath')
                }
            }
            catch {
                Write-Warning "Erreur lors du traitement des hashes pour $familyName : $_"
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
            
                if ($results) {
                    try {
                        # Création d'une nouvelle feuille de calcul
                        $worksheet = Add-Worksheet -ExcelPackage $excel -WorksheetName $worksheetName -ErrorAction Stop
                        
                        # Export des résultats vers Excel avec vérification
                        $results | Export-Excel -ExcelPackage $excel -WorksheetName $worksheetName -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ErrorAction Stop
                        
                        # Vérification et application du formatage conditionnel
                        $ws = $excel.Workbook.Worksheets[$worksheetName]
                        if ($ws -and $ws.Dimension) {
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText -ConditionValue "File" -BackgroundColor LightBlue -Column 3
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText -ConditionValue "Process" -BackgroundColor LightGreen -Column 3
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText -ConditionValue "Network" -BackgroundColor LightYellow -Column 3
                        }
                        
                        Write-Host "  $($results.Count) détections trouvées"
                    }
                    catch {
                        Write-Warning "Erreur lors de l'export Excel pour $originalName : $_"
                        # Création d'une feuille avec message d'erreur
                        try {
                            $errorWorksheet = Add-Worksheet -ExcelPackage $excel -WorksheetName $worksheetName -ErrorAction Stop
                            Set-ExcelRange -Worksheet $errorWorksheet -Range "A1" -Value "Erreur lors de l'export des résultats" -Bold
                        }
                        catch {
                            Write-Warning "Impossible de créer la feuille d'erreur : $_"
                        }
                    }
                } 
                else {
                    # Création d'une feuille pour les résultats vides
                    try {
                        $emptyWorksheet = Add-Worksheet -ExcelPackage $excel -WorksheetName $worksheetName -ErrorAction Stop
                        Set-ExcelRange -Worksheet $emptyWorksheet -Range "A1" -Value "Aucune détection trouvée" -Bold
                        Write-Host "  Aucune détection trouvée"
                    }
                    catch {
                        Write-Warning "Erreur lors de la création de la feuille vide pour $originalName : $_"
                    }
                }
            }
            catch {
                Write-Warning "Erreur lors de l'exécution de la requête pour $originalName : $_"
                continue
            }
        } # Fin de la boucle foreach

        # Création de la feuille de résumé
        # Initialisation du fichier Excel et de la collection de résumé
        Write-Host "Création du fichier Excel..."
        $excelPath = Join-Path $sessionExportPath "DefenderHunting_Results_$timestamp.xlsx"
        $summaryResults = [System.Collections.ArrayList]::new()
        $excelCreated = $false
        
        # Traitement de chaque famille
        foreach ($safeName in $hashesByFamily.Keys) {
            $originalName = $hashesByFamily[$safeName].OriginalName
            Write-Host "`nTraitement de la famille: $originalName"
            
            # Génération et exécution de la requête
            try {
                $query = # ... votre code de génération de requête existant ...
                Write-Host "Exécution de la requête pour $originalName..."
                $results = Invoke-DefenderQuery -Query $query -Token $token -FamilyName $originalName
                
                # Préparation des données de résumé
                $summaryEntry = [PSCustomObject]@{
                    'Famille' = $originalName
                    'Nombre de détections' = 0
                    'SHA256 uniques' = $hashesByFamily[$safeName].SHA256.Count
                    'SHA1 uniques' = $hashesByFamily[$safeName].SHA1.Count
                    'MD5 uniques' = $hashesByFamily[$safeName].MD5.Count
                    'Date analyse' = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                }
                
                if ($results -and $results.Count -gt 0) {
                    # Création du fichier Excel si c'est la première fois qu'on trouve des résultats
                    if (-not $excelCreated) {
                        $excel = Open-ExcelPackage -Path $excelPath -Create
                        $excelCreated = $true
                    }
                    
                    try {
                        # Préparation du nom de feuille
                        $worksheetName = if ($safeName.Length -gt 31) {
                            $safeName.Substring(0, 28) + "..."
                        } else {
                            $safeName
                        }
                        
                        # Export des résultats
                        $results | Export-Excel -ExcelPackage $excel -WorksheetName $worksheetName `
                                 -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
                        
                        # Application du formatage conditionnel
                        $ws = $excel.Workbook.Worksheets[$worksheetName]
                        if ($ws) {
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText `
                                                    -ConditionValue "File" -BackgroundColor LightBlue -Column 3
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText `
                                                    -ConditionValue "Process" -BackgroundColor LightGreen -Column 3
                            Add-ConditionalFormatting -WorkSheet $ws -RuleType ContainsText `
                                                    -ConditionValue "Network" -BackgroundColor LightYellow -Column 3
                        }
                        
                        $summaryEntry.'Nombre de détections' = $results.Count
                        Write-Host "  $($results.Count) détections trouvées"
                    }
                    catch {
                        Write-Warning "Erreur lors de l'export des résultats pour $originalName : $_"
                    }
                } else {
                    Write-Host "  Aucune détection trouvée"
                }
                
                [void]$summaryResults.Add($summaryEntry)
            }
            catch {
                Write-Warning "Erreur lors du traitement de $originalName : $_"
                continue
            }
        }
        
        # Création de la feuille de résumé seulement si des résultats ont été trouvés
        if ($excelCreated) {
            try {
                $summaryResults | Export-Excel -ExcelPackage $excel -WorksheetName 'Résumé' `
                                -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
                Write-Host "`nFeuille de résumé créée avec succès"
                
                # Sauvegarde et fermeture du fichier Excel
                Close-ExcelPackage $excel -Show
                Write-Host "Rapport Excel généré : $excelPath"
            }
            catch {
                Write-Warning "Erreur lors de la finalisation du fichier Excel : $_"
            }
        } else {
            Write-Host "`nAucune détection trouvée pour toutes les familles - Pas de fichier Excel généré"
        }
    }
    catch {
        Write-Error "Erreur lors de l'analyse : $_"
    }

} # Fin de la fonction Start-DefenderHunting

# Export de la fonction principale
Export-ModuleMember -Function Start-DefenderHunting