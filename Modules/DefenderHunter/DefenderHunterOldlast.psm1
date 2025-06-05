# DefenderHunter.psm1
# Module pour l'analyse des IoCs via Microsoft Defender Advanced Hunting
# Auteur : Wiston Lestin

# Fonction : Connexion à l'API Defender et récupération du token d'accès
function Connect-DefenderAPI {
    param(
        [string]$ConfigFile = "C:\DefenderHunting\config.json"
    )

    Write-Verbose "Lecture de la configuration API depuis $ConfigFile..."

    # Vérifier que le fichier de configuration existe
    if (-not (Test-Path -Path $ConfigFile)) {
        throw "Le fichier de configuration $ConfigFile est introuvable. Vérifiez le chemin."
    }

    # Charger les informations d’authentification
    $config = Get-Content -Path $ConfigFile | ConvertFrom-Json

    $tenantId = $config.TenantID
    $clientId = $config.ClientID
    $clientSecret = $config.ClientSecret

    # Vérification des paramètres
    if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
        throw "Le fichier de configuration doit contenir TenantID, ClientID et ClientSecret."
    }

    Write-Verbose "Tentative de connexion à Microsoft Defender API..."
    
    # URL de connexion au token
    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Corps de la requête pour récupérer le token
    $body = @{
        client_id     = $clientId
        scope         = "https://api.security.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    try {
        # Appel à l'API pour récupérer le token
        $response = Invoke-RestMethod -Method Post -Uri $authUrl -Body $body -ContentType "application/x-www-form-urlencoded"

        if ($response.access_token) {
            Write-Verbose "Connexion réussie. Token récupéré avec succès."
            return $response.access_token
        } else {
            throw "Token non reçu. Vérifiez vos informations d'authentification."
        }
    } catch {
        throw "Erreur lors de la connexion à l'API Defender : $_"
    }
}

# Fonction : Importation et traitement des IoCs depuis un fichier CSV
function Import-IoCs {
    param(
        [string]$IocsFile
    )

    Write-Verbose "Importation des IoCs depuis le fichier : $IocsFile..."
    $IoCs = Import-Csv $IocsFile
    foreach ($ioc in $IoCs) {
        try {
            # Extraire les hashes depuis JSON
            $hashes = $ioc.FileHashes | ConvertFrom-Json
            $ioc | Add-Member -MemberType NoteProperty -Name "MD5" -Value $hashes.md5
            $ioc | Add-Member -MemberType NoteProperty -Name "SHA1" -Value $hashes.sha1
            $ioc | Add-Member -MemberType NoteProperty -Name "SHA256" -Value $hashes.sha256
        } catch {
            Write-Warning "Erreur lors de l'analyse des IoCs pour $($ioc.FullPath): $_"
        }
    }
    return $IoCs
}

# Fonction : Générer une requête KQL pour une famille IoC
function New-KqlQuery {
    param(
        [string]$FamilyName,
        [array]$IoCs,
        [string]$LookbackDays = "30d"
    )
    Write-Verbose "Création d'une nouvelle requête KQL pour la famille : $FamilyName..."

    $hashFilters = @()
    foreach ($ioc in $IoCs) {
        if ($ioc.SHA256) {
            $hashFilters += "SHA256 == '$($ioc.SHA256)'"
        }
    }

    if (-not $hashFilters) {
        Write-Warning "Aucun SHA256 trouvé pour la famille $FamilyName."
        return $null
    }

    $filterClause = $hashFilters -join " or "
    $kqlQuery = @"
DeviceFileEvents
| where Timestamp > ago($LookbackDays)
| where $filterClause
| project Timestamp, DeviceName, FileName, FolderPath, SHA256, InitiatingProcessCommandLine
"@
    return $kqlQuery
}

# Fonction : Envoi de la requête via l'API Defender
function Invoke-DefenderApi {
    param(
        [string]$KqlQuery,
        [string]$ApiToken
    )

    Write-Verbose "Envoi de la requête KQL via l'API Defender Advanced Hunting..."

    # Construisez la requête KQL
$kqlQuery = @{
    Query = @"
        DeviceEvents
        | summarize count() by DeviceName
"@
}
    
    try {
        # Exécuter la requête via l'API
        $response = Invoke-RestMethod -Uri "https://api-us.securitycenter.microsoft.com/api/advancedhunting/run" `
                                            -Headers @{ "Authorization" = "Bearer $ApiToken" } `
                                            -Method POST `
                                            -Body (@{ Query = $KqlQuery } | ConvertTo-Json -Depth 10) `
                                            -ContentType "application/json"
        return $response
        Write-Output "Token: $response.access_token"

    } catch {
        Write-Warning "Erreur API : $($_.Exception.Message)"
        if ($null -ne $_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorDetails = $reader.ReadToEnd()
            Write-Warning "Détails de l'erreur : $errorDetails"
        }
        return $null
    }
}


<#
.SYNOPSIS
Démarre l'analyse des IoCs avec Microsoft Defender Advanced Hunting.

.DESCRIPTION
Cette fonction importe une liste d'IoCs depuis un fichier CSV, se connecte à l'API de Defender pour exécuter des requêtes KQL générées pour chaque IoC,
et exporte les résultats dans un fichier CSV.

.PARAMETER IocsFile
Chemin du fichier CSV contenant les IoCs à analyser.

.PARAMETER ConfigFile
Chemin du fichier de configuration contenant TenantID, ClientID et ClientSecret pour se connecter à l'API.

.PARAMETER LookbackDays
Période (en jours) sur laquelle exécuter les recherches. Défaut : 30 jours.

.PARAMETER OutputFile
Chemin où exporter les résultats de l'analyse. Défaut : ***REMOVED***\HuntingResults.csv.

.PARAMETER BatchSize
Taille des lots d'IoCs traités par requête pour éviter des erreurs API. Défaut : 50.

.EXAMPLE
Start-DefenderHunting -IocsFile "C:\DefenderHunting\IoCs.csv" -ConfigFile "C:\DefenderHunting\config.json" -LookbackDays 30

.EXAMPLE
Start-DefenderHunting -IocsFile "IoCs.csv" -ConfigFile "config.json" -OutputFile "results.csv" -BatchSize 100 -Verbose

#>

# Fonction principale : Démarrage de l'analyse des IoCs
function Start-DefenderHunting {
    param(
        [string]$IocsFile,
        [int]$LookbackDays = 30,
        [string]$ConfigFile = "C:\DefenderHunting\config.json",
        [string]$OutputFile = "***REMOVED***\HuntingResults.csv",
        [int]$BatchSize = 50
    )

    Write-Verbose "Démarrage de l'analyse des IoCs depuis $IocsFile..."

    # Obtenir le token via Connect-DefenderAPI
    $apiToken = Connect-DefenderAPI -ConfigFile "C:\DefenderHunting\config.json"
    Write-Output "Token: $apiToken"

    if (-not $apiToken) {
        throw "Impossible de récupérer un token d'authentification. Vérifiez vos informations."
    }

    Write-Verbose "Token récupéré avec succès. Début du traitement des IoCs..."

    # Charger les IoCs
    $IoCs = Import-IoCs -IocsFile $IocsFile
    if (-not $IoCs) {
        Write-Warning "Aucun IoC trouvé dans le fichier fourni."
        return
    }

    # Grouper les IoCs par famille
    $families = $IoCs | Group-Object -Property FamilyName
    $allResults = @()

    foreach ($family in $families) {
        Write-Verbose "Traitement de la famille : $($family.Name)"
        $familyIoCs = $family.Group

        # Diviser les IoCs en lots pour éviter des requêtes trop volumineuses
        $batches = @()
        for ($i = 0; $i -lt $familyIoCs.Count; $i += $BatchSize) {
            $batches += ,$familyIoCs[$i..([math]::Min($i + $BatchSize - 1, $familyIoCs.Count - 1))]
        }

        foreach ($batch in $batches) {
            # Générer et exécuter la requête KQL
            $kqlQuery = New-KqlQuery -FamilyName $family.Name -IoCs $batch -LookbackDays "$LookbackDays"d
            if ($kqlQuery) {
                $results = Invoke-DefenderApi -KqlQuery $kqlQuery -ApiToken $apiToken
                if ($results) {
                    $allResults += $results
                }
            }
        }
    }

    # Exporter les résultats
    if ($allResults) {
        Write-Verbose "Exportation des résultats vers $OutputFile..."
        $allResults | Export-Csv -Path $OutputFile -NoTypeInformation -Force
        Write-Verbose "Export terminé."
    } else {
        Write-Warning "Aucun résultat à exporter."
    }
}

Export-ModuleMember -Function Start-DefenderHunting, Connect-DefenderAPI, Import-IoCs, New-KqlQuery, Invoke-DefenderApi
