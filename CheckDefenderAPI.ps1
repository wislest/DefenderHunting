# Script de test mis à jour avec les paramètres de Logic Apps
$config = Get-Content "C:\DefenderHunting\config.json" | ConvertFrom-Json

Write-Host "Tentative d'obtention du token..."
$tokenUrl = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"

# Mise à jour avec les valeurs correspondant à Logic Apps
$body = @{
    client_id = $config.ClientId
    client_secret = $config.ClientSecret
    # Utilisation de l'audience correcte vue dans Logic Apps
    scope = "https://api.securitycenter.microsoft.com/.default"
    grant_type = "client_credentials"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -Verbose
    Write-Host "Token obtenu avec succès"
    
    $headers = @{
        'Content-Type' = 'application/json'
        'Accept' = 'application/json'
        'Authorization' = "Bearer $($tokenResponse.access_token)"
    }

    # Utilisation de l'URL vue dans Logic Apps
    $apiUrl = "https://api.securitycenter.microsoft.com/api/advancedqueries/run"

    # Requête KQL simple
    $query = @{
        'Query' = "DeviceInfo | limit 1"
    } | ConvertTo-Json

    Write-Host "Envoi de la requête à l'API..."
    Write-Host "URL utilisée: $apiUrl"
    
    $response = Invoke-RestMethod -Uri $apiUrl `
        -Method Post `
        -Headers $headers `
        -Body $query `
        -Verbose

    Write-Host "Réponse reçue avec succès :"
    $response.Results | Format-Table
} catch {
    Write-Host "Erreur détaillée :"
    Write-Host "StatusCode: $($_.Exception.Response.StatusCode.value__)"
    Write-Host "StatusDescription: $($_.Exception.Response.StatusDescription)"
    
    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd()
    Write-Host "Corps de la réponse : $responseBody"
}