# Test-DefenderConfig.ps1
param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigPath
)

try {
    # Lire la configuration
    $config = Get-Content $ConfigPath | ConvertFrom-Json
    
    # Vérifier les champs requis
    $requiredFields = @('TenantId', 'ClientId', 'ClientSecret')
    foreach ($field in $requiredFields) {
        if (-not $config.$field) {
            Write-Error "Champ manquant dans la configuration: $field"
            return
        }
    }
    
    # Tester l'authentification
    $tokenUrl = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"
    $body = @{
        client_id     = $config.ClientId
        client_secret = $config.ClientSecret
        scope         = "https://api.securitycenter.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    
    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
    
    if ($response.access_token) {
        Write-Host "Configuration valide! Token obtenu avec succès."
        Write-Host "Vous pouvez maintenant utiliser le script IoC-Hunter.ps1"
    }
}
catch {
    Write-Error "Erreur lors du test de la configuration: $_"
}