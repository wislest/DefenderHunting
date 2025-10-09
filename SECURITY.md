# Sécurité et Gestion des Identifiants

## Configuration des Identifiants

⚠️ **Ne commitez JAMAIS vos vraies credentials sur Git!**

### Option 1: Fichier config.json (Par défaut)

1. Copiez le fichier d'exemple:
   ```powershell
   Copy-Item config.json.example config.json
   ```

2. Éditez `config.json` avec vos vraies valeurs Azure AD

3. Vérifiez que `config.json` est dans `.gitignore` ✓

### Option 2: Variables d'environnement (Recommandé pour CI/CD)

1. Copiez le fichier d'exemple:
   ```powershell
   Copy-Item .env.example .env
   ```

2. Éditez `.env` avec vos valeurs

3. Chargez les variables avant d'exécuter:
   ```powershell
   # Charger les variables d'environnement depuis .env
   Get-Content .env | ForEach-Object {
       if ($_ -match '^\s*([^#][^=]+)=(.+)$') {
           $name = $matches[1].Trim()
           $value = $matches[2].Trim()
           [Environment]::SetEnvironmentVariable($name, $value, 'Process')
       }
   }
   ```

### Option 3: Azure Key Vault (Production)

Pour les environnements de production, utilisez Azure Key Vault:

```powershell
# Récupérer les secrets depuis Key Vault
$TenantId = (Get-AzKeyVaultSecret -VaultName "YourVault" -Name "DefenderTenantId").SecretValueText
$ClientId = (Get-AzKeyVaultSecret -VaultName "YourVault" -Name "DefenderClientId").SecretValueText
$ClientSecret = (Get-AzKeyVaultSecret -VaultName "YourVault" -Name "DefenderClientSecret").SecretValueText

# Exécuter
Start-DefenderHunting -IocsFile ".\iocs.csv" -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
```

## Fichiers Protégés par .gitignore

Les fichiers suivants ne seront **jamais** committés:
- `config.json` - Configuration avec credentials
- `.env` - Variables d'environnement
- `Logs/` - Fichiers de logs
- `Exports/` - Résultats d'analyse

## Rotation des Secrets

Changez régulièrement vos secrets Azure AD:
1. Azure Portal → App registrations → Votre app → Certificates & secrets
2. Créez un nouveau secret
3. Mettez à jour `config.json` ou `.env`
4. Supprimez l'ancien secret

## Permissions Requises

L'application Azure AD doit avoir:
- **API**: Microsoft Threat Protection
- **Permission**: `AdvancedHunting.Read.All`
- **Type**: Application (pas délégué)
