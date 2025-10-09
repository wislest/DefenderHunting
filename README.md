# DefenderHunting

## Description
Solution automatisée pour l'analyse des Indicateurs de Compromission (IoCs) via Microsoft Defender Advanced Hunting.

## Démarrage Rapide

### 1. Configuration des Identifiants

⚠️ **Important**: Ne commitez JAMAIS vos credentials! Consultez [SECURITY.md](SECURITY.md) pour les meilleures pratiques.

**Option A - Fichier config.json (Recommandé)**
```powershell
# Copier le template
Copy-Item config.json.example config.json

# Éditer config.json avec vos vraies valeurs Azure AD
# TenantId, ClientId, ClientSecret
```

**Option B - Variables d'environnement**
```powershell
Copy-Item .env.example .env
# Éditer .env puis charger les variables
Get-Content .env | ForEach-Object {
    if ($_ -match '^\s*([^#][^=]+)=(.+)$') {
        [Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), 'Process')
    }
}
```

### 2. Préparer le Fichier IoCs

**Format requis**: CSV avec 3 colonnes obligatoires:
- `FullPath`: Chemin du fichier (peut être "Unknown" si inconnu)
- `FamilyName`: Nom de la famille de malware
- `FileHashes`: Objet JSON avec `md5`, `sha1`, `sha256`

**Exemple de format correct**:
```csv
FullPath,FamilyName,FileHashes
C:\temp\malware.exe,Akira,"{""md5"":""176b7a50bbcbc34acedb46561a04c3d0"",""sha1"":""38a29d7b1345d3221a550b5b5909436a451a53ca"",""sha256"":""c5bcfd00d0b8fda7c4b20cdc9649713d9f01dd12f61ce8ee9c45ec424a6bbdf2""}"
```

#### Conversion de Fichiers Simplifiés

Si vous avez un CSV simple avec seulement `FamilyName` et un hash SHA256:

```csv
FamilyName,FileHashes
EDR killer,f51397bb18e166c933fe090320ec23397fed73b68157ce86406db9f07847d355
```

Utilisez le script de conversion:
```powershell
.\Convert-IocsFormat.ps1 -InputFile "votre-fichier.csv" -OutputFile "votre-fichier-formatted.csv"
```

Le script détecte automatiquement le type de hash (MD5/SHA1/SHA256) selon sa longueur et génère le format JSON requis.

### 3. Exécuter l'Analyse

```powershell
# Importer le module
Import-Module .\Modules\DefenderHunter\DefenderHunter.psd1 -Force

# Lancer la chasse aux menaces
Start-DefenderHunting -IocsFile ".\votre-fichier-formatted.csv" -ConfigPath ".\config.json" -Verbose
```

### 4. Résultats

Les résultats sont générés dans:
- **Excel**: `Exports\DefenderHunting_Results_{timestamp}.xlsx`
- **Requêtes KQL**: `Queries\{FamilyName}.kql`
- **Logs**: `Logs\`

## Fichiers Disponibles

- `cover-iocs.csv` - Exemple complet au bon format
- `test_iocs.csv` - Fichier de test minimal
- `config.json.example` - Template de configuration
- `.env.example` - Template pour variables d'environnement
- `Convert-IocsFormat.ps1` - Convertisseur de format CSV

## Permissions Azure AD Requises

Votre application Azure AD doit avoir:
- **API**: Microsoft Threat Protection
- **Permission**: `AdvancedHunting.Read.All` (Application)
- **Grant admin consent**: Oui

## Dépannage

### Erreur: "Conversion from JSON failed"
→ Format CSV incorrect. Utilisez `Convert-IocsFormat.ps1` pour convertir votre fichier.

### Erreur: "Specified tenant identifier is neither a valid DNS name"
→ TenantId invalide dans config.json. Vérifiez vos credentials Azure AD.

### Erreur: "Fichier IoCs introuvable"
→ Vérifiez le nom du fichier et le chemin. Utilisez `Get-ChildItem *.csv` pour lister les fichiers disponibles.

## Documentation Complète

- [CLAUDE.md](CLAUDE.md) - Guide technique détaillé
- [SECURITY.md](SECURITY.md) - Gestion sécurisée des identifiants 
