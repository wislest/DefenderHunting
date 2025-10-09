#Requires -Version 5.1
<#
.SYNOPSIS
    Convertit un fichier CSV simple (FamilyName,Hash) au format attendu par DefenderHunter

.DESCRIPTION
    Transforme un CSV avec seulement FamilyName et FileHashes (hash simple)
    vers le format complet: FullPath,FamilyName,FileHashes (JSON avec md5/sha1/sha256)

.PARAMETER InputFile
    Fichier CSV source avec colonnes: FamilyName, FileHashes (hash simple)

.PARAMETER OutputFile
    Fichier CSV de sortie au format DefenderHunter

.EXAMPLE
    .\Convert-IocsFormat.ps1 -InputFile "edrkiller-iocs.csv" -OutputFile "edrkiller-iocs-formatted.csv"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [Parameter(Mandatory=$true)]
    [string]$OutputFile
)

# Vérifier que le fichier existe
if (-not (Test-Path $InputFile)) {
    Write-Error "Fichier introuvable: $InputFile"
    exit 1
}

Write-Host "Lecture de $InputFile..." -ForegroundColor Cyan

# Lire le CSV source
$iocs = Import-Csv $InputFile

# Préparer les données converties
$convertedIocs = @()

foreach ($ioc in $iocs) {
    $familyName = $ioc.FamilyName
    $hash = $ioc.FileHashes

    # Ignorer les lignes sans hash valide (URLs, en-têtes, etc.)
    if ([string]::IsNullOrWhiteSpace($hash) -or $hash -notmatch '^[a-fA-F0-9]+$') {
        Write-Warning "Ligne ignorée (hash invalide): FamilyName='$familyName', Hash='$hash'"
        continue
    }

    # Déterminer le type de hash basé sur la longueur
    $hashType = switch ($hash.Length) {
        32  { "md5" }
        40  { "sha1" }
        64  { "sha256" }
        default {
            Write-Warning "Hash de longueur inconnue ($($hash.Length)): $hash - Supposé SHA256"
            "sha256"
        }
    }

    # Créer l'objet JSON pour FileHashes
    $hashesJson = @{
        md5 = ""
        sha1 = ""
        sha256 = ""
    }

    $hashesJson[$hashType] = $hash.ToLower()

    # Convertir en JSON - PowerShell Export-Csv gérera l'échappement automatiquement
    $jsonString = ($hashesJson | ConvertTo-Json -Compress)

    # Créer l'entrée formatée
    $convertedIocs += [PSCustomObject]@{
        FullPath = "Unknown\$($familyName.Replace(' ', '_')).exe"
        FamilyName = $familyName
        FileHashes = $jsonString
    }
}

Write-Host "Conversion terminée: $($convertedIocs.Count) IoCs valides sur $($iocs.Count) lignes" -ForegroundColor Green

# Exporter au format CSV
$convertedIocs | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "Fichier créé: $OutputFile" -ForegroundColor Green
Write-Host "`nVous pouvez maintenant exécuter:" -ForegroundColor Yellow
Write-Host "Start-DefenderHunting -IocsFile `"$OutputFile`" -ConfigPath `".\config.json`" -Verbose" -ForegroundColor Cyan
