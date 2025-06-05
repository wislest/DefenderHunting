# Install-DefenderHunter.ps1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$BasePath = "C:\DefenderHunting"
)

# 1. Créer la structure de dossiers
$folders = @(
    "Modules\DefenderHunter",
    "Queries",
    "Exports",
    "Logs"
)

foreach ($folder in $folders) {
    $path = Join-Path $BasePath $folder
    if (!(Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
        Write-Host "Créé: $path"
    }
}

# 2. Créer/Mettre à jour le fichier module principal
$modulePath = Join-Path $BasePath "Modules\DefenderHunter"
$moduleContent = Get-Content -Path "$PSScriptRoot\DefenderHunter.psm1" -Raw
$moduleContent | Out-File -FilePath (Join-Path $modulePath "DefenderHunter.psm1") -Force -Encoding UTF8

# 3. Créer/Mettre à jour le manifeste
$manifestPath = Join-Path $modulePath "DefenderHunter.psd1"
if (Test-Path $manifestPath) {
    Write-Host "Mise à jour du manifeste existant"
    Update-ModuleManifest -Path $manifestPath -ModuleVersion "1.0.1"
} else {
    Write-Host "Création d'un nouveau manifeste"
    New-ModuleManifest -Path $manifestPath `
        -ModuleVersion "1.0.0" `
        -Author "Wiston Lestin - Security Team" `
        -RootModule "DefenderHunter.psm1" `
        -Description "Module pour la chasse aux IoCs dans Microsoft Defender"
}

Write-Host "`nInstallation terminée. Pour utiliser le module:"
Write-Host "1. Import-Module $modulePath -Force"
Write-Host "2. Start-DefenderHunting -IocsFile 'chemin\vers\iocs.csv'"