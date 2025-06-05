# Setup-DefenderHunter.ps1
param(
    [Parameter(Mandatory=$false)]
    [string]$BasePath = "C:\DefenderHunting"
)

# Créer les dossiers
$folders = @(
    "Modules\DefenderHunter",
    "Queries",
    "Exports",
    "Logs"
)

foreach ($folder in $folders) {
    New-Item -ItemType Directory -Path (Join-Path $BasePath $folder) -Force
}

# Créer le module
$ModulePath = Join-Path $BasePath "Modules\DefenderHunter"

# Créer le manifest
$manifestParams = @{
    Path              = "$ModulePath\DefenderHunter.psd1"
    RootModule       = "DefenderHunter.psm1"
    Author           = "Wiston Lestin"
    Description      = "Module pour la recherche d'IoCs dans Microsoft Defender"
    ModuleVersion    = "1.0.0"
}
New-ModuleManifest @manifestParams

# Créer le module principal
@"
# Importer les fonctions
. `$PSScriptRoot\DefenderExport-Core.ps1
. `$PSScriptRoot\DefenderExport-Queries.ps1
. `$PSScriptRoot\DefenderExport-Functions.ps1
. `$PSScriptRoot\DefenderExport-Main.ps1

# Exporter les fonctions publiques
Export-ModuleMember -Function Start-DefenderHunting, Import-IoCs, Export-Results
"@ | Out-File "$ModulePath\DefenderHunter.psm1"

# Copier les scripts existants
Copy-Item "Scripts\*.ps1" -Destination $ModulePath

Write-Host "Configuration terminée. Pour utiliser le module:"
Write-Host "1. Import-Module $ModulePath\DefenderHunter.psd1"
Write-Host "2. Start-DefenderHunting -IocsFile 'chemin\vers\iocs.csv'"