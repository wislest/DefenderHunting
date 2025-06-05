# Create-DefenderModule.ps1
$modulePath = "C:\DefenderHunting\Modules\DefenderHunter"
New-Item -Type Directory -Path $modulePath -Force

# Créer le manifeste du module
$manifestParams = @{
    Path = "$modulePath\DefenderHunter.psd1"
    ModuleVersion = "1.0.0"
    Author = "Wiston Lestin - Security Team"
    Description = "Module pour la chasse aux IoCs dans Microsoft Defender"
    RootModule = "DefenderHunter.psm1"
}
New-ModuleManifest @manifestParams