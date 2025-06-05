# C:\DefenderHunting\Setup-Environment.ps1
# Créer la structure des dossiers
$basePath = $PWD
$folders = @(
    "Scripts",
    "Queries",
    "Exports",
    "Logs"
)

foreach ($folder in $folders) {
    $path = Join-Path $basePath $folder
    if (!(Test-Path $path)) {
        New-Item -ItemType Directory -Path $path
    }
}

# Copier les fichiers de script si nécessaire
Copy-Item .\DefenderExport-Main.ps1 .\Scripts\ -Force