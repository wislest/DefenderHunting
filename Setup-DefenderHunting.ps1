# Setup-DefenderHunting.ps1

# Créer la structure des dossiers
$basePath = "C:\DefenderHunting"
$folders = @(
    "Scripts",
    "Queries",
    "Exports",
    "Logs"
)

# Créer les dossiers
foreach ($folder in $folders) {
    $path = Join-Path $basePath $folder
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force
    }
}

# Copier les fichiers de requêtes
$queries = @{
    "FileEvents.kql" = $fileEventsQuery
    "ProcessEvents.kql" = $processEventsQuery
    "NetworkEvents.kql" = $networkEventsQuery
}

foreach ($query in $queries.GetEnumerator()) {
    $path = Join-Path $basePath "Queries\$($query.Key)"
    $query.Value | Out-File -FilePath $path -Encoding UTF8
}

# Créer le fichier de configuration
$config = @{
    TenantId = ""
    ClientId = ""
    ClientSecret = ""
    ExportPath = Join-Path $basePath "Exports"
    LogPath = Join-Path $basePath "Logs"
} | ConvertTo-Json

$config | Out-File -FilePath (Join-Path $basePath "config.json") -Encoding UTF8