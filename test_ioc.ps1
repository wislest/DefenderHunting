# Test avec des données manquantes
$testCsv = @"
FileHashes,FamilyName,FullPath (REDACTED)
{"sha256":"abc123"},,
{"sha256":"def456"},"TestFamily",
"","EmptyHash","C:\test"
"@ | ConvertFrom-Csv

# Sauvegarder dans un fichier temporaire
$testPath = ".\test_iocs.csv"
$testCsv | Export-Csv $testPath -NoTypeInformation