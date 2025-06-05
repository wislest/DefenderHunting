# Script PowerShell simplifié pour créer un runbook Excel d'attachement de ressources aux DCRs
# Pour la migration de Splunk vers Microsoft Sentinel - Version COM compatible

# Définir le chemin du répertoire et du fichier de sortie
$baseDirectory = "C:\Users\lestinw\log-ingestion"
$outputFile = Join-Path $baseDirectory "Migration_Splunk_Sentinel_Runbook.xlsx"

# Lire les fichiers CSV pour obtenir les listes de serveurs
$csvFiles = @{
    'AzureMySqlServer' = Join-Path $baseDirectory "AzureMySqlServer.csv"
    'AzurePostgresSqlServer' = Join-Path $baseDirectory "AzurePostgresSqlServer.csv"
    'AzuresqlServerInstances' = Join-Path $baseDirectory "AzuresqlServerInstances.csv"
    'DC_Servers_Sensors' = Join-Path $baseDirectory "DC_Servers_Sensors.csv"
    'Linux_servers_azure-arc' = Join-Path $baseDirectory "Linux_servers_azure-arc.csv"
    'Windows_servers_azure-arc' = Join-Path $baseDirectory "Windows_servers_azure-arc.csv"
}

# Vérifier que tous les fichiers existent
$missingFiles = $csvFiles.GetEnumerator() | Where-Object { -not (Test-Path $_.Value) } | ForEach-Object { $_.Value }
if ($missingFiles) {
    Write-Error "Fichiers manquants: $($missingFiles -join ', ')"
    exit
}

# Lire les fichiers CSV
try {
    $dfMysql = Import-Csv $csvFiles['AzureMySqlServer'] -Encoding UTF8
    $dfPostgres = Import-Csv $csvFiles['AzurePostgresSqlServer'] -Encoding UTF8
    $dfSqlserver = Import-Csv $csvFiles['AzuresqlServerInstances'] -Encoding UTF8
    $dfDc = Import-Csv $csvFiles['DC_Servers_Sensors'] -Encoding UTF8
    $dfLinux = Import-Csv $csvFiles['Linux_servers_azure-arc'] -Encoding UTF8
    $dfWindows = Import-Csv $csvFiles['Windows_servers_azure-arc'] -Encoding UTF8
    
    Write-Host "Tous les fichiers CSV ont été chargés avec succès"
    
    # Afficher les dimensions des DataFrames pour vérification
    Write-Host "MySQL servers: $($dfMysql.Count) lignes"
    Write-Host "PostgreSQL servers: $($dfPostgres.Count) lignes"
    Write-Host "SQL Server instances: $($dfSqlserver.Count) lignes"
    Write-Host "Domain Controllers: $($dfDc.Count) lignes"
    Write-Host "Linux servers: $($dfLinux.Count) lignes"
    Write-Host "Windows servers: $($dfWindows.Count) lignes"
    
} catch {
    Write-Error "Erreur lors de la lecture des fichiers CSV: $_"
    exit
}

# Créer un objet Excel via COM
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()

    # Fonction pour créer la feuille résumé
    function Create-SummarySheet {
        $summarySheet = $workbook.Worksheets.Item(1)
        $summarySheet.Name = "Résumé"
        
        # Titre principal
        $summarySheet.Cells.Item(1, 1).Value = "Runbook - Test d'ingestion Splunk vers Microsoft Sentinel"
        $summarySheet.Cells.Item(1, 1).Font.Bold = $true
        $summarySheet.Cells.Item(1, 1).Font.Size = 16
        $summarySheet.Range("A1:H1").Merge() | Out-Null
        $summarySheet.Cells.Item(1, 1).HorizontalAlignment = -4108 # xlCenter
        
        # Date et informations du projet
        $summarySheet.Cells.Item(3, 1).Value = "Date du test d'ingestion:"
        $summarySheet.Cells.Item(3, 3).Value = (Get-Date).ToString("dd/MM/yyyy")
        $summarySheet.Cells.Item(3, 1).Font.Bold = $true
        $summarySheet.Cells.Item(3, 5).Value = "Environnement:"
        $summarySheet.Cells.Item(3, 7).Value = "Production"
        $summarySheet.Cells.Item(3, 5).Font.Bold = $true
        
        $summarySheet.Cells.Item(4, 1).Value = "Responsable:"
        $summarySheet.Cells.Item(4, 3).Value = "[Nom du responsable]"
        $summarySheet.Cells.Item(4, 1).Font.Bold = $true
        $summarySheet.Cells.Item(4, 5).Value = "Durée prévue:"
        $summarySheet.Cells.Item(4, 7).Value = "1 journée"
        $summarySheet.Cells.Item(4, 5).Font.Bold = $true
        
        # Résumé de l'inventaire
        $summarySheet.Cells.Item(6, 1).Value = "Résumé des ressources à attacher aux DCRs"
        $summarySheet.Cells.Item(6, 1).Font.Bold = $true
        $summarySheet.Range("A6:H6").Merge() | Out-Null
        $summarySheet.Cells.Item(6, 1).HorizontalAlignment = -4108 # xlCenter
        $summarySheet.Cells.Item(6, 1).Interior.ColorIndex = 34 # Light blue
        
        # En-têtes du tableau de résumé
        $headers = @("Type de serveur", "Nombre de serveurs", "DCR associée", "Statut", "Progression", "Temps estimé", "Temps réel")
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $summarySheet.Cells.Item(7, $i+1).Value = $headers[$i]
            $summarySheet.Cells.Item(7, $i+1).Font.Bold = $true
            $summarySheet.Cells.Item(7, $i+1).Interior.ColorIndex = 5 # Blue
            $summarySheet.Cells.Item(7, $i+1).Font.ColorIndex = 2 # White
            $summarySheet.Cells.Item(7, $i+1).HorizontalAlignment = -4108 # xlCenter
            $summarySheet.Cells.Item(7, $i+1).Borders.LineStyle = 1 # xlContinuous
        }
        
        # Données du tableau de résumé
        $serverTypes = @(
            @("Contrôleurs de domaine", $dfDc.Count, "DCR-DomainController"),
            @("Serveurs Windows", $dfWindows.Count, "DCR-WindowsServer"),
            @("Serveurs Linux", $dfLinux.Count, "DCR-LinuxServer"),
            @("SQL Server", $dfSqlserver.Count, "DCR-SQLServer"),
            @("MySQL Server", $dfMysql.Count, "DCR-MySQLServer"),
            @("PostgreSQL Server", $dfPostgres.Count, "DCR-PostgresServer")
        )
        
        for ($i = 0; $i -lt $serverTypes.Count; $i++) {
            $row = $i + 8
            $serverType, $count, $dcr = $serverTypes[$i]
            
            $summarySheet.Cells.Item($row, 1).Value = $serverType
            $summarySheet.Cells.Item($row, 2).Value = $count
            $summarySheet.Cells.Item($row, 3).Value = $dcr
            $summarySheet.Cells.Item($row, 4).Value = "Non commencé"
            $summarySheet.Cells.Item($row, 5).Value = "0%"
            $summarySheet.Cells.Item($row, 6).Value = [math]::Max(1, [math]::Ceiling($count / 60)) + " heures"
            
            # Appliquer les styles
            for ($col = 1; $col -le 7; $col++) {
                $summarySheet.Cells.Item($row, $col).Borders.LineStyle = 1 # xlContinuous
                if ($col -in 2, 5) {
                    $summarySheet.Cells.Item($row, $col).HorizontalAlignment = -4108 # xlCenter
                }
            }
        }
        
        # Ajouter une formule pour calculer le total
        $totalRow = $serverTypes.Count + 8
        $summarySheet.Cells.Item($totalRow, 1).Value = "Total"
        $summarySheet.Cells.Item($totalRow, 1).Font.Bold = $true
        $summarySheet.Cells.Item($totalRow, 2).Formula = "=SUM(B8:B$($totalRow-1))"
        $summarySheet.Cells.Item($totalRow, 2).Font.Bold = $true
        $summarySheet.Cells.Item($totalRow, 2).HorizontalAlignment = -4108 # xlCenter
        $summarySheet.Cells.Item($totalRow, 6).Formula = "=SUM(F8:F$($totalRow-1))"
        $summarySheet.Cells.Item($totalRow, 6).Font.Bold = $true
        
        for ($col = 1; $col -le 7; $col++) {
            $summarySheet.Cells.Item($totalRow, $col).Borders.LineStyle = 1 # xlContinuous
        }
        
        # Instructions générales
        $summarySheet.Cells.Item(15, 1).Value = "Instructions générales"
        $summarySheet.Cells.Item(15, 1).Font.Bold = $true
        $summarySheet.Cells.Item(15, 1).Interior.ColorIndex = 34 # Light blue
        $summarySheet.Range("A15:H15").Merge() | Out-Null
        
        $instructions = @(
            "1. Ce runbook est conçu pour guider l'attachement des ressources aux DCRs (Data Collection Rules) pour le test d'ingestion.",
            "2. Suivez l'ordre des onglets pour une exécution optimale (commencez par les contrôleurs de domaine).",
            "3. Pour chaque type de serveur, utilisez la liste de vérification et mettez à jour le statut au fur et à mesure.",
            "4. En cas de problème, documentez-le dans la colonne 'Commentaires' et marquez le statut 'Problème rencontré'.",
            "5. Après avoir complété chaque section, revenez à cette feuille de résumé pour mettre à jour la progression globale.",
            "6. Prévoyez 60 serveurs par heure en moyenne pour l'attachement aux DCRs."
        )
        
        for ($i = 0; $i -lt $instructions.Count; $i++) {
            $row = $i + 16
            $summarySheet.Cells.Item($row, 1).Value = $instructions[$i]
            $summarySheet.Range("A$row:H$row").Merge() | Out-Null
            $summarySheet.Cells.Item($row, 1).HorizontalAlignment = -4131 # xlLeft
        }
        
        # Section planification
        $summarySheet.Cells.Item(23, 1).Value = "Planification horaire suggérée"
        $summarySheet.Cells.Item(23, 1).Font.Bold = $true
        $summarySheet.Cells.Item(23, 1).Interior.ColorIndex = 34 # Light blue
        $summarySheet.Range("A23:H23").Merge() | Out-Null
        
        $timelineHeaders = @("Heure", "Activité", "Responsable", "Détails")
        for ($i = 0; $i -lt $timelineHeaders.Count; $i++) {
            $col = $i*2 + 1
            $summarySheet.Cells.Item(24, $col).Value = $timelineHeaders[$i]
            $summarySheet.Cells.Item(24, $col).Font.Bold = $true
            $summarySheet.Cells.Item(24, $col).Interior.ColorIndex = 5 # Blue
            $summarySheet.Cells.Item(24, $col).Font.ColorIndex = 2 # White
            $summarySheet.Cells.Item(24, $col).HorizontalAlignment = -4108 # xlCenter
            $summarySheet.Cells.Item(24, $col).Borders.LineStyle = 1 # xlContinuous
            $summarySheet.Range($summarySheet.Cells.Item(24, $col), $summarySheet.Cells.Item(24, $col+1)).Merge() | Out-Null
        }
        
        $timeline = @(
            @("08:00 - 09:00", "Préparation et configuration initiale", "[Nom]", "Vérification de l'accès à tous les serveurs et aux DCRs"),
            @("09:00 - 10:00", "Attachement des contrôleurs de domaine", "[Nom]", "Prioriser les contrôleurs de domaine principaux"),
            @("10:00 - 12:00", "Attachement des serveurs Windows", "[Nom]", "Commencer par les serveurs applicatifs critiques"),
            @("13:00 - 14:00", "Attachement des serveurs SQL", "[Nom]", "Inclure SQL Server, MySQL et PostgreSQL"),
            @("14:00 - 15:30", "Attachement des serveurs Linux", "[Nom]", "Prioriser les serveurs Web et applications"),
            @("15:30 - 16:30", "Vérification et résolution des problèmes", "[Nom]", "Traiter les serveurs en erreur"),
            @("16:30 - 17:00", "Validation finale et documentation", "[Nom]", "Confirmer les flux de données pour chaque type de DCR")
        )
        
        for ($i = 0; $i -lt $timeline.Count; $i++) {
            $row = $i + 25
            $time, $activity, $responsible, $details = $timeline[$i]
            
            $summarySheet.Cells.Item($row, 1).Value = $time
            $summarySheet.Range("A$row:B$row").Merge() | Out-Null
            $summarySheet.Range("A$row:B$row").Borders.LineStyle = 1 # xlContinuous
            
            $summarySheet.Cells.Item($row, 3).Value = $activity
            $summarySheet.Range("C$row:D$row").Merge() | Out-Null
            $summarySheet.Range("C$row:D$row").Borders.LineStyle = 1 # xlContinuous
            
            $summarySheet.Cells.Item($row, 5).Value = $responsible
            $summarySheet.Range("E$row:F$row").Merge() | Out-Null
            $summarySheet.Range("E$row:F$row").Borders.LineStyle = 1 # xlContinuous
            
            $summarySheet.Cells.Item($row, 7).Value = $details
            $summarySheet.Range("G$row:H$row").Merge() | Out-Null
            $summarySheet.Range("G$row:H$row").Borders.LineStyle = 1 # xlContinuous
        }
        
        # Ajuster la largeur des colonnes
        $columnWidths = @(16, 20, 16, 20, 16, 20, 16, 20)
        for ($col = 1; $col -le $columnWidths.Count; $col++) {
            $summarySheet.Columns.Item($col).ColumnWidth = $columnWidths[$col-1]
        }
        
        return $summarySheet
    }

    # Fonction pour créer une feuille par type de serveur
    function Create-ServerSheet {
        param (
            $Title,
            $ServerData,
            $DcrName,
            $ServerType
        )
        
        $sheet = $workbook.Worksheets.Add()
        $sheet.Name = $Title
        
        # Titre
        $sheet.Cells.Item(1, 1).Value = "Attachement des $Title à $DcrName"
        $sheet.Cells.Item(1, 1).Font.Bold = $true
        $sheet.Cells.Item(1, 1).Font.Size = 14
        $sheet.Range("A1:I1").Merge() | Out-Null
        $sheet.Cells.Item(1, 1).HorizontalAlignment = -4108 # xlCenter
        
        # Description et instructions
        $sheet.Cells.Item(3, 1).Value = "Instructions pour l'attachement des $Title"
        $sheet.Cells.Item(3, 1).Font.Bold = $true
        $sheet.Cells.Item(3, 1).Interior.ColorIndex = 34 # Light blue
        $sheet.Range("A3:I3").Merge() | Out-Null
        
        $instructions = @(
            "1. Utilisez la commande PowerShell ou Azure CLI pour attacher les serveurs à la DCR $DcrName",
            "2. Pour chaque serveur, cochez la case 'Attaché' une fois l'opération réussie",
            "3. En cas d'erreur, notez le message d'erreur dans la colonne 'Commentaires' et marquez 'Erreur'",
            "4. Vérifiez que le flux de données commence à apparaître dans Log Analytics"
        )
        
        for ($i = 0; $i -lt $instructions.Count; $i++) {
            $row = $i + 4
            $sheet.Cells.Item($row, 1).Value = $instructions[$i]
            $sheet.Range("A$row:I$row").Merge() | Out-Null
            $sheet.Cells.Item($row, 1).HorizontalAlignment = -4131 # xlLeft
        }
        
        # Exemple de commande
        $sheet.Cells.Item(9, 1).Value = "Exemple de commande PowerShell:"
        $sheet.Cells.Item(9, 1).Font.Bold = $true
        $sheet.Range("A9:I9").Merge() | Out-Null
        
        $command = @"
# Pour l'agent Azure Monitor
`$dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "$DcrName"
`$vm = Get-AzVM -ResourceGroupName "<RG>" -Name "<SERVER_NAME>"
New-AzDataCollectionRuleAssociation -TargetResourceId `$vm.Id -AssociationName `$vm.Name -DataCollectionRuleId `$dcr.Id
"@
        
        $sheet.Cells.Item(10, 1).Value = $command
        $sheet.Range("A10:I12").Merge() | Out-Null
        $sheet.Cells.Item(10, 1).HorizontalAlignment = -4131 # xlLeft
        $sheet.Cells.Item(10, 1).VerticalAlignment = -4160 # xlTop
        $sheet.Cells.Item(10, 1).WrapText = $true
        
        # Section de progression
        $sheet.Cells.Item(14, 1).Value = "Progression:"
        $sheet.Cells.Item(14, 1).Font.Bold = $true
        $sheet.Cells.Item(14, 3).Value = "0%"
        $sheet.Cells.Item(14, 5).Value = "Temps estimé:"
        $sheet.Cells.Item(14, 5).Font.Bold = $true
        $sheet.Cells.Item(14, 7).Value = [math]::Max(1, [math]::Ceiling($ServerData.Count / 60)) + " heures"
        
        # En-têtes du tableau
        $headers = @("Serveur", "Système d'exploitation", "Version", "Statut Agent", "Groupe de ressources", "Attaché", "Statut", "Horodatage", "Commentaires")
        
        $row = 16
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $sheet.Cells.Item($row, $i+1).Value = $headers[$i]
            $sheet.Cells.Item($row, $i+1).Font.Bold = $true
            $sheet.Cells.Item($row, $i+1).Interior.ColorIndex = 5 # Blue
            $sheet.Cells.Item($row, $i+1).Font.ColorIndex = 2 # White
            $sheet.Cells.Item($row, $i+1).HorizontalAlignment = -4108 # xlCenter
            $sheet.Cells.Item($row, $i+1).Borders.LineStyle = 1 # xlContinuous
        }
        
        # Déterminer les noms de colonnes en fonction du type de serveur
        switch ($ServerType) {
            "dc" {
                $serverNameCol = "Sensor"
                $osCol = "Type"
                $versionCol = "Version"
                $statusCol = "Sensor status"
                $rgCol = $null
            }
            "windows" {
                $serverNameCol = "NAME"
                $osCol = "OPERATING SYSTEM"
                $versionCol = "OS VERSION"
                $statusCol = "ARC AGENT STATUS"
                $rgCol = "RESOURCE GROUP"
            }
            "linux" {
                $serverNameCol = "NAME"
                $osCol = "OPERATING SYSTEM"
                $versionCol = "OS VERSION"
                $statusCol = "ARC AGENT STATUS"
                $rgCol = "RESOURCE GROUP"
            }
            default {
                $serverNameCol = "NAME"
                $osCol = "OPERATING SYSTEM"
                $versionCol = "VERSION"
                $statusCol = "ARC AGENT STATUS"
                $rgCol = "RESOURCE GROUP"
            }
        }
        
        # Limiter à 100 premières lignes pour performance
        $maxRows = [math]::Min(100, $ServerData.Count)
        $startRow = $row + 1
        
        try {
            for ($i = 0; $i -lt $maxRows; $i++) {
                $currentRow = $startRow + $i
                $rowData = $ServerData[$i]
                
                # Serveur
                if ($rowData.$serverNameCol) {
                    $sheet.Cells.Item($currentRow, 1).Value = $rowData.$serverNameCol
                }
                
                # Système d'exploitation
                if ($osCol -and $rowData.$osCol) {
                    $sheet.Cells.Item($currentRow, 2).Value = $rowData.$osCol
                }
                
                # Version
                if ($versionCol -and $rowData.$versionCol) {
                    $sheet.Cells.Item($currentRow, 3).Value = $rowData.$versionCol
                }
                
                # Statut Agent
                if ($statusCol -and $rowData.$statusCol) {
                    $sheet.Cells.Item($currentRow, 4).Value = $rowData.$statusCol
                }
                
                # Groupe de ressources
                if ($rgCol -and $rowData.$rgCol) {
                    $sheet.Cells.Item($currentRow, 5).Value = $rowData.$rgCol
                }
                
                # Valeurs par défaut pour les autres colonnes
                $sheet.Cells.Item($currentRow, 6).Value = "Non"  # Attaché
                $sheet.Cells.Item($currentRow, 7).Value = "En attente"  # Statut
                
                # Ajouter les bordures à toutes les cellules
                for ($col = 1; $col -le $headers.Count; $col++) {
                    $sheet.Cells.Item($currentRow, $col).Borders.LineStyle = 1 # xlContinuous
                    
                    # Alignement spécifique pour certaines colonnes
                    if ($col -in 6, 7) {  # Attaché et Statut
                        $sheet.Cells.Item($currentRow, $col).HorizontalAlignment = -4108 # xlCenter
                    }
                }
            }
            
            # Ajouter des listes déroulantes pour les colonnes Attaché et Statut
            $startRange = [char]([int][char]'F') + $startRow.ToString()
            $endRange = [char]([int][char]'F') + ($startRow + $maxRows - 1).ToString()
            $validationRange = $sheet.Range($startRange + ":" + $endRange)
            
            $attachedValidation = $validationRange.Validation
            $attachedValidation.Delete()
            $attachedValidation.Add(3, 1, 1, "Oui,Non") # 3 = xlValidateList, 1 = xlValidAlertStop
            $attachedValidation.IgnoreBlank = $true
            $attachedValidation.InCellDropdown = $true
            
            $startRange = [char]([int][char]'G') + $startRow.ToString()
            $endRange = [char]([int][char]'G') + ($startRow + $maxRows - 1).ToString()
            $validationRange = $sheet.Range($startRange + ":" + $endRange)
            
            $statusValidation = $validationRange.Validation
            $statusValidation.Delete()
            $statusValidation.Add(3, 1, 1, "Réussi,Erreur,En attente") # 3 = xlValidateList, 1 = xlValidAlertStop
            $statusValidation.IgnoreBlank = $true
            $statusValidation.InCellDropdown = $true
            
            # Ajouter une formule pour la progression automatique
            if ($maxRows -gt 0) {
                $sheet.Cells.Item(14, 3).Formula = "=COUNTIF(F$startRow:F$($startRow+$maxRows-1),""Oui"")/COUNTA(A$startRow:A$($startRow+$maxRows-1))"
                $sheet.Cells.Item(14, 3).NumberFormat = "0.00%"
            }
        } catch {
            Write-Error "Erreur lors de l'ajout des données à la feuille $Title: $_"
        }
        
        # Ajuster la largeur des colonnes
        $columnWidths = @(25, 25, 15, 15, 20, 10, 10, 15, 40)
        for ($i = 0; $i -lt $columnWidths.Count; $i++) {
            $sheet.Columns.Item($i+1).ColumnWidth = $columnWidths[$i]
        }
        
        # Lien vers la feuille résumé
        $sheet.Cells.Item(3, 9).Value = "Retour au résumé"
        $sheet.Cells.Item(3, 9).Font.ColorIndex = 5
        $sheet.Cells.Item(3, 9).Font.Underline = $true
        $sheet.Hyperlinks.Add($sheet.Cells.Item(3, 9), "", "Résumé!A1", "", "Résumé") | Out-Null
        
        return $sheet
    }

    # Créer la feuille résumé
    $summarySheet = Create-SummarySheet
    
    # Créer les feuilles pour chaque type de serveur
    $dcSheet = Create-ServerSheet -Title "Contrôleurs de domaine" -ServerData $dfDc -DcrName "DCR-DomainController" -ServerType "dc"
    $winSheet = Create-ServerSheet -Title "Serveurs Windows" -ServerData $dfWindows -DcrName "DCR-WindowsServer" -ServerType "windows"
    $linuxSheet = Create-ServerSheet -Title "Serveurs Linux" -ServerData $dfLinux -DcrName "DCR-LinuxServer" -ServerType "linux"
    $sqlserverSheet = Create-ServerSheet -Title "SQL Server" -ServerData $dfSqlserver -DcrName "DCR-SQLServer" -ServerType "sqlserver"
    $mysqlSheet = Create-ServerSheet -Title "MySQL Server" -ServerData $dfMysql -DcrName "DCR-MySQLServer" -ServerType "mysql"
    $postgresSheet = Create-ServerSheet -Title "PostgreSQL Server" -ServerData $dfPostgres -DcrName "DCR-PostgresServer" -ServerType "postgres"
    
    # Ajouter une feuille Vérification finale
    $verificationSheet = $workbook.Worksheets.Add()
    $verificationSheet.Name = "Vérification finale"
    
    $verificationSheet.Cells.Item(1, 1).Value = "Vérification finale du test d'ingestion"
    $verificationSheet.Cells.Item(1, 1).Font.Bold = $true
    $verificationSheet.Cells.Item(1, 1).Font.Size = 14
    $verificationSheet.Range("A1:F1").Merge() | Out-Null
    $verificationSheet.Cells.Item(1, 1).HorizontalAlignment = -4108 # xlCenter
    
    $verificationSheet.Cells.Item(3, 1).Value = "Points de vérification et validation"
    $verificationSheet.Cells.Item(3, 1).Font.Bold = $true
    $verificationSheet.Cells.Item(3, 1).Interior.ColorIndex = 34 # Light blue
    $verificationSheet.Range("A3:F3").Merge() | Out-Null
    
    $checklist = @(
        @("Contrôleurs de domaine", "Vérifier les événements Windows Security (EventID 4624, 4625, etc.)"),
        @("Contrôleurs de domaine", "Vérifier les événements Directory Service"),
        @("Serveurs Windows", "Vérifier les événements Windows Security et System"),
        @("Serveurs Windows", "Vérifier les logs d'application critiques"),
        @("Serveurs SQL", "Vérifier les logs SQL Server et erreurs"),
        @("Serveurs MySQL", "Vérifier les logs d'erreur MySQL"),
        @("Serveurs PostgreSQL", "Vérifier les logs d'erreur PostgreSQL"),
        @("Serveurs Linux", "Vérifier les logs Syslog (auth, daemon, etc.)"),
        @("Général", "Vérifier le volume d'ingestion dans Log Analytics"),
        @("Général", "Valider que tous les serveurs sont affichés dans la section 'Agents' de Sentinel"),
        @("Général", "Confirmer qu'aucun problème n'est signalé dans le moniteur d'ingestion"),
        @("Estimation des coûts", "Extrapoler le coût journalier à partir des premières heures d'ingestion"),
        @("Estimation des coûts", "Comparer avec les estimations initiales")
    )
    
    # En-têtes du tableau
    $headers = @("Catégorie", "Point de vérification", "Statut", "Heure", "Vérificateur", "Commentaires")
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $verificationSheet.Cells.Item(5, $i+1).Value = $headers[$i]
        $verificationSheet.Cells.Item(5, $i+1).Font.Bold = $true
        $verificationSheet.Cells.Item(5, $i+1).Interior.ColorIndex = 5 # Blue
        $verificationSheet.Cells.Item(5, $i+1).Font.ColorIndex = 2 # White
        $verificationSheet.Cells.Item(5, $i+1).HorizontalAlignment = -4108 # xlCenter
        $verificationSheet.Cells.Item(5, $i+1).Borders.LineStyle = 1 # xlContinuous
    }
    
    # Ajouter les points de vérification
    for ($i = 0; $i -lt $checklist.Count; $i++) {
        $row = $i + 6
        $category, $check = $checklist[$i]
        
        $verificationSheet.Cells.Item($row, 1).Value = $category
        $verificationSheet.Cells.Item($row, 2).Value = $check
        $verificationSheet.Cells.Item($row, 3).Value = "Non vérifié"
        
        # Appliquer les styles
        for ($col = 1; $col -le $headers.Count; $col++) {
            $verificationSheet.Cells.Item($row, $col).Borders.LineStyle = 1 # xlContinuous
        }
    }
    
    # Ajouter des listes déroulantes pour la colonne Statut
    for ($i = 0; $i -lt $checklist.Count; $i++) {
        $row = $i + 6
        $statusValidation = $verificationSheet.Cells.Item($row, 3).Validation
        $statusValidation.Delete()
        $statusValidation.Add(3, 1, 1, "Non vérifié,Conforme,Non conforme,Partiellement conforme") # 3 = xlValidateList, 1 = xlValidAlertStop
        $statusValidation.IgnoreBlank = $true
        $statusValidation.InCellDropdown = $true
    }
    
    # Ajuster la largeur des colonnes
    $column_widths = @(20, 50, 15, 15, 15, 40)
    for ($i = 0; $i -lt $column_widths.Count; $i++) {
        $verificationSheet.Columns.Item($i+1).ColumnWidth = $column_widths[$i]
    }
    
    # Section pour la vérification des volumes d'ingestion
    $verificationSheet.Cells.Item(20, 1).Value = "Suivi des volumes d'ingestion"
    $verificationSheet.Cells.Item(20, 1).Font.Bold = $true
    $verificationSheet.Cells.Item(20, 1).Interior.ColorIndex = 34 # Light blue
    $verificationSheet.Range("A20:F20").Merge() | Out-Null
    
    $ingestion_headers = @("Type de source", "Volume après 1h (GB)", "Volume après 4h (GB)", "Volume après 8h (GB)", "Projection 24h (GB)", "Coût estimé ($)")
    for ($i = 0; $i -lt $ingestion_headers.Count; $i++) {
        $verificationSheet.Cells.Item(21, $i+1).Value = $ingestion_headers[$i]
        $verificationSheet.Cells.Item(21, $i+1).Font.Bold = $true
        $verificationSheet.Cells.Item(21, $i+1).Interior.ColorIndex = 5 # Blue
        $verificationSheet.Cells.Item(21, $i+1).Font.ColorIndex = 2 # White
        $verificationSheet.Cells.Item(21, $i+1).HorizontalAlignment = -4108 # xlCenter
        $verificationSheet.Cells.Item(21, $i+1).Borders.LineStyle = 1 # xlContinuous
    }
    
    # Types de sources à surveiller
    $source_types = @(
        "Contrôleurs de domaine",
        "Serveurs Windows",
        "Serveurs Linux",
        "SQL Server",
        "MySQL Server",
        "PostgreSQL Server",
        "Total"
    )
    
    for ($i = 0; $i -lt $source_types.Count; $i++) {
        $row = $i + 22
        $source = $source_types[$i]
        
        $verificationSheet.Cells.Item($row, 1).Value = $source
        # Les autres cellules sont vides mais formatées
        for ($col = 1; $col -le $ingestion_headers.Count; $col++) {
            $verificationSheet.Cells.Item($row, $col).Borders.LineStyle = 1 # xlContinuous
            if ($col -gt 1) {  # Colonnes numériques
                $verificationSheet.Cells.Item($row, $col).HorizontalAlignment = -4108 # xlCenter
                # Si c'est la ligne Total, ajouter une formule de somme
                if ($source -eq "Total" -and $col -lt 6) {  # Pas de formule pour la dernière colonne (coût)
                    $verificationSheet.Cells.Item($row, $col).Formula = "=SUM(" + [char]([int][char]'A' + $col-1) + "22:" + [char]([int][char]'A' + $col-1) + ($row-1) + ")"
                }
            }
        }
        
        # Pour la dernière colonne (coût estimé), ajouter une formule
        if ($source -eq "Total") {
            $verificationSheet.Cells.Item($row, 6).Formula = "=E" + $row + "*4.5"  # 4.5$ par GB en exemple
        } else {
            $verificationSheet.Cells.Item($row, 6).Formula = "=E" + $row + "*4.5"  # 4.5$ par GB en exemple
        }
    }
    
    # Lien vers la feuille résumé
    $verificationSheet.Cells.Item(3, 6).Value = "Retour au résumé"
    $verificationSheet.Cells.Item(3, 6).Font.ColorIndex = 5
    $verificationSheet.Cells.Item(3, 6).Font.Underline = $true
    $verificationSheet.Hyperlinks.Add($verificationSheet.Cells.Item(3, 6), "", "Résumé!A1", "", "Résumé") | Out-Null
    
    # Ajouter une feuille Commandes utiles
    $commandsSheet = $workbook.Worksheets.Add()
    $commandsSheet.Name = "Commandes utiles"
    
    $commandsSheet.Cells.Item(1, 1).Value = "Commandes utiles pour le test d'ingestion"
    $commandsSheet.Cells.Item(1, 1).Font.Bold = $true
    $commandsSheet.Cells.Item(1, 1).Font.Size = 14
    $commandsSheet.Range("A1:H1").Merge() | Out-Null
    $commandsSheet.Cells.Item(1, 1).HorizontalAlignment = -4108 # xlCenter
    
    $commandsSheet.Cells.Item(3, 1).Value = "Commandes PowerShell pour l'attachement des DCRs"
    $commandsSheet.Cells.Item(3, 1).Font.Bold = $true
    $commandsSheet.Cells.Item(3, 1).Interior.ColorIndex = 34 # Light blue
    $commandsSheet.Range("A3:H3").Merge() | Out-Null
    
    # Commandes PowerShell pour les différents types de serveurs
    $powershell_commands = @(
        @("Contrôleurs de domaine", @"
# Récupérer la DCR pour les contrôleurs de domaine
$dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-DomainController"

# Attacher un contrôleur de domaine à la DCR
$vm = Get-AzVM -ResourceGroupName "rg-infra" -Name "DOMAIN-DC01"
New-AzDataCollectionRuleAssociation -TargetResourceId $vm.Id -AssociationName $vm.Name -DataCollectionRuleId $dcr.Id

# Pour attacher plusieurs contrôleurs de domaine en une commande
$dcs = Get-AzVM -ResourceGroupName "rg-infra" | Where-Object {$_.Name -like "*DC*"}
foreach ($dc in $dcs) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $dc.Id -AssociationName $dc.Name -DataCollectionRuleId $dcr.Id
    Write-Host "Attaché: $($dc.Name)"
}
"@),
        
        @("Serveurs Windows", @"
# Récupérer la DCR pour les serveurs Windows
$dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-WindowsServer"

# Attacher un serveur Windows à la DCR
$vm = Get-AzVM -ResourceGroupName "rg-apps" -Name "APP-SRV01"
New-AzDataCollectionRuleAssociation -TargetResourceId $vm.Id -AssociationName $vm.Name -DataCollectionRuleId $dcr.Id

# Pour attacher plusieurs serveurs Windows en une commande (par lot de 10)
$windows_servers = Get-AzVM -ResourceGroupName "rg-apps" | Where-Object {$_.StorageProfile.OsDisk.OsType -eq "Windows"}
$counter = 0
foreach ($server in $windows_servers) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $server.Id -AssociationName $server.Name -DataCollectionRuleId $dcr.Id
    Write-Host "Attaché: $($server.Name)"
    $counter++
    if ($counter % 10 -eq 0) {
        Write-Host "Traitement de $counter serveurs terminé..."
        Start-Sleep -Seconds 5
    }
}
"@),
        
        @("Serveurs Linux", @"
# Récupérer la DCR pour les serveurs Linux
$dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-LinuxServer"

# Attacher un serveur Linux à la DCR
$vm = Get-AzVM -ResourceGroupName "rg-linux" -Name "LINUX-WEB01"
New-AzDataCollectionRuleAssociation -TargetResourceId $vm.Id -AssociationName $vm.Name -DataCollectionRuleId $dcr.Id

# Pour attacher plusieurs serveurs Linux en une commande
$linux_servers = Get-AzVM -ResourceGroupName "rg-linux" | Where-Object {$_.StorageProfile.OsDisk.OsType -eq "Linux"}
foreach ($server in $linux_servers) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $server.Id -AssociationName $server.Name -DataCollectionRuleId $dcr.Id
    Write-Host "Attaché: $($server.Name)"
}
"@),
        
        @("Serveurs SQL, MySQL et PostgreSQL", @"
# Récupérer les DCRs pour les différents types de bases de données
$sql_dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-SQLServer"
$mysql_dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-MySQLServer"
$postgres_dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName "DCR-PostgresServer"

# Attacher les serveurs SQL
$sql_servers = Get-AzVM | Where-Object {$_.Name -like "*SQL*" -or $_.Tags.DBType -eq "MSSQL"}
foreach ($server in $sql_servers) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $server.Id -AssociationName $server.Name -DataCollectionRuleId $sql_dcr.Id
    Write-Host "Attaché à SQL Server DCR: $($server.Name)"
}

# Attacher les serveurs MySQL
$mysql_servers = Get-AzVM | Where-Object {$_.Name -like "*MYSQL*" -or $_.Tags.DBType -eq "MySQL"}
foreach ($server in $mysql_servers) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $server.Id -AssociationName $server.Name -DataCollectionRuleId $mysql_dcr.Id
    Write-Host "Attaché à MySQL DCR: $($server.Name)"
}

# Attacher les serveurs PostgreSQL
$postgres_servers = Get-AzVM | Where-Object {$_.Name -like "*POSTGRES*" -or $_.Tags.DBType -eq "PostgreSQL"}
foreach ($server in $postgres_servers) {
    New-AzDataCollectionRuleAssociation -TargetResourceId $server.Id -AssociationName $server.Name -DataCollectionRuleId $postgres_dcr.Id
    Write-Host "Attaché à PostgreSQL DCR: $($server.Name)"
}
"@),
        
        @("Script pour générer un rapport des associations", @"
# Script pour générer un rapport des associations DCR créées
$associations = @()

# Fonction pour obtenir les associations et le statut pour une DCR
function Get-DCRAssociationStatus {
    param (
        [string]$DcrName
    )
    
    $dcr = Get-AzDataCollectionRule -ResourceGroupName "rg-sentinel-mxdr" -RuleName $DcrName
    $associations = Get-AzDataCollectionRuleAssociation -TargetResourceId /subscriptions/<SUBSCRIPTION_ID>/resourceGroups/*
    
    $dcr_associations = $associations | Where-Object { $_.DataCollectionRuleId -eq $dcr.Id }
    return $dcr_associations
}

# Obtenir les associations pour chaque DCR
$dcrs = @(
    "DCR-DomainController",
    "DCR-WindowsServer",
    "DCR-LinuxServer",
    "DCR-SQLServer",
    "DCR-MySQLServer",
    "DCR-PostgresServer"
)

foreach ($dcr in $dcrs) {
    $assoc = Get-DCRAssociationStatus -DcrName $dcr
    foreach ($a in $assoc) {
        $vm_id = $a.TargetResourceId
        $vm_name = $vm_id.Split('/')[-1]
        
        # Ajouter à la liste
        $associations += [PSCustomObject]@{
            ServerName = $vm_name
            DCR = $dcr
            AssociationName = $a.Name
            ProvisioningState = $a.ProvisioningState
            CreationTime = $a.CreationTime
        }
    }
}

# Exporter le rapport en CSV
$associations | Export-Csv -Path "C:\\Users\\lestinw\\log-ingestion\\DCR_Associations_Report.csv" -NoTypeInformation
"@),
        
        @("Vérifier l'ingestion des données", @"
# Pour vérifier si les données sont ingérées dans Log Analytics
# 1. Se connecter au portail Azure
# 2. Accéder à Log Analytics > rg-sentinel-mxdr-workspace
# 3. Exécuter la requête suivante pour vérifier les événements Windows
SecurityEvent
| where TimeGenerated > ago(1h)
| summarize count() by Computer
| order by count_ desc

# Pour vérifier les logs Syslog Linux
Syslog
| where TimeGenerated > ago(1h)
| summarize count() by Computer, Facility
| order by count_ desc

# Pour vérifier les logs SQL
AzureDiagnostics
| where TimeGenerated > ago(1h)
| where Category == "SQLSecurityAuditEvents"
| summarize count() by Resource
| order by count_ desc
"@)
    )
    
    # Ajouter ces commandes au document
    $row = 4
    for ($i = 0; $i -lt $powershell_commands.Count; $i++) {
        $title, $command = $powershell_commands[$i]
        
        # Titre de la commande
        $commandsSheet.Cells.Item($row, 1).Value = $title
        $commandsSheet.Cells.Item($row, 1).Font.Bold = $true
        $commandsSheet.Range("A$row" + ":H$row").Merge() | Out-Null
        
        # Commande
        $commandsSheet.Cells.Item(($row+1), 1).Value = $command.Trim()
        $commandsSheet.Range("A" + ($row+1) + ":H" + ($row+5)).Merge() | Out-Null
        $commandsSheet.Cells.Item(($row+1), 1).HorizontalAlignment = -4131 # xlLeft
        $commandsSheet.Cells.Item(($row+1), 1).VerticalAlignment = -4160 # xlTop
        $commandsSheet.Cells.Item(($row+1), 1).WrapText = $true
        $commandsSheet.Cells.Item(($row+1), 1).Font.Name = "Consolas"
        $commandsSheet.Cells.Item(($row+1), 1).Font.Size = 9
        
        # Ajouter un espacement
        $row += 7
    }
    
    # Ajuster les largeurs de colonnes
    for ($i = 1; $i -le 8; $i++) {
        $commandsSheet.Columns.Item($i).ColumnWidth = 15
    }
    
    # Lien vers la feuille résumé
    $commandsSheet.Cells.Item(3, 8).Value = "Retour au résumé"
    $commandsSheet.Cells.Item(3, 8).Font.ColorIndex = 5
    $commandsSheet.Cells.Item(3, 8).Font.Underline = $true
    $commandsSheet.Hyperlinks.Add($commandsSheet.Cells.Item(3, 8), "", "Résumé!A1", "", "Résumé") | Out-Null
    
    # Réorganiser les feuilles
    $summarySheet.Move($null, $workbook.Sheets.Item($workbook.Sheets.Count))
    $dcSheet.Move($null, $summarySheet)
    $winSheet.Move($null, $dcSheet)
    $linuxSheet.Move($null, $winSheet)
    $sqlserverSheet.Move($null, $linuxSheet)
    $mysqlSheet.Move($null, $sqlserverSheet)
    $postgresSheet.Move($null, $mysqlSheet)
    $verificationSheet.Move($null, $postgresSheet)
    $commandsSheet.Move($null, $verificationSheet)
    
    # Sauvegarder le workbook
    $workbook.SaveAs($outputFile)
    $workbook.Close($true)
    $excel.Quit()
    
    Write-Host "Le runbook a été créé avec succès: $outputFile"
} catch {
    Write-Error "Erreur lors de la création du fichier Excel: $_"
} finally {
    # Libérer les ressources COM
    if ($excel -ne $null) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Création du runbook terminée!"