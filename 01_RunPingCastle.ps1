Start-Sleep -Seconds 5
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# ====================
# == VARIABLES =======
# ====================
$ApplicationName = 'PingCastle'
$PingCastle = [pscustomobject]@{
    Name            = $ApplicationName
    ProgramPath     = Join-Path $PSScriptRoot $ApplicationName
    ProgramName     = '{0}.exe' -f $ApplicationName
    Argument1       = '--healthcheck'
    Argument2       = '--level Full'
    ReportFileName  = 'ad_hc_{0}' -f ($env:USERDNSDOMAIN).ToLower()
    ReportFolder    = "Reports"
    LogFolder       = "Logs"
    ScoreFileName   = '{0}Score.txt' -f $ApplicationName
    ProgramUpdate   = '{0}AutoUpdater.exe' -f $ApplicationName
    ArgumentsUpdate = '--wait-for-days 30'
}

$pingCastleFullpath            = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramName
$pingCastleUpdateFullpath      = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramUpdate
$pingCastleReportLogs          = Join-Path $PingCastle.ProgramPath $PingCastle.ReportFolder

$pingCastleScoreFileFullpath   = Join-Path $pingCastleReportLogs $PingCastle.ScoreFileName
$pingCastleReportFullpath      = Join-Path $PingCastle.ProgramPath ('{0}.html' -f $PingCastle.ReportFileName)
$pingCastleReportXMLFullpath   = Join-Path $PingCastle.ProgramPath ('{0}.xml' -f $PingCastle.ReportFileName)

$pingCastleReportDate          = Get-Date -UFormat %Y%m%d_%H%M%S
$pingCastleDate                = Get-Date -Format "dd/MM/yyyy"
$pingCastleReportFileNameDate  = ('{0}_{1}.html' -f $pingCastleReportDate, $PingCastle.ReportFileName)

$sentNotification = $false
$scoreTrend = "same"

$splatProcess = @{
    FilePath = $pingcastleFullpath
    ArgumentList = $PingCastle.Arguments
    WindowStyle = 'Hidden'
    Wait        = $true
}

# ====================
# == LOGGING SETUP  ==
# ====================

$logDirectory = Join-Path $PingCastle.ProgramPath "Logs"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$LogFilePath = Join-Path $logDirectory ("PingCastle.log" -f $timestamp)
function Write-Log {
    param (
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO", "ERROR", "WARN")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFilePath -Value $logMessage
    Write-Output $logMessage
}



# Créer le dossier Logs s'il n'existe pas
if (-not (Test-Path $logDirectory)) {
    try {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
        Write-Log "Répertoire de logs créé : $logDirectory" 
    } catch {
        Write-Log "Erreur lors de la création du répertoire de logs : $logDirectory" -Level "ERROR"
        exit 1
    }
} else {
    Write-Log "Répertoire de logs déjà existant : $logDirectory" -Level "INFO"
}

# ====================
# == START SCRIPT ====
# ====================
Write-Log "=== Début du script PingCastle ===" -Level "INFO"

# Check if PingCastle is present
if (-not(Test-Path $pingCastleFullpath)) {
    Write-Log "PingCastle introuvable à l'emplacement : $pingCastleFullpath" -Level "ERROR"
    exit 1
}
Write-Log "PingCastle trouvé : $pingCastleFullpath" -Level "INFO"
Write-Log "Commande exécutée : $pingCastleFullpath $($PingCastle.Argument1)" -Level "INFO"

# Create report directory if not exists
if (-not (Test-Path $pingCastleReportLogs)) {
    try {
        $null = New-Item -Path $pingCastleReportLogs -ItemType directory
        Write-Log "Répertoire de report créé : $pingCastleReportLogs" -Level "INFO"
    } catch {
        Write-Log "Erreur lors de la création du dossier $pingCastleReportLogs" -Level "ERROR"
        exit 1
    }
}


# Launch PingCastle HealthCheck
Write-Log "Lancement de PingCastle..." -Level "INFO"
try {
    Set-Location -Path $PingCastle.ProgramPath
    $pingcastlepathprogram = $PingCastle.ProgramPath
    try {
        #Start-Process @splatProcess
        & "$pingCastleFullpath" $PingCastle.Argument1
        Write-Log "Code retour PingCastle : $LASTEXITCODE" -Level "INFO"
        Write-Log "PingCastle exécuté" -Level "INFO"
    }catch {
        Write-Log "Erreur dans le lancement de PingCastle" -Level "INFO"
    }
    Write-Log "Analyse PingCastle terminée." -Level "INFO"
} catch {
    Write-Log "Erreur lors de l'exécution de PingCastle 2." -Level "ERROR"
    exit 1
}

# Vérification de la génération des rapports
foreach ($file in @($pingCastleReportFullpath, $pingCastleReportXMLFullpath)) {
    if (-not (Test-Path $file)) {
        Write-Log "Fichier de rapport manquant : $file" -Level "ERROR"
        exit 1
    } else {
        Write-Log "Fichier généré : $file" -Level "INFO"
    }
}

# Lecture XML
try {
    $contentPingCastleReportXML = (Select-Xml -Path $pingCastleReportXMLFullpath -XPath "/HealthcheckData").node
    Write-Log "Contenu XML chargé avec succès" -Level "INFO"
} catch {
    Write-Log "Impossible de lire le contenu du fichier XML." -Level "ERROR"
    exit 1
}

# === Récupération des scores depuis le fichier PingCastle XML déjà chargé ===
$reportDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$globalScore           = $contentPingCastleReportXML.GlobalScore
$staleObjectsScore     = $contentPingCastleReportXML.StaleObjectsScore
$privilegedGroupScore  = $contentPingCastleReportXML.PrivilegiedGroupScore
$trustScore            = $contentPingCastleReportXML.TrustScore
$anomalyScore          = $contentPingCastleReportXML.AnomalyScore

# === Définition des chemins ===
$scoreHistoryPath = Join-Path $PingCastle.ProgramPath "Scores/score_historique.xml"
$lastScorePath = Join-Path $PingCastle.ProgramPath "Scores/last_score.xml"
$lastReportPath = Join-Path $PingCastle.ProgramPath "Logs/last_report.xml"

# Calcul du numéro de semaine (ISO 8601)
$calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
$weekRule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
$firstDayOfWeek = [System.DayOfWeek]::Monday
$weekNumber = $calendar.GetWeekOfYear([datetime]$reportDateTime, $weekRule, $firstDayOfWeek)

# === Création ou mise à jour du fichier score_historique.xml ===
if (-not (Test-Path $scoreHistoryPath)) {
    $xmlHeader = '<?xml version="1.0" encoding="utf-8"?>'
    $xmlBody = "<Scores>
    <Score date='$reportDateTime' week='$weekNumber' Global='$globalScore' Stale='$staleObjectsScore' Privileged='$privilegedGroupScore' Trust='$trustScore' Anomaly='$anomalyScore' />
</Scores>"
    "$xmlHeader`n$xmlBody" | Set-Content -Encoding UTF8 -Path $scoreHistoryPath
    Write-Log "Fichier historique créé avec une première entrée." -Level "INFO"
} else {
    [xml]$xmlDoc = Get-Content -Path $scoreHistoryPath

    $newNode = $xmlDoc.CreateElement("Score")
    $newNode.SetAttribute("date", $reportDateTime)
    $newNode.SetAttribute("week", $weekNumber)
    $newNode.SetAttribute("Global", $globalScore)
    $newNode.SetAttribute("Stale", $staleObjectsScore)
    $newNode.SetAttribute("Privileged", $privilegedGroupScore)
    $newNode.SetAttribute("Trust", $trustScore)
    $newNode.SetAttribute("Anomaly", $anomalyScore)

    $xmlDoc.Scores.AppendChild($newNode) | Out-Null
    $xmlDoc.Save($scoreHistoryPath)
    Write-Log "Score ajouté au fichier historique (semaine $weekNumber)." -Level "INFO"
}


# === Création/Mise à jour du fichier last_score.xml ===

try {
    # Lecture de la valeur du dernier GlobalScore dans le fichier last_score.xml
    [xml]$lastScoreXml = Get-Content $lastScorePath
    $lastGlobalScore = $lastScoreXml.LastScore.Global
    $lastTrustScore = $lastScoreXml.LastScore.Trust
    $lastPrivilegedScore = $lastScoreXml.LastScore.Privileged
    $lastStaleScore = $lastScoreXml.LastScore.Stale
    $lastAnomalyScore = $lastScoreXml.LastScore.Anomaly


    # Comparaison avec le score actuel
    if ($globalScore -ne $lastGlobalScore) {
        Write-Log "Le GlobalScore a changé : $lastGlobalScore -> $globalScore. Une notification va être envoyé !" -Level "INFO"
        $sentNotification = $true
    } else {
        Write-Log "Le GlobalScore est identique ($globalScore). Aucune notification nécessaire." -Level "INFO"
        $sentNotification = $false
    }

    if ($trustScore -ne $lastTrustScore) {
        Write-Log "Le TrustScore a changé : $lastTrustScore -> $trustScore. Une notification va être envoyé !" -Level "INFO"
        $sentNotification = $true
    } else {
        Write-Log "Le TrustScore est identique ($trustScore). Aucune notification nécessaire." -Level "INFO"
        $sentNotification = $false
    }

    if ($anomalyScore -ne $lastAnomalyScore) {
        Write-Log "Le anomalyScore a changé : $lastAnomalyScore -> $anomalyScore. Une notification va être envoyé !" -Level "INFO"
        $sentNotification = $true
    } else {
        Write-Log "Le anomalyScore est identique ($anomalyScore). Aucune notification nécessaire." -Level "INFO"
        $sentNotification = $false
    }

    if ($privilegedGroupScore -ne $lastPrivilegedScore) {
        Write-Log "Le PriviligedScore a changé : $lastPrivilegedScore -> $privilegedGroupScore. Une notification va être envoyé !" -Level "INFO"
        $sentNotification = $true
    } else {
        Write-Log "Le PriviligedScore est identique ($privilegedGroupScore). Aucune notification nécessaire." -Level "INFO"
        $sentNotification = $false
    }

    if ($staleObjectsScore -ne $lastStaleScore) {
        Write-Log "Le StaleScore a changé : $lastStaleScore -> $staleObjectsScore. Une notification va être envoyé !" -Level "INFO"
        $sentNotification = $true
    } else {
        Write-Log "Le StaleScore est identique ($staleObjectsScore). Aucune notification nécessaire." -Level "INFO"
        $sentNotification = $false
    }
    # Comparaison avec le score actuel

    [int]$globalScore = [int]$globalScore
    [int]$lastGlobalScore = [int]$lastGlobalScore

    if ($globalScore -ne $lastGlobalScore) {
        $sentNotification = $true

        if ($globalScore -gt $lastGlobalScore) {
            $scoreTrend = "up"
            Write-Log "Le GlobalScore a augmenté : $lastGlobalScore -> $globalScore. Une notification va être envoyée !" -Level "INFO"
        } elseif ($globalScore -lt $lastGlobalScore) {
            $scoreTrend = "down"
            Write-Log "Le GlobalScore a diminué : $lastGlobalScore -> $globalScore. Une notification va être envoyée !" -Level "INFO"
        }
    } else {
        Write-Log "Le GlobalScore est identique ($globalScore). Aucune notification nécessaire." -Level "INFO"
    }
}
catch {
    Write-Log "Erreur lors de la lecture ou comparaison du GlobalScore." -Level "ERROR"
    $sentNotification = $true
}

$lastScoreContent = @"
<?xml version="1.0" encoding="utf-8"?>
<LastScore 
    date="$reportDateTime"
    Global="$globalScore"
    Stale="$staleObjectsScore"
    Privileged="$privilegedGroupScore"
    Trust="$trustScore"
    Anomaly="$anomalyScore" />
"@

$lastScoreContent | Set-Content -Encoding UTF8 -Path $lastScorePath
Write-Log "Fichier last_score.xml mis à jour." -Level "INFO"

# Envoi de mail
if ($sentNotification) {
    # Chargement du rapport HTML complet
    $reportHtmlContent = Get-Content -Path $pingCastleReportFullpath -Raw

    # Corps personnalisé du mail selon l'évolution du score
    switch ($scoreTrend) {
        "up" {
            $statusMessage = "<p style='color:red; font-weight:bold;'>⚠️ Le score de sécurité a <u>augmenté</u> de $lastGlobalScore à $globalScore. Une régression potentielle a été détectée.</p>"
        }
        "down" {
            $statusMessage = "<p style='color:green; font-weight:bold;'>✅ Le score de sécurité a <u>diminué</u> de $lastGlobalScore à $globalScore. Amélioration détectée.</p>"
        }
    }

    # Récupération des variables scores
    $globalScore           = [int]$contentPingCastleReportXML.GlobalScore
    $staleObjectsScore     = [int]$contentPingCastleReportXML.StaleObjectsScore
    $privilegedGroupScore  = [int]$contentPingCastleReportXML.PrivilegiedGroupScore
    $trustScore            = [int]$contentPingCastleReportXML.TrustScore
    $anomalyScore          = [int]$contentPingCastleReportXML.AnomalyScore

    function Get-GaugeHtml {
        param(
            [string]$title,
            [int]$value
        )

        # Détermination de la couleur
        if ($value -lt 20) {
            $color = "green"
        } elseif ($value -lt 50) {
            $color = "orange"
        } else {
            $color = "red"
        }

        return @"
        <div style='margin-bottom: 10px;'>
            <p style='margin: 0; font-weight: bold;'>$title : $value</p>
            <div style='width: 100%; background-color: #eee; height: 20px; border-radius: 5px;'>
                <div style='width: $value%; background-color: $color; height: 100%; border-radius: 5px;'></div>
            </div>
        </div>
"@
    }

    $gaugesHtml = Get-GaugeHtml "Global Score" $globalScore
    $gaugesHtml += Get-GaugeHtml "Stale Object Score" $staleObjectsScore
    $gaugesHtml += Get-GaugeHtml "Trust Score" $privilegedGroupScore
    $gaugesHtml += Get-GaugeHtml "Privilege Score" $trustScore
    $gaugesHtml += Get-GaugeHtml "Anomaly Score" $anomalyScore

        # Construction du corps HTML
    $mailBody = @"
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            background-color: #ffffff;
            color: #333333;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
        }

        .header {
            background-color: #005A9C;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .header img {
            height: 50px;
        }

        .container {
            padding: 30px;
            width : 850px
        }

        h2 {
            color: #005A9C;
            margin-bottom: 10px;
        }

        h3 {
            color: #0078D7;
            margin-top: 30px;
        }

        .score {
            margin-top: 15px;
            font-size: 1.1em;
        }

        .warning {
            color: #d9534f;
        }

        .success {
            color: #5cb85c;
        }

        .footer {
            margin-top: 40px;
            font-size: 0.85em;
            color: #666666;
            border-top: 1px solid #cccccc;
            padding-top: 20px;
        }

        a {
            color: #0078D7;
        }
    </style>
</head>
<body>
    <!-- Bandeau en-tête -->

    <!-- Corps principal -->
    <div class="container">
        <h2>📊 Rapport de Sécurité PingCastle - $pingCastleDate</h2>

        <p class="score">$statusMessage</p>

        <h3>Détail des scores :</h3>
        $gaugesHtml

        <div class="footer">
            <p><strong>Note :</strong> Ceci est un message automatique généré par le système de surveillance PingCastle. Merci de ne pas répondre directement à cet e-mail.</p>
            <p>Pour toute question, vous pouvez contacter L'<a href="mailto:support-infra@eure.fr">équipe infra</a>.</p>
            <p>Le rapport complet est disponible en pièce jointe.</p>
        </div>
    </div>
</body>
</html>
"@
    $mailBodyBytes = [System.Text.Encoding]::UTF8.GetBytes($mailBody)
    $mailBodyEncoded = [System.Text.Encoding]::UTF8.GetString($mailBodyBytes)
    $alternateView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($htmlBody, $null, "text/html")

    # Paramètres SMTP
    $smtpServer = "SERVER_SMTP"
    $smtpPort = X
    $from = "pingcastle@gmail.fr"
    $to = "xxx@xxxx.fr"
    $subject = "Rapport journalier PingCastle - $pingCastleDate"

    try {
        Send-MailMessage -From $from -To $to -Subject $subject -Body $mailBody -BodyAsHtml -SmtpServer $smtpServer -Port $smtpPort -Attachments $pingCastleReportFullpath -Encoding 'utf8'
        Write-Log "Mail envoyé à $to avec les scores" -Level "INFO"
    } catch {
        Write-Log "Erreur lors de l'envoi de l'e-mail : $_ " -Level "ERROR"
    }
} else {
    Write-Log "Mail non envoyé, non nécessaire !" -Level "INFO"
}


# Déplacement des rapports
try {
    # Vérifie que les fichiers existent encore
    if (-not (Test-Path $pingCastleReportFullpath)) {
        Write-Log "Le fichier HTML n'existe plus : $pingCastleReportFullpath" -Level "ERROR"
        exit 1
    }

    if (-not (Test-Path $pingCastleReportXMLFullpath)) {
        Write-Log "Le fichier XML n'existe plus : $pingCastleReportXMLFullpath" -Level "ERROR"
        exit 1
    }

    # Assure-toi que le dossier de destination existe
    if (-not (Test-Path $pingCastleReportLogs)) {
        New-Item -Path $pingCastleReportLogs -ItemType Directory -Force | Out-Null
        Write-Log "Répertoire de rapport recréé : $pingCastleReportLogs" -Level "INFO"
    }

    # Construction des chemins finaux
    $htmlDestination = Join-Path $pingCastleReportLogs $pingCastleReportFileNameDate
    $xmlDestination = Join-Path $pingCastleReportLogs ("{0}_{1}.xml" -f $pingCastleReportDate, $PingCastle.ReportFileName)
    $LastReportDestination = Join-Path $pingCastleReportLogs "Last_report.xml"

    Write-Log "Déplacement du fichier HTML vers : $htmlDestination" -Level "INFO"
    Write-Log "Déplacement du fichier XML vers : $xmlDestination" -Level "INFO"

    Copy-Item $pingCastleReportFullpath $LastReportDestination -Force
    Move-Item $pingCastleReportFullpath $htmlDestination -Force
    Move-Item $pingCastleReportXMLFullpath $xmlDestination -Force

    Write-Log "Rapports déplacés avec succès dans le répertoire : $pingCastleReportLogs" -Level "INFO"
} catch {
    Write-Log "Erreur lors du déplacement des fichiers de rapport : $($_.Exception.Message)" -Level "ERROR"
}


# Mise à jour PingCastle
try {
    Write-Log "Mise à jour de PingCastle en cours..." -Level "INFO"
    Start-Process -FilePath $pingCastleUpdateFullpath -ArgumentList $PingCastle.ArgumentsUpdate @splatProcess
    Write-Log "Mise à jour PingCastle terminée." -Level "INFO"
} catch {
    Write-Log "Erreur lors de la mise à jour de PingCastle." -Level "ERROR"
}

Write-Log "=== Fin du script PingCastle ===" -Level "INFO"
