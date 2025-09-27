# ====================
# == VARIABLES =======
# ====================
$ApplicationName = 'PingCastle'
$PingCastle = [pscustomobject]@{
    Name            = $ApplicationName
    ProgramPath     = Join-Path $PSScriptRoot $ApplicationName
    ProgramName     = '{0}.exe' -f $ApplicationName
    Arguments       = '--healthcheck --level Full'
    ReportFileName  = 'ad_hc_{0}' -f ($env:USERDNSDOMAIN).ToLower()
    ReportFolder    = "Reports"
    LogFolder       = "Logs"
    ScoreFileName   = '{0}Score.txt' -f $ApplicationName
    ProgramUpdate   = '{0}AutoUpdater.exe' -f $ApplicationName
    ArgumentsUpdate = '--wait-for-days 30'
}

$logDirectory = Join-Path $PingCastle.ProgramPath "Logs"
$LogFilePath = Join-Path $logDirectory "PingCastle.log"

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

# ========================
# == CHEMINS & DONNEES ==
# ========================
$scoreHistoryPath = Join-Path $PingCastle.ProgramPath "\Scores\score_historique.xml"

if (-not (Test-Path $scoreHistoryPath)) {
    Write-Log "Fichier historique introuvable : $scoreHistoryPath" -Level "ERROR"
    return
}

[xml]$xmlDoc = Get-Content -Path $scoreHistoryPath
$reportDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
$weekRule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
$firstDayOfWeek = [System.DayOfWeek]::Monday
$currentweekNumber = $calendar.GetWeekOfYear([datetime]$reportDateTime, $weekRule, $firstDayOfWeek)

Write-Log "Semaine en cours : $currentWeekNumber" -Level "INFO"

$scores = $xmlDoc.Scores.Score | Where-Object { $_.week -eq $currentWeekNumber }
if (-not $scores) {
    Write-Log "Aucune donnée trouvée pour la semaine $currentWeekNumber" -Level "WARN"
    return
}

# Moyennes de la semaine
$avgGlobal = [math]::Round(($scores.Global | Measure-Object -Average).Average, 2)
$avgStale = [math]::Round(($scores.Stale | Measure-Object -Average).Average, 2)
$avgPrivileged = [math]::Round(($scores.Privileged | Measure-Object -Average).Average, 2)
$avgTrust = [math]::Round(($scores.Trust | Measure-Object -Average).Average, 2)
$avgAnomaly = [math]::Round(($scores.Anomaly | Measure-Object -Average).Average, 2)

Write-Log "Calcul des moyennes de la semaine $currentWeekNumber : Global=$avgGlobal, Stale=$avgStale, Privileged=$avgPrivileged, Trust=$avgTrust, Anomaly=$avgAnomaly" -Level "INFO"

# Moyennes totales
$totalScores = $xmlDoc.Scores.Score
$totalAvgGlobal = [math]::Round(($totalScores.Global | Measure-Object -Average).Average, 2)
$totalAvgStale = [math]::Round(($totalScores.Stale | Measure-Object -Average).Average, 2)
$totalAvgPrivileged = [math]::Round(($totalScores.Privileged | Measure-Object -Average).Average, 2)
$totalAvgTrust = [math]::Round(($totalScores.Trust | Measure-Object -Average).Average, 2)
$totalAvgAnomaly = [math]::Round(($totalScores.Anomaly | Measure-Object -Average).Average, 2)

Write-Log "Calcul des moyennes totales : Global=$totalAvgGlobal, Stale=$totalAvgStale, Privileged=$totalAvgPrivileged, Trust=$totalAvgTrust, Anomaly=$totalAvgAnomaly" -Level "INFO"

# Préparation des données pour le graphique
$weeks = @()
$globalData = @()
$staleData = @()
$privilegedData = @()
$trustData = @()
$anomalyData = @()

foreach ($entry in $totalScores) {
    $entryDate = [datetime]$entry.date
    $weekNum = $calendar.GetWeekOfYear($entryDate, $weekRule, $firstDayOfWeek)
    $weeks += "S$weekNum"

    $globalData += [int]$entry.Global
    $staleData += [int]$entry.Stale
    $privilegedData += [int]$entry.Privileged
    $trustData += [int]$entry.Trust
    $anomalyData += [int]$entry.Anomaly
}

# Nouvelle méthode pour éviter les superpositions de courbes, peu importe la valeur
$globalDataAdjusted     = @()
$staleDataAdjusted      = @()
$privilegedDataAdjusted = @()
$trustDataAdjusted      = @()
$anomalyDataAdjusted    = @()

for ($i = 0; $i -lt $globalData.Count; $i++) {
    # On crée un dictionnaire de la valeur => liste des courbes qui ont cette valeur à l'index $i
    $valueGroups = @{}
    $allValues = @{
        "global"     = $globalData[$i]
        "stale"      = $staleData[$i]
        "privileged" = $privilegedData[$i]
        "trust"      = $trustData[$i]
        "anomaly"    = $anomalyData[$i]
    }

    foreach ($key in $allValues.Keys) {
        $val = $allValues[$key]
        if (-not $valueGroups.ContainsKey($val)) {
            $valueGroups[$val] = @()
        }
        $valueGroups[$val] += $key
    }

    # Maintenant, on applique un petit décalage pour les doublons
    $adjustedValues = @{}
    foreach ($val in $valueGroups.Keys) {
        $keys = $valueGroups[$val]
        for ($j = 0; $j -lt $keys.Count; $j++) {
            $adjusted = [math]::Round($val - ($j * 1), 2)
            $adjustedValues[$keys[$j]] = $adjusted
        }
    }

    # On remplit les listes ajustées
    $globalDataAdjusted     += $adjustedValues["global"]
    $staleDataAdjusted      += $adjustedValues["stale"]
    $privilegedDataAdjusted += $adjustedValues["privileged"]
    $trustDataAdjusted      += $adjustedValues["trust"]
    $anomalyDataAdjusted    += $adjustedValues["anomaly"]
}

Write-Log "Calcul des moyennes totales : Global=$globalDataAdjusted, Stale=$staleDataAdjusted, Privileged=$privilegedDataAdjusted, Trust=$trustDataAdjusted, Anomaly=$anomalyDataAdjusted" -Level "INFO"
# =====================
# == GÉNÉRATION HTML ==
# =====================


Add-Type -AssemblyName System.Drawing

# Crée l'image
$width = 1000
$height = 600
$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::White)

# Polices
$font = New-Object System.Drawing.Font(
    "Arial",
    10,
    [System.Drawing.FontStyle]::Regular,
    [System.Drawing.GraphicsUnit]::Point
)

$titleFont = New-Object System.Drawing.Font(
    "Arial",
    14,
    [System.Drawing.FontStyle]::Bold,
    [System.Drawing.GraphicsUnit]::Point
)


# Axes
$leftMargin = 60
$topMargin = 50
$bottomMargin = 50
$rightMargin = 30
$chartHeight = $height - $topMargin - $bottomMargin
$chartWidth = $width - $leftMargin - $rightMargin

# Crée un stylo pour la grille (graduations de l'axe Y)
$gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::LightGray)
$gridPen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash

# === Légende des axes ===
# Définir les polices et brosses
$axisFont = New-Object System.Drawing.Font("Arial", 10)
$labelBrush = [System.Drawing.Brushes]::Black

# Label Axe Y (Ordonnée)
$graphics.DrawString("Score", $axisFont, $labelBrush, 10, 10)

# Label Axe X (Abscisse)
$graphics.DrawString("Semaines", $axisFont, $labelBrush, $width / 2 - 30, $height - 30)

# Graduation Axe Y
for ($y = 0; $y -le 100; $y += 20) {
    $posY = $height - 50 - ($y * ($chartHeight / 100))
    $graphics.DrawLine($gridPen, 45, $posY, $width - 50, $posY)
    $graphics.DrawString($y.ToString(), $axisFont, $labelBrush, 5, $posY - 7)
}

# Graduation Axe X (semaines)
$stepX = $chartWidth / ($weeks.Count - 1)
for ($i = 0; $i -lt $weeks.Count; $i++) {
    $x = 50 + ($i * $stepX)
    $graphics.DrawString($weeks[$i], $axisFont, $labelBrush, $x - 10, $height - 45)
}

# Valeur max pour l'échelle
$allValues = $globalData + $staleData + $privilegedData + $trustData + $anomalyData
$maxValue = ($allValues | Measure-Object -Maximum).Maximum


# Tracer chaque série
function Draw-Line {
    param (
        [int[]]$data,
        [System.Drawing.Pen]$pen
    )
    for ($i = 1; $i -lt $data.Count; $i++) {
        $x1 = $leftMargin + ($i - 1) * ($chartWidth / ($data.Count - 1))
        $x2 = $leftMargin + $i * ($chartWidth / ($data.Count - 1))
        $y1 = $topMargin + $chartHeight - ($data[$i - 1] / $maxValue * $chartHeight)
        $y2 = $topMargin + $chartHeight - ($data[$i] / $maxValue * $chartHeight)
        $graphics.DrawLine($pen, $x1, $y1, $x2, $y2)
    }
}

# Créer des stylos pour chaque série
$penGlobal = New-Object System.Drawing.Pen ([System.Drawing.Color]::Blue), 2
$penStale = New-Object System.Drawing.Pen ([System.Drawing.Color]::Green), 2
$penPriv = New-Object System.Drawing.Pen ([System.Drawing.Color]::Orange), 2
$penTrust = New-Object System.Drawing.Pen ([System.Drawing.Color]::Purple), 2
$penAnomaly = New-Object System.Drawing.Pen ([System.Drawing.Color]::Red), 2

Draw-Line -data $globalDataAdjusted -pen $penGlobal
Draw-Line -data $staleDataAdjusted -pen $penStale
Draw-Line -data $privilegedDataAdjusted -pen $penPriv
Draw-Line -data $trustDataAdjusted -pen $penTrust
Draw-Line -data $anomalyDataAdjusted -pen $penAnomaly

# Légende
$graphics.FillRectangle([System.Drawing.Brushes]::Blue, 60, $height - 15, 10, 10)
$graphics.DrawString("Global", $font, [System.Drawing.Brushes]::Black, 75, $height - 17)
$graphics.FillRectangle([System.Drawing.Brushes]::Green, 140, $height - 15, 10, 10)
$graphics.DrawString("Stale", $font, [System.Drawing.Brushes]::Black, 155, $height - 17)
$graphics.FillRectangle([System.Drawing.Brushes]::Orange, 220, $height - 15, 10, 10)
$graphics.DrawString("Privileged", $font, [System.Drawing.Brushes]::Black, 235, $height - 17)
$graphics.FillRectangle([System.Drawing.Brushes]::Purple, 320, $height - 15, 10, 10)
$graphics.DrawString("Trust", $font, [System.Drawing.Brushes]::Black, 335, $height - 17)
$graphics.FillRectangle([System.Drawing.Brushes]::Red, 400, $height - 15, 10, 10)
$graphics.DrawString("Anomaly", $font, [System.Drawing.Brushes]::Black, 415, $height - 17)

# Sauvegarde en PNG
$outputPath = "..path\PingCastle\Logs\graph"+$currentWeekNumber+".png"
$bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()
Write-Log "Graphique sauvegardé à l'emplacement : $outputPath" -Level INFO

# Création du message
$mailMessage = New-Object System.Net.Mail.MailMessage
$mailMessage.From = "pingcastle@gmail.fr"
$mailMessage.To.Add("xxxx@xxxx.fr")
$mailMessage.Subject = "📬 Rapport PingCastle - Semaine $currentWeekNumber"
$mailMessage.IsBodyHtml = $true
$mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

# Construction du HTML avec image inline (en tant que LinkedResource)
$linkedImage = New-Object System.Net.Mail.LinkedResource("$outputPath", "image/png")
$linkedImage.ContentId = "chartImage"
$linkedImage.ContentType.MediaType = "image/png"
$linkedImage.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64

# HTML avec image intégrée via CID
$html = @"
<html>
<body>
<h2>📊 Rapport Hebdomadaire PingCastle</h2>
<h3>Moyennes historiques</h3>
<ul>
    <li>Global : <strong>$avgGlobal</strong></li>
    <li>Stale : <strong>$avgStale</strong></li>
    <li>Privileged : <strong>$avgPrivileged</strong></li>
    <li>Trust : <strong>$avgTrust</strong></li>
    <li>Anomaly : <strong>$avgAnomaly</strong></li>
</ul>
<br>
<h2>Voici le graphique d'évolution des scores :</h2>
<img src='cid:chartImage' style='width:100%; max-width:1000px;' />

<br>
<p style="font-size: 0.9em; color: #888">
Ce message est généré automatiquement. Merci de ne pas y répondre. Pour toute question, contactez le 
<a href="mailto:test@test.fr" style="color:#4FC3F7">Équipe informatique</a>.
</p>
</body>
</html>
"@

# Création de l'alternate view HTML avec ressource liée
$alternateView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($html, $null, "text/html")
$alternateView.LinkedResources.Add($linkedImage)
$mailMessage.AlternateViews.Add($alternateView)

# Configuration SMTP
try {
    $smtpClient = New-Object System.Net.Mail.SmtpClient("SMTP_SERVER", 25)
    $smtpClient.Send($mailMessage)
    Write-Log "Email hebdomadaire envoyé à $to" -Level "INFO"
} catch {
    Write-Log "Erreur lors de l'envoi de l'email hebdomadaire : $_" -Level "ERROR"
}
