
# Pfad zur Quelltextdatei aus Archivexport NLS
$sourceFilePath = "\\SGWVDSFS01.netzbetrieb.local\Share_NLS$\Excel-Reports\Auswertungen\Betriebstagebuch_3Tage.csv"

# Pfad zu den Zieltxt-Dateien
$sourceFilePath_time = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\temp\Betriebstagebuch_3Tage_time.csv"
$targetFilePath = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\temp\Betriebstagebuch_3Tage.txt"    #Tempdatei
$sourceFilePath_time = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\temp\Betriebstagebuch_3Tage_time.csv"
$targetCsvFilePath = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\temp\Betriebstagebuch_3Tage.csv" #Tempdatei
$csvSavePathHtml = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\BT_3Tage.html"


################################################################################################################################

# Pfad zur Empfängerdatei
$filePath = "\\tsgswvwi01.netzbetrieb.local\data$\Daten\HL\Betriebstagebuch\PS Scripts\Empfänger.txt"

# Variablen löschen
Get-Variable | Where-Object { !$_.Options -match 'ReadOnly' } | ForEach-Object { Remove-Variable -Name $_.Name -Force }


# Korrektur Sommmerzeit
# Zeitzone und Datumformat festlegen
$zurichTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById("W. Europe Standard Time")
$dateTimeFormat = "dd.MM.yyyy HH:mm:ss.fff"

# Inhalte der Datei einlesen
$content = Get-Content -Path $sourceFilePath

# Header und erste Kommentarzeilen beibehalten
$processedLines = $content[0..7]

# Durchgehen der Inhalte ab der ersten Datenzeile
foreach ($line in $content[8..$content.Length]) {
    $fields = $line -split ","  # Zerlegen der Zeile in Felder
    
    foreach ($i in 1,2) {
        if ($fields[$i] -match "\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}\.\d{3}") {
            $dateTime = [DateTime]::ParseExact($fields[$i], $dateTimeFormat, [Globalization.CultureInfo]::InvariantCulture)

            # Keine Konvertierung, wenn die Zeit bereits in lokaler Zeitzone (Zürich) vorliegt
            if ($zurichTimeZone.IsDaylightSavingTime($dateTime)) {
                $dateTime = $dateTime.AddHours(1)  # Nur Anpassung um eine Stunde bei Sommerzeit
            }

            $fields[$i] = $dateTime.ToString($dateTimeFormat)
        }
    }
    
    # Aktualisierte Zeile zu den verarbeiteten Zeilen hinzufügen
    $processedLines += ($fields -join ",")
}

$processedLines | Set-Content -Path $sourceFilePath
# Ende Korrektur Sommmerzeit

# Anzahl der Zeilen, die entfernt werden sollen
$anzahlZeilenEntfernen = 7

# Archivexport NLS einlesen mit der richtigen Kodierung # Die ersten 7 Zeilen entfernen # Die bereinigten Daten in eine Test Datei speichern mit der richtigen Kodierung
$lines = Get-Content -Path $sourceFilePath -Encoding Default
$lines = $lines | Select-Object -Skip $anzahlZeilenEntfernen

$lines | Set-Content -Path $targetFilePath -Encoding UTF8

#Write-Host "Die ersten $anzahlZeilenEntfernen Zeilen wurden entfernt und die bereinigten Daten wurden als Betriebstagebuch1.txt gespeichert."




# Initialisieren Sie die Zählvariablen
$AnzBetriebt = 0
$AnzAlarmt = 0
$AnzStoert = 0
$AnzWarnungt = 0

$AnzBetrieb = 0
$AnzAlarm = 0
$AnzStoer = 0
$AnzWarnung = 0





# Lesen Sie die Datei ein und überspringen Sie die erste Zeile mit den Spaltenköpfen
$content = Get-Content -Path $targetFilePath | Select-Object -Skip 1

# Prüfen Sie, ob die Datei Einträge enthält
if ($content.Count -gt 0) {
    # Extrahiere Datum und Uhrzeit ohne Sekunden für 'Von' (erster LSZeitpunkt)
    $ersterLSZeitpunkt = ($content[0] -split ",")[1].Trim()
    $Von = $ersterLSZeitpunkt.Substring(0, 16)

    # Extrahiere Datum und Uhrzeit ohne Sekunden für 'Bis' (letzter LSZeitpunkt)
    $letzterLSZeitpunkt = ($content[-1] -split ",")[1].Trim()
    $Bis = $letzterLSZeitpunkt.Substring(0, 16)
} else {
    Write-Host "Die Datei enthält keine Einträge."
    return
}

# Schleife durch jede Zeile in der Datei
foreach ($line in $content) {
    # Trennen Sie die Zeile an den Kommas und erstellen Sie ein Array aus den Spalten
    $columns = $line -split ","

    # Zähle, wie oft die Werte in der Spalte 'Klasse' erscheinen
    switch ($columns[8]) {
        "Betrieb" { $AnzBetriebt++ }
        "Alarm"   { $AnzAlarmt++ }
        "Stoer"   { $AnzStoert++ }
        "Warnung" { $AnzWarnungt++ }
    }
}









# Von und Bis mit Wochentag ergänzen

try {
    # Parsen des ersten Zeitpunktes
    $VonDateTime = [DateTime]::ParseExact($ersterLSZeitpunkt, "dd.MM.yyyy HH:mm:ss.fff", [Globalization.CultureInfo]::InvariantCulture)
    # Verwendung der deutschen Kultur
    $deCulture = New-Object System.Globalization.CultureInfo "de-DE"
    $Von = $VonDateTime.ToString("ddd dd.MM.yyyy HH:mm", $deCulture)
} catch {
    Write-Host "Es gab ein Problem beim Parsen des 'Von'-Datums: $_"
}

try {
    # Parsen des letzten Zeitpunktes
    $BisDateTime = [DateTime]::ParseExact($letzterLSZeitpunkt, "dd.MM.yyyy HH:mm:ss.fff", [Globalization.CultureInfo]::InvariantCulture)
    # Verwendung der deutschen Kultur
    $deCulture = New-Object System.Globalization.CultureInfo "de-DE"
    $Bis = $BisDateTime.ToString("ddd dd.MM.yyyy HH:mm", $deCulture)
} catch {
    Write-Host "Es gab ein Problem beim Parsen des 'Bis'-Datums: $_"
}


# Ausgabe der Variablen
Write-Host "Von (erster LSZeitpunkt): $Von"
Write-Host "Bis (letzter LSZeitpunkt): $Bis"






# Zeilen mit den gewünschten Spaltenüberschriften
$spaltenUeberschriften = "StatusBZ", "Prio", "LSZeitpunkt", "QuitZeitpunkt", "QuitUser", "AktionUser", "Text", "Klasse"  # Spalten und Reihenfolge hier setzen

# CSV einlesen mit der richtigen Kodierung
$lines = Get-Content -Path $targetFilePath -Encoding UTF8

# Spaltenüberschriften aus der ersten Zeile extrahieren
$headerLine = $lines[0] -split ','

# Indexe der Spaltenüberschriften ermitteln
$spaltenIndexe = @{}
foreach ($spalte in $spaltenUeberschriften) {
    $spaltenIndexe[$spalte] = [array]::IndexOf($headerLine, $spalte)
}

# Die Zeilen mit den gewünschten Spaltenüberschriften auswählen und in Betriebstagebuch2.csv speichern
$csvData = foreach ($line in $lines | Select-Object -Skip 1) {
    $values = $line -split ','
    $selectedValues = @{}
    foreach ($spalte in $spaltenUeberschriften) {
        $selectedValues[$spalte] = $values[$spaltenIndexe[$spalte]]
    }

    # Hinzugefügte benutzerdefinierte Spalte "Prio" basierend auf der "Klasse" Spalte
    $selectedValues["Prio"] = if ($selectedValues["Klasse"] -eq "Betrieb") { "BT" }
                              elseif ($selectedValues["Klasse"] -eq "Warn") { "U2" }
                              elseif ($selectedValues["Klasse"] -eq "Alarm") { "K" }
                              elseif ($selectedValues["Klasse"] -eq "Stoer") { "U1" }
                              else { "" }

    # Hinzufügen einer leeren "StatusBZ"-Spalte
    $selectedValues["StatusBZ"] = ""

    [PSCustomObject]$selectedValues
}

# Sortieren Sie die Daten nach LSZeitpunkt absteigend
$csvData = $csvData | Sort-Object {[datetime]::ParseExact($_.LSZeitpunkt, 'dd.MM.yyyy HH:mm:ss.fff', $null)} -Descending

# Initialstatus für Spalte StatusBZ festlegen
$initialStatus = $null

# Zähler für die Prioritätsklassen initialisieren
$AnzAlarm = 0
$AnzStoer = 0
$AnzWarnung = 0

# Durchlaufe die Daten einmal, um den Anfangsstatus StatusB festzustellen
foreach ($row in $csvData) {
    if ($null -eq $initialStatus) {
        if ($row.Text -match 'Status anwesend') {
            $initialStatus = 'anwesend'
            break # Beende die Schleife, sobald der Anfangsstatus festgestellt wurde
        } elseif ($row.Text -match 'Status abwesend') {
            $initialStatus = 'abwesend'
            break # Beende die Schleife, sobald der Anfangsstatus festgestellt wurde
        }
    }
}

# Setze den initialen Status StatusBZ, falls keiner gefunden wurde
if ($null -eq $initialStatus) {
    $initialStatus = 'abwesend' # oder 'anwesend', abhängig von der Geschäftslogik
}

# Setze den Anfangsstatus StatusB als aktuellen Status
$currentStatus = $initialStatus

# Durchlaufe die Daten und setze den Wert für StatusBZ basierend auf dem Text
foreach ($row in $csvData) {
    # Prüfe auf Statusänderung in der Zeile
    if ($row.Text -match 'Status anwesend') {
        $currentStatus = 'abwesend'
    } elseif ($row.Text -match 'Status abwesend') {
        $currentStatus = 'anwesend'
    }

    # Setze den StatusBZ für die aktuelle Zeile
    $row.StatusBZ = $currentStatus

    # Zähle die Prioritätsklassen, wenn der aktuelle Status 'abwesend' ist
    if ($currentStatus -eq 'abwesend') {
        switch ($row.Prio) {
            'K'  { $AnzAlarm++ }
            'U1' { $AnzStoer++ }
            'U2' { $AnzWarnung++ }
        }
    }
}

# Exportiere die aktualisierten Daten in die CSV-Datei mit der richtigen Kodierung
$csvData | Export-Csv -Path $targetCsvFilePath -Delimiter ',' -NoTypeInformation -Encoding UTF8

#Write-Host "Die Daten wurden in Betriebstagebuch2.csv gespeichert."

# CSV-Inhalt einlesen
$csvData = Import-Csv $targetCsvFilePath -Encoding UTF8


# Definieren der Anfang des HTML-Strings mit Styles für bedingte Formatierung
$html = @"
<style>
  th { 
    background-color: black; 
    color: white; 
    font-weight: bold; 
  }
  td { 
    background-color: #f2f2f2; 
    padding: 5px; 
  }
  .k  { background-color: red;    color: black; font-weight: bold; }
  .u1 { background-color: yellow; color: black; font-weight: bold; }
  .u2 { background-color: white;  color: black; font-weight: bold; }
  .bt { background-color: black;  color: white; font-weight: bold; }
  .abw { background-color: blue;   color: white; font-weight: bold; }
  .anw { background-color: green;  color: white; font-weight: bold; }
</style>
"@



# Einfügen Texte vor der Tabelle
#$html += "<h2>Betriebstagebuch</h2>"#
$html += '<h4>Alarme in Abwesenheit&nbsp;&nbsp;&nbsp;<span class="k">K:</span>' +'&nbsp;'+ $AnzAlarm +'&nbsp;&nbsp;'+ ' <span class="u1">U1:</span>'+'&nbsp;' + $AnzStoer +'&nbsp;&nbsp;'+ ' <span class="u2">U2:</span>'+'&nbsp;' + $AnzWarnung +'&nbsp;&nbsp;'+ '</h4>';
$html += '<h5>' + $Von + '&nbsp;&nbsp;&nbsp;bis&nbsp;&nbsp;&nbsp' + $Bis + '</h5>'
#$html += '<h4>Alarme in Abwesenheit&nbsp;&nbsp;&nbsp;<span class="k">K:</span>' +'&nbsp;'+ $AnzAlarm +'&nbsp;&nbsp;'+ ' <span class="u1">U1:</span>'+'&nbsp;' + $AnzStoer +'&nbsp;&nbsp;'+ ' <span class="u2">U2:</span>'+'&nbsp;' + $AnzWarnung +'&nbsp;&nbsp;'+ '&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; Alarme Total&nbsp;&nbsp;&nbsp; <span class="k">K:</span>'+'&nbsp;' + $AnzAlarmt +'&nbsp;&nbsp;'+ ' <span class="u1">U1:</span>' +'&nbsp;'+ $AnzStoert +'&nbsp;&nbsp;'+ ' <span class="u2">U2:</span>'+'&nbsp;' + $AnzWarnungt + '</h4>';






# Start der Tabelle
$html += "<table border='1'>
<tr>
"

# Hinzufügen der Spaltenüberschriften zur Tabelle
foreach ($column in $spaltenUeberschriften) {
    $html += "<th>$column</th>"
}

# Schließen des Tabellenkopfes
$html += "</tr>"

# Hinzufügen der Datenzeilen zur HTML-Tabelle
foreach ($row in $csvData) {
    $html += "<tr>"
    foreach ($column in $spaltenUeberschriften) {
        # Initial keine Klasse setzen
        $styleClass = ""
        
        # Bedingte Formatierung für die 'Prio', 'Bereich' und 'Text'-Spalten
        switch ($column) {
            'Prio' {
                switch ($row.Prio) {
                    'K'  { $styleClass = " class='k'" }
                    'U1' { $styleClass = " class='u1'" }
                    'U2' { $styleClass = " class='u2'" }
                    'BT' { $styleClass = " class='bt'" }
                }
            }
            'Bereich' {
                switch ($row.Bereich) {
                    'abw' { $styleClass = " class='abw'" }
                    'anw' { $styleClass = " class='anw'" }
                }
            }
            'Text' {
                if ($row.Text -match 'Status anwesend') {
                    $styleClass = " class='anw'"
                } elseif ($row.Text -match 'Status abwesend') {
                    $styleClass = " class='abw'"
                }
            }
            'StatusBZ' {
                if ($row.StatusBZ -match 'anwesend') {
                    $styleClass = " class='anw'"
                } elseif ($row.StatusBZ -match 'abwesend') {
                    $styleClass = " class='abw'"
                }
            }
        }

        # Hinzufügen der Zelle zur Reihe
        $html += "<td$styleClass>$($row.$column)</td>"
    }
    $html += "</tr>"
}

# Schließen des HTML-Strings
$html += "</table>"

# Ausgabe oder Speichern des HTML-Codes nach Bedarf


# Speichern des HTML-Strings in einer HTML-Datei
$html | Out-File -FilePath $csvSavePathHtml -Encoding UTF8

# Ausgabe in der Konsole, dass die Datei erstellt wurde
Write-Host "Die HTML-Datei wurde erstellt und nach LSZeitpunkt absteigend sortiert."



# Ausgabe der Zählvariablen und der formatierten LSZeitpunkt Variablen

Write-Host "Von (erster LSZeitpunkt): $Von"
Write-Host "Bis (letzter LSZeitpunkt): $Bis"
Write-Host "Anzahl Betrieb: $AnzBetrieb"
Write-Host "Anzahl Alarm: $AnzAlarm"
Write-Host "Anzahl Störung: $AnzStoer"
Write-Host "Anzahl Warnung: $AnzWarnung"
Write-Host "------------------------------"

Write-Host "Total"
Write-Host "Anzahl Betrieb: $AnzBetriebt"
Write-Host "Anzahl Alarm: $AnzAlarmt"
Write-Host "Anzahl Störung: $AnzStoert"
Write-Host "Anzahl Warnung: $AnzWarnungt"










# Datei lesen und jede Zeile in ein Array einlesen
$recipients = Get-Content -Path $filePath

# Ausgabe der Empfänger
foreach ($recipient in $recipients) {
    Write-Host $recipient
}


# SMTP-Serverdetails
$smtpServer = "smtprelay.net.work"
$smtpPort = 25
$sender = "nls@sgsw.ch"
#$recipient = "christian.angele@sgsw.ch"

# E-Mails senden
foreach ($recipient in $recipients) {
    # Erstellen der E-Mail-Parameter mit HTML-Body für jeden Empfänger
    $mailParams = @{
        SmtpServer = $smtpServer
        Port = $smtpPort
        From = $sender
        To = $recipient        
        Subject = "Betriebstagebuch"
        Body = $html
        BodyAsHtml = $true
    }

    # Senden der E-Mail
    Send-MailMessage @mailParams
}