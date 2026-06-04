# ⚙️ Konfiguration & Anpassungen

Detaillierte Konfigurationsoptionen und Anpassungsmöglichkeiten.

---

## 🔧 Skript-interne Konfiguration

Im Skript-Header (ca. Zeile 100-120) können folgende Variablen angepasst werden:

### Festplatte & Speicher
```powershell
# Warnung, wenn weniger als 20 GB freier Speicher verfügbar ist
$script:WarningDiskSpaceGB = 20

# Beispiel: Auf 50 GB erhöhen
$script:WarningDiskSpaceGB = 50
```

### Zertifikat-Überwachung
```powershell
# Warnung, wenn Zertifikat in weniger als 30 Tagen abläuft
$script:WarningCertDaysExpiry = 30

# Beispiel: Auf 60 Tage erhöhen (früher warnen)
$script:WarningCertDaysExpiry = 60
```

### Mailbox-Statistiken
```powershell
# Maximum 500 Mailboxen für detaillierte Statistiken
# Bei mehr Mailboxen wird ein Summary angezeigt
$script:MaxMailboxesForStats = 500

# Beispiel: Auf 1000 erhöhen (mehr Details, länger Laufzeit)
$script:MaxMailboxesForStats = 1000
```

### DNS-Server
```powershell
# Leer = Standard-DNS des Systems
$script:DNSServer = ""

# Beispiel: Spezifischen DNS-Server verwenden
$script:DNSServer = "8.8.8.8"

# oder
$script:DNSServer = "dns.contoso.com"
```

---

## 🎨 HTML-Anpassungen

### CSS-Styling
Das Skript generiert CSS mit folgende Farben (ca. Zeile 3000+):

```powershell
# Im Skript suchen nach: $cssStyle = @"
# Dort können CSS-Werte angepasst werden

# Beispiel: Andere Farbe für Header
background-color: #0078D4;     # Microsoft-Blau
→ #009CDE;                     # Helleres Blau
```

### HTML-Header anpassen
```powershell
# DocTitle anpassen (ca. Zeile 85)
$script:DocTitle = "Exchange Server - Umgebungsdokumentation"

# DocSubTitle anpassen (wird über -CompanyName gesetzt)
$script:DocSubTitle = $CompanyName
```

---

## 📋 Sektionen konfigurieren

### Standardsektionen ausschließen
```powershell
# Alle außer Hardware dokumentieren
$excludeSections = @('Hardware')
$sections = $allSections | Where-Object { $_ -notin $excludeSections }

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections $sections
```

### Nur Sicherheit dokumentieren
```powershell
$securitySections = @(
    'Certificates',
    'SecurityAndAuthentication',
    'AntiSpamAntiMalware',
    'Compliance'
)

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections $securitySections
```

---

## 🔐 Sicherheits-Konfiguration

### Netzwerk-Verbindung
Das Skript versucht automatisch (Fallback-Mechanismus):

1. **WsMan (WinRM)** – Port 5985/5986
2. **DCOM/RPC** – Port 135
3. **Remote Registry** – .NET-basiert

```powershell
# Falls DCOM problematisch ist, nur WinRM verwenden
# Im Skript Line 335 anpassen:
# (Manuell Dcom-Code auskommentieren)
```

### Firewall-Anforderungen

**Port 135 (DCOM/RPC):**
```powershell
# Firewall-Regel hinzufügen
New-NetFirewallRule -DisplayName "DCOM-RPC" `
  -Direction Inbound -Protocol tcp -LocalPort 135 `
  -Action Allow
```

**Port 5985/5986 (WinRM):**
```powershell
# WinRM aktivieren
Enable-PSRemoting -Force

# Firewall-Regel automatisch erstellt
```

---

## 🔄 Automation & Scheduling

### Tägliche Dokumentation
```powershell
# File: C:\Scripts\Daily-Exchange-Doc.ps1
$date = Get-Date -Format "yyyyMMdd"
$outputPath = "D:\ExchangeReports\Daily_$date"

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -CompanyName 'Contoso' `
  -OutputPath $outputPath `
  -OutputFormats @('HTML','PDF') `
  -NoGui
```

### Wöchentliche Dokumentation
```powershell
# Scheduled Task erstellen
$scriptPath = "C:\Scripts\Weekly-Exchange-Doc.ps1"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 02:00
$action = New-ScheduledTaskAction -Execute PowerShell `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File $scriptPath"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
Register-ScheduledTask -TaskName "Exchange Weekly Docs" `
  -Trigger $trigger -Action $action -Settings $settings `
  -RunLevel Highest
```

### Monatliche Reports mit E-Mail
```powershell
# File: C:\Scripts\Monthly-Exchange-Report.ps1
$date = Get-Date -Format "yyyyMM"
$outputPath = "D:\ExchangeReports\Monthly_$date"
$pdfFile = "$outputPath\Exchange_Dokumentation_*.pdf"

# Dokumentation erstellen
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03') `
  -OutputPath $outputPath `
  -OutputFormats @('PDF') `
  -NoGui

# Per Mail versenden
Get-ChildItem $pdfFile | Foreach-Object {
    Send-MailMessage -From "automation@contoso.com" `
      -To "admins@contoso.com" `
      -Subject "Exchange Monthly Report - $date" `
      -Body "Siehe Anhang" `
      -Attachments $_.FullName `
      -SmtpServer "smtp.contoso.com"
}
```

---

## 📊 Logging & Error Handling

### Log-Pfad ändern
```powershell
# Standard (in OutputPath)
C:\ExchangeDoku\Exchange-Dokumentation_yyyyMMdd_HHmmss.log

# Custom: Im Skript Line 120 anpassen
$script:LogPath = "D:\CustomLogPath"
```

### Log-Level konfigurieren
```powershell
# Im Skript sind Log-Level verfügbar:
Write-Log -Message "Text" -Level "INFO"      # Informationen
Write-Log -Message "Text" -Level "WARNING"   # Warnungen
Write-Log -Message "Text" -Level "ERROR"     # Fehler
```

---

## 🌍 Multi-Site-Konfiguration

### Separate Dokumentationen per Site
```powershell
# Site 1
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX-Site1-01','EX-Site1-02') `
  -CompanyName 'Contoso - Site 1' `
  -OutputPath 'D:\Reports\Site1'

# Site 2
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX-Site2-01','EX-Site2-02') `
  -CompanyName 'Contoso - Site 2' `
  -OutputPath 'D:\Reports\Site2'
```

### Kombinierte Dokumentation
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX-Site1-01','EX-Site1-02','EX-Site2-01','EX-Site2-02') `
  -CompanyName 'Contoso - Alle Sites' `
  -OutputPath 'D:\Reports\Combined'
```

---

## 📈 Performance-Optimierung

### Schnellere Dokumentation
```powershell
# Nur kritische Sektionen
$fastSections = @(
    'Hardware',
    'ExchangeServer',
    'Certificates',
    'DatabaseAndDAG'
)

time {
  .\Exchange_Dokumentation.ps1 `
    -ExchangeServers @('EX01') `
    -Sections $fastSections
}
```

### Parallele Verarbeitung (mehre Server)
```powershell
# Das Skript verarbeitet Server sequenziell
# Bei vielen Servern länger, aber zuverlässiger
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03','EX04')
```

---

## 🔍 Debugging

### Verbose-Modus
```powershell
.\Exchange_Dokumentation.ps1 -Verbose -NoGui
```

### Debug-Ausgaben
```powershell
$DebugPreference = "Continue"
.\Exchange_Dokumentation.ps1 -Debug -NoGui
```

### Fehlerbehandlung
```powershell
$ErrorActionPreference = "Stop"
try {
    .\Exchange_Dokumentation.ps1 -ExchangeServers @('EX01')
} catch {
    Write-Host "Fehler: $_" -ForegroundColor Red
}
```

---

## 🎯 Nächste Schritte

- ➡️ [Beispiele](Beispiele) – Praktische Anwendungen
- ➡️ [Best Practices](Best-Practices) – Tipps & Tricks
- ➡️ [Troubleshooting](Troubleshooting) – Probleme lösen
