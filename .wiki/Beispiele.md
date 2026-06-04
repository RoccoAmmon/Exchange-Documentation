# 💡 Praktische Beispiele

10+ praktische Verwendungsszenarien und Anwendungsfälle.

---

## 📌 Verwendungsszenarios

### Szenario 1: Schnelle Inventur eines Servers

Sie möchten schnell eine Übersicht des aktuellen Exchange-Servers bekommen.

```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Schnelle Inventur' `
  -OutputPath 'C:\Temp' `
  -NoGui

# Ergebnis: C:\Temp\Exchange_Dokumentation_yyyyMMdd_HHmmss.html
```

**Dauer:** ~5-10 Minuten

---

### Szenario 2: Komplette Exchange-Umgebung dokumentieren

Vollständige Dokumentation aller Exchange-Server einer Organisation.

```powershell
$allServers = Get-ExchangeServer | Select-Object -ExpandProperty Name

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers $allServers `
  -CompanyName 'Contoso GmbH - Vollständige Inventur' `
  -OutputPath 'D:\Exchange_Inventory' `
  -OutputFormats @('HTML','PDF') `
  -NoGui
```

**Dauer:** ~30-60 Minuten (je nach Anzahl Server)

---

### Szenario 3: Audit & Compliance-Report

Erstellen Sie einen detaillierten Report für Audit- oder Compliance-Anforderungen.

```powershell
$complianceSections = @(
    'Certificates',
    'SecurityAndAuthentication',
    'AntiSpamAntiMalware',
    'Compliance',
    'RBAC​Configuration'
)

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -CompanyName 'Contoso GmbH - Compliance Audit 2026' `
  -Sections $complianceSections `
  -OutputPath 'D:\Compliance_Reports' `
  -OutputFormats @('PDF') `
  -NoGui
```

**Use Case:** ISO 27001, SOC2, GDPR Audits

---

### Szenario 4: Tägliche Überwachung

Automatische tägliche Dokumentation für Monitoring.

**Datei:** `C:\Scripts\Daily-Exchange-Monitoring.ps1`

```powershell
param(
    [string]$ServerName = 'EX01',
    [string]$ReportPath = 'D:\DailyReports'
)

$date = Get-Date -Format "yyyyMMdd"
$dailyPath = "$ReportPath\$date"

# Verzeichnis erstellen
if (-not (Test-Path $dailyPath)) {
    New-Item -ItemType Directory -Path $dailyPath | Out-Null
}

# Dokumentation (nur Hardware, Zertifikate, Datenbanken)
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @($ServerName) `
  -CompanyName "Tägliche Überwachung - $ServerName" `
  -Sections @('Hardware','Certificates','DatabaseAndDAG') `
  -OutputPath $dailyPath `
  -NoGui

Write-Host "Report erstellt: $dailyPath"
```

**Scheduled Task:**
```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At 08:00
$action = New-ScheduledTaskAction -Execute PowerShell `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\Scripts\Daily-Exchange-Monitoring.ps1"
Register-ScheduledTask -TaskName "Daily Exchange Report" -Trigger $trigger -Action $action -RunLevel Highest
```

---

### Szenario 5: Monatliche Management-Reports

Erstellen Sie professionelle monatliche Reports für das Management.

```powershell
$month = Get-Date -Format "MMMM yyyy"
$outputDir = "D:\Monthly_Reports\$(Get-Date -Format 'yyyyMM')"

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03') `
  -CompanyName "Contoso - Monthly Exchange Report - $month" `
  -OutputPath $outputDir `
  -OutputFormats @('HTML','PDF') `
  -NoGui

# PDF per Mail versenden
$pdfFile = Get-ChildItem "$outputDir\*.pdf" | Select-Object -First 1
Send-MailMessage -From automation@contoso.com `
  -To "management@contoso.com" `
  -Subject "Exchange Monthly Report - $month" `
  -Body "Siehe Anhang für detaillierten Report" `
  -Attachments $pdfFile.FullName `
  -SmtpServer smtp.contoso.com
```

---

### Szenario 6: Vor Software-Updates

Erstellen Sie einen Baseline-Report vor wichtigen Updates.

```powershell
$baselineDate = Get-Date -Format "yyyyMMdd_HHmm"
$outputDir = "D:\Baselines\Pre_Update_$baselineDate"

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -CompanyName "Exchange Baseline vor Update - $baselineDate" `
  -OutputPath $outputDir `
  -OutputFormats @('HTML','PDF') `
  -NoGui

Write-Host "Baseline-Report erstellt: $outputDir"
Write-Host "Nach dem Update denselben Befehl mit neuem Datum ausführen und Unterschiede vergleichen"
```

---

### Szenario 7: Notfall-Dokumentation

Schnelle Dokumentation im Notfall (z.B. beim Ausfall).

```powershell
# Nur kritische Infos (schnell!)
$emergencySections = @(
    'ExchangeServer',
    'DatabaseAndDAG',
    'Certificates'
)

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'NOTFALL: Exchange Status' `
  -Sections $emergencySections `
  -OutputPath 'C:\Emergency_Reports' `
  -NoGui
```

**Dauer:** ~2-3 Minuten

---

### Szenario 8: Hybrid-Migration

Dokumentieren Sie Exchange vor einer Hybrid- oder Cloud-Migration.

```powershell
$allServers = Get-ExchangeServer | Select-Object -ExpandProperty Name

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers $allServers `
  -CompanyName 'Contoso - Pre-Migration Inventory' `
  -OutputPath 'D:\Migration_Docs' `
  -OutputFormats @('HTML','PDF','Markdown') `
  -NoGui

Write-Host "Migration-Dokumentation abgeschlossen"
Write-Host "Diese Dokumentation ist wichtig für die Migration!"
```

---

### Szenario 9: Vergleich vor/nach Änderung

Dokumentieren Sie vor und nach einer Konfigurationsänderung.

```powershell
# VOR der Änderung
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Pre-Config-Change' `
  -OutputPath 'D:\Before_Change' `
  -NoGui

Write-Host "VOR-Dokumentation abgeschlossen. Bitte Änderungen durchführen..."
Read-Host "Nach Änderung Enter drücken"

# NACH der Änderung
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Post-Config-Change' `
  -OutputPath 'D:\After_Change' `
  -NoGui

Write-Host "NACH-Dokumentation abgeschlossen"
Write-Host "Bitte beide Reports vergleichen!"
```

---

### Szenario 10: Archivierung für Dokumentation

Archivieren Sie regelmäßig alle Reports.

```powershell
$archivePath = "D:\Exchange_Documentation_Archive"
$reportDate = Get-Date -Format "yyyyMMdd"

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03','EX04') `
  -CompanyName "Archiv $reportDate" `
  -OutputPath "$archivePath\$reportDate" `
  -OutputFormats @('HTML','PDF') `
  -NoGui

# Alte Reports komprimieren (älter als 30 Tage)
Get-ChildItem "$archivePath" -Directory | 
  Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
  ForEach-Object {
    Compress-Archive -Path $_.FullName -DestinationPath "$_.zip" -CompressionLevel Optimal
    Remove-Item $_.FullName -Recurse
  }
```

---

### Szenario 11: Mehrsprachige Berichte

Dokumentation mit unterschiedlichen Firmennamen.

```powershell
# Deutsch
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Contoso GmbH (Deutsch)' `
  -OutputPath 'D:\Reports_DE' `
  -NoGui

# English
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Contoso Inc (English)' `
  -OutputPath 'D:\Reports_EN' `
  -NoGui
```

---

## 🎯 PowerShell-Snippets

### Alle Reports des heutigen Tages auflisten
```powershell
$today = Get-Date -Format "yyyyMMdd"
Get-ChildItem "D:\ExchangeReports" -Filter "*$today*" -Recurse
```

### Größe der letzten 5 Reports
```powershell
Get-ChildItem "D:\ExchangeReports" -Filter "*.pdf" | 
  Sort-Object -Property LastWriteTime -Descending | 
  Select-Object -First 5 -Property Name, @{N='Size (MB)';E={[math]::Round($_.Length/1MB,2)}}
```

### Report öffnen nach Erstellung
```powershell
$result = .\Exchange_Dokumentation.ps1 -ExchangeServers @('EX01') -NoGui
$reportFile = Get-ChildItem "C:\ExchangeDoku" -Filter "*.html" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Start-Process $reportFile.FullName
```

---

## 🎯 Nächste Schritte

- ➡️ [Best Practices](Best-Practices) – Tipps für häufige Szenarien
- ➡️ [Troubleshooting](Troubleshooting) – Probleme lösen
- ➡️ [FAQ](FAQ) – Häufige Fragen
