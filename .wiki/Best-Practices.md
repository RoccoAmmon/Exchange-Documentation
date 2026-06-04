# 🎯 Best Practices

Bewährte Praktiken und Optimierungen für die Verwendung des Exchange Dokumentations-Tools.

---

## 📋 Allgemeine Best Practices

### 1. Regelmäßige Dokumentation
**Empfehlung:** Erstellen Sie mindestens **monatlich** ein Baseline-Dokument.

```powershell
# Monatliche Automatisierung
Scheduled Task um 02:00 Uhr nachts
Ausführung jeden 1. Monatag
```

**Vorteil:** Verfügbarkeit von historischen Daten für Vergleiche

---

### 2. Versionierung & Archivierung
**Empfehlung:** Speichern Sie Reports mit **Datumsangaben**.

```powershell
# ✅ GUT: Mit Datum
D:\Reports\Exchange_20260601_0200.html
D:\Reports\Exchange_20260701_0200.html

# ❌ SCHLECHT: Ohne Datum (überschreibt sich)
D:\Reports\Exchange_Report.html
```

**Vorteil:** Ermöglicht Historien-Vergleiche

---

### 3. Separate Verzeichnisse für Sektionen
**Empfehlung:** Organisieren Sie nach Art des Reports.

```powershell
D:\Exchange_Reports\
├─ Daily\          # Schnelle tägliche Checks
├─ Weekly\         # Umfassende wöchentliche Reports
├─ Monthly\        # Detaillierte monatliche Inventare
├─ Audits\         # Compliance & Audit-Reports
├─ Before_Change\  # Pre-Change Baselines
└─ Emergency\      # Notfall-Reports
```

---

### 4. Mehrere Ausgabeformate nutzen
**Empfehlung:** Kombinieren Sie Formate je nach Zweck.

```powershell
# Für Verwaltung: PDF (sicher, unverändert)
-OutputFormats @('PDF')

# Für Techniker: HTML (bearbeitbar, einfach)
-OutputFormats @('HTML')

# Für Wiki/Dokumentation: Markdown
-OutputFormats @('Markdown')

# Für Archivierung: Alle drei
-OutputFormats @('HTML','PDF','Markdown')
```

---

### 5. Filterte Sektionen für schnelle Reports
**Empfehlung:** Verwenden Sie gezielt nur benötigte Sektionen.

```powershell
# Daily Check (schnell, 5 Min)
$sections = @('Hardware','ExchangeServer','Certificates')

# Weekly Review (umfassend, 15 Min)
$sections = @('Hardware','ExchangeServer','Certificates','DatabaseAndDAG','TransportRules')

# Full Audit (komplett, 30+ Min)
# (keine Sections = alle)
```

---

## 🚀 Performance-Optimierungen

### 1. Parallele Ausführungen vermeiden
**Empfehlung:** Führen Sie das Skript **nicht parallel** aus.

```powershell
# ❌ NICHT EMPFOHLEN
$server1 = @('EX01')
$server2 = @('EX02')

.\Script.ps1 -ExchangeServers $server1 &
.\Script.ps1 -ExchangeServers $server2 &

# ✅ EMPFOHLEN: Sequenziell oder kombiniert
.\Script.ps1 -ExchangeServers @('EX01','EX02')
```

**Grund:** Verhindert Ressourcen-Konflikte und Fehler

---

### 2. Off-Peak Hours
**Empfehlung:** Führen Sie umfangreiche Dokumentationen **nachts** aus.

```powershell
# Scheduled Task: 02:00 Uhr nachts
$trigger = New-ScheduledTaskTrigger -Daily -At 02:00
```

**Grund:** Minimales Impact auf Production-Systeme

---

### 3. Remote ausführen
**Empfehlung:** Führen Sie das Skript von einem **Admin-Workstation** aus, nicht vom Exchange-Server.

```powershell
# ✅ GUT: Von Admin-PC/Admin-Server
C:\Scripts\Exchange_Dokumentation.ps1

# ⚠️ VORSICHT: Auf Exchange-Server selbst
# Verursacht Last auf dem System
```

---

## 🔐 Sicherheits-Best Practices

### 1. Sichere Datenspeicherung
**Empfehlung:** Speichern Sie Reports in **sicheren Verzeichnissen**.

```powershell
# Reports enthalten sensible Infos (Zertifikate, IPs, etc)
# → Auf gesicherten Shares speichern

# ACL: Nur Administratoren
icacls D:\Exchange_Reports /grant "DOMAIN\Exchange Admins:F" /inheritance:r
```

---

### 2. E-Mail-Verteilung nur über sichere Kanäle
**Empfehlung:** Reports **nie unverschlüsselt** per E-Mail versenden.

```powershell
# ✅ SICHER: Mit Verschlüsselung
Send-MailMessage -SMTPServer smtp.contoso.com -UseSsl `
  -From admin@contoso.com -To receiver@contoso.com

# ❌ UNSICHER: Unverschlüsselt
Send-MailMessage -SMTPServer smtp.mail.com -From ... # Kein SSL!
```

---

### 3. Zugriffsrechte beschränken
**Empfehlung:** Reports **nur für Administratoren** freigeben.

```powershell
# Nur Administratoren
icacls D:\Exchange_Reports /grant "DOMAIN\Domain Admins:F" `
  /grant "DOMAIN\Exchange Admins:F" /inheritance:r
```

---

## 🛠️ Automatisierungs-Best Practices

### 1. Error Handling
**Empfehlung:** Verwenden Sie robustes Error Handling.

```powershell
param(
    [string]$ExchangeServer = 'EX01',
    [string]$OutputPath = 'D:\Reports'
)

try {
    .\Exchange_Dokumentation.ps1 `
      -ExchangeServers @($ExchangeServer) `
      -OutputPath $OutputPath `
      -NoGui -ErrorAction Stop
    
    Write-EventLog -LogName "Application" -Source "Exchange-Docs" `
      -EventId 1000 -Message "Erfolg: Report erstellt" -Type Information
}
catch {
    Write-EventLog -LogName "Application" -Source "Exchange-Docs" `
      -EventId 1001 -Message "Fehler: $_" -Type Error
    
    # Alert senden
    Send-MailMessage -To admins@contoso.com -Subject "Exchange-Docs Fehler" `
      -Body "Fehler: $_"
}
```

---

### 2. Monitoring & Alerts
**Empfehlung:** Überwachen Sie Fehler und Warnungen.

```powershell
# Log-Datei prüfen auf Fehler
$logFile = Get-ChildItem "C:\ExchangeDoku\*.log" | 
  Sort-Object LastWriteTime -Descending | Select-Object -First 1

$errors = Select-String -Path $logFile.FullName -Pattern "\[ERROR\]"

if ($errors) {
    # Alert
    Send-MailMessage -To admins@contoso.com -Subject "⚠️ Exchange-Docs Fehler" `
      -Body ($errors | Out-String)
}
```

---

### 3. Cleanup alte Reports
**Empfehlung:** Alte Reports automatisch **archivieren/löschen**.

```powershell
# Reports älter als 90 Tage in ZIP archivieren und löschen
Get-ChildItem D:\Exchange_Reports -Filter "*.pdf" | 
  Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } |
  ForEach-Object {
    $zipFile = "$($_.DirectoryName)\Archive_$($_.BaseName).zip"
    Compress-Archive -Path $_.FullName -DestinationPath $zipFile
    Remove-Item $_.FullName
  }
```

---

## 📊 Reporting-Best Practices

### 1. Konsistente Nomenklatur
**Empfehlung:** Verwenden Sie einheitliche Namensgebung.

```powershell
# Format: [Firma]_[Umgebung]_[Typ]_[Datum]
Exchange_Contoso_Prod_Daily_20260604.pdf
Exchange_Contoso_Audit_Monthly_202606.pdf
Exchange_Contoso_PreUpgrade_20260604_1400.pdf
```

---

### 2. Zusätzliche Informationen erfassen
**Empfehlung:** Dokumentieren Sie **warum** und **wann** Reports erstellt wurden.

```powershell
# Zusatz-Info in Datei speichern
$info = @{
    CreatedDate = Get-Date
    Reason = "Monatliche Inventur"
    ExecutedBy = $env:USERNAME
    ComputerName = $env:COMPUTERNAME
    ExchangeServers = @('EX01','EX02')
}

$info | ConvertTo-Json | Out-File "D:\Reports\Metadata.json"
```

---

### 3. Versionskontrolle für kritische Reports
**Empfehlung:** Wichtige Reports in Git speichern.

```powershell
cd D:\Exchange_Reports
git init
git add *.pdf *.html
git commit -m "Exchange Reports $(Get-Date -Format 'yyyyMMdd')"
git tag "report-$(Get-Date -Format 'yyyyMMdd')"
```

---

## 📈 Skalierung

### 1. Große Umgebungen (50+ Server)
**Empfehlung:** Teilen Sie nach **Standorten** auf.

```powershell
# Dokumentation pro Standort
$sites = @{
    'Berlin' = @('EX-BER-01','EX-BER-02')
    'München' = @('EX-MUC-01','EX-MUC-02')
    'Köln' = @('EX-KOL-01','EX-KOL-02')
}

foreach ($site in $sites.GetEnumerator()) {
    .\Exchange_Dokumentation.ps1 `
      -ExchangeServers $site.Value `
      -CompanyName "Contoso - $($site.Key)" `
      -OutputPath "D:\Reports\$($site.Key)" `
      -NoGui
}
```

---

### 2. Hybrid-Umgebungen
**Empfehlung:** Dokumentieren Sie **separate** On-Premises & Online.

```powershell
# On-Premises
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -CompanyName 'Contoso - On-Premises' `
  -OutputPath 'D:\Reports\OnPrem'

# Hybrid-Status (zusätzlich in anderem Tool oder Tool)
```

---

## ✅ Checkliste

- [ ] **Regelmäßige Dokumentation** – Mind. monatlich
- [ ] **Versionierung** – Mit Datumsstempel
- [ ] **Sicherung** – Reports sichern
- [ ] **Zugriffskontrolle** – Nur für Admins
- [ ] **Error Handling** – Logging & Alerts
- [ ] **Automatisierung** – Scheduled Tasks
- [ ] **Cleanup** – Alte Reports archivieren
- [ ] **Monitoring** – Log-Dateien überwachen
- [ ] **Dokumentation** – Warum, Wann, Wer
- [ ] **Testen** – Im Test-Umfeld testen vor Prod

---

## 🎯 Nächste Schritte

- ➡️ [FAQ](FAQ) – Häufige Fragen
- ➡️ [Troubleshooting](Troubleshooting) – Probleme lösen
- ➡️ [Beispiele](Beispiele) – Mehr Szenarien
