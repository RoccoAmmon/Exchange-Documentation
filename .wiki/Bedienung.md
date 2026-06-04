# 🎮 Bedienung & Parameter

Umfassende Anleitung zur Verwendung des Exchange Dokumentations-Tools.

---

## 📝 Parameter-Übersicht

### Grundlegende Parameter

| Parameter | Typ | Standard | Beschreibung | Beispiel |
|---|---|---|---|---|
| `-ExchangeServers` | string[] | (GUI) | Zu dokumentierende Server | `@('EX01','EX02')` |
| `-CompanyName` | string | Meine Organisation | Firmenname im Bericht | `'Contoso GmbH'` |
| `-OutputPath` | string | C:\ExchangeDoku | Ausgabeverzeichnis | `'D:\Reports'` |
| `-OutputFormats` | string[] | @('HTML') | Ausgabeformate | `@('HTML','PDF')` |
| `-Sections` | string[] | (alle) | Dokumentationsbereiche | Siehe unten |
| `-ShowGui` | switch | - | GUI erzwingen | `-ShowGui` |
| `-NoGui` | switch | - | GUI unterdrücken | `-NoGui` |

---

## 🎯 Verwendungsbeispiele

### Beispiel 1: Einfach mit GUI
```powershell
.\Exchange_Dokumentation.ps1
```
→ Öffnet grafische Oberfläche zur Server-Auswahl

### Beispiel 2: Einen Server dokumentieren
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Contoso GmbH' `
  -NoGui
```

### Beispiel 3: Mehrere Server
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03') `
  -CompanyName 'Meine Firma' `
  -OutputPath 'D:\Inventory' `
  -NoGui
```

### Beispiel 4: Alle Ausgabeformate
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -OutputFormats @('HTML','PDF','Markdown') `
  -OutputPath 'C:\Reports'
```

### Beispiel 5: Nur bestimmte Sektionen
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections @('Hardware','ExchangeServer','Certificates') `
  -CompanyName 'Test'
```

### Beispiel 6: Automatisiert (Scheduled Task)
```powershell
# Script: C:\Scripts\DocumentExchange.ps1
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -CompanyName 'Contoso' `
  -OutputPath 'D:\WeeklyReports\Exchange_$(Get-Date -Format yyyyMMdd)' `
  -OutputFormats @('HTML','PDF') `
  -NoGui

# Errorcodes prüfen
if ($LASTEXITCODE -ne 0) {
    Send-MailMessage -To admin@contoso.com -Subject "Exchange Dokumentation fehlgeschlagen"
}
```

---

## 📊 Dokumentationssektionen

### Verfügbare Sektionen

```powershell
$availableSections = @(
    'Hardware',                    # Hardware & System-Infos
    'PatchInfo',                   # Windows & Exchange Patches
    'FSMORoles',                   # Active Directory FSMO-Rollen
    'ADInfo',                      # AD-Informationen & Schema
    'ExchangeServer',              # Exchange Server Übersicht
    'OrganizationConfig',          # Exchange Org-Konfiguration
    'ExchangeURLs',                # URLs & virtuelle Verzeichnisse
    'DatabaseAndDAG',              # Datenbanken & DAG
    'PublicFolders',               # Öffentliche Ordner
    'ConnectorsAndRouting',        # Send-/Receive-Connectoren
    'RemoteDomains',               # Remote Domains & Accepted Domains
    'DNSAndEmailPolicies',         # DNS-Records & E-Mail Policies
    'AddressLists',                # Adresslisten & OAB
    'MobileDevicePolicies',        # Mobile Device Management
    'Certificates',                # SSL/TLS Zertifikate
    'TransportRules',              # Transport-Regeln
    'MailboxStatistics',           # Mailbox-Statistiken
    'ThrottlingPolicies',          # Throttling-Richtlinien
    'RetentionPolicies',           # Retention & Journal Rules
    'RBACConfiguration',           # RBAC-Rollen
    'SecurityAndAuthentication',   # Sicherheit & TLS
    'AntiSpamAntiMalware',        # Anti-Spam/Anti-Malware
    'Compliance',                  # DLP, Litigation Hold, etc.
    'SMTPRelay',                   # SMTP-Relay-Konfiguration
    'EventLogs'                    # Event Logs (letzte 7 Tage)
)
```

### Beispiel: Nur wichtige Sektionen
```powershell
$sections = @(
    'Hardware',
    'ExchangeServer',
    'Certificates',
    'DatabaseAndDAG',
    'SecurityAndAuthentication'
)

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections $sections `
  -OutputPath 'C:\QuickReport'
```

---

## 🖥️ Grafische Benutzeroberfläche (GUI)

Die GUI bietet folgende Möglichkeiten:

### Oberfläche
```
┌─────────────────────────────────────────────┐
│ Exchange Server Dokumentation v1.0          │
├─────────────────────────────────────────────┤
│                                             │
│ [HEADER] Automatische Dokumentation         │
│                                             │
│ ☑ EX01 (erkannt)                           │
│ ☑ EX02 (erkannt)                           │
│ ☐ Weitere Server hinzufügen...             │
│                                             │
│ Firmenname: [____________Contoso GmbH]     │
│ Ausgabepfad: [______C:\ExchangeDoku]       │
│                                             │
│ ☑ HTML  ☐ PDF  ☐ Markdown                  │
│                                             │
│ [Sektionen wählen...] [Alle/Keine]         │
│                                             │
│           [Start] [Abbrechen]               │
│                                             │
└─────────────────────────────────────────────┘
```

### GUI-Features
- ✅ Automatische Exchange-Server-Erkennung
- ✅ Sektionen nach Kategorie gruppiert
- ✅ Schnelle Auswahloptionen (Alle/Keine)
- ✅ Livefortschritt während Dokumentation
- ✅ Optionales Öffnen des Ergebnisses

---

## 🛠️ Konfigurierbare Variablen

Im Skript (Anfang) können angepasst werden:

```powershell
# Festplattenwarnungs-Schwelle
$script:WarningDiskSpaceGB = 20

# Zertifikat-Ablauf-Warnung (Tage)
$script:WarningCertDaysExpiry = 30

# Limit für detaillierte Mailbox-Statistiken
$script:MaxMailboxesForStats = 500

# DNS-Server für MX-Abfragen (leer = Standard)
$script:DNSServer = ""
```

---

## 📤 Ausgabeformate

### HTML (Standard)
- ✅ Formatiert mit CSS
- ✅ Inhaltsverzeichnis
- ✅ Direkt in Word importierbar
- ✅ Seitennummerierung für Druck

```powershell
-OutputFormats @('HTML')
```

### PDF
- ✅ Professionelle Ausgabe
- ✅ Unveränderbar
- ✅ Einfaches Verteilen

```powershell
-OutputFormats @('PDF')
```

### Markdown
- ✅ GitHub-kompatibel
- ✅ Wiki-freundlich
- ✅ Versionskontrolle-freundlich

```powershell
-OutputFormats @('Markdown')
```

---

## 🔍 Debugging & Logging

### Verbose-Ausgabe
```powershell
.\Exchange_Dokumentation.ps1 -Verbose -NoGui
```

### Log-Dateien
```
C:\ExchangeDoku\Exchange-Dokumentation_20260604_120000.log
```

### Exit-Codes
```powershell
0   = Erfolgreich
1   = Fehler bei Exchange-Tools
2   = Fehler bei Netzwerk/Servern
3   = Fehler bei Dateigeschrieb
```

---

## 💡 Tipps & Tricks

### Tipp 1: Mehrmals pro Woche dokumentieren
```powershell
# PowerShell-Scheduled-Task erstellen
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Wednesday,Friday -At 2AM
$action = New-ScheduledTaskAction -Execute powershell.exe -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\Scripts\DocumentExchange.ps1"
Register-ScheduledTask -TaskName "Exchange Weekly Documentation" -Trigger $trigger -Action $action -RunLevel Highest
```

### Tipp 2: Ausgaben mit Datum im Namen
```powershell
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputDir = "D:\Reports\Exchange_$timestamp"

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -OutputPath $outputDir
```

### Tipp 3: Nur kritische Sektionen für schnelle Reports
```powershell
$quickSections = @('Hardware','ExchangeServer','Certificates','DatabaseAndDAG')

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections $quickSections `
  -OutputPath 'C:\QuickReport'
```

---

## 🎯 Nächste Schritte

- ➡️ [Konfiguration](Konfiguration) – Erweiterte Anpassungen
- ➡️ [Beispiele](Beispiele) – Weitere Szenarien
- ➡️ [Best Practices](Best-Practices) – Optimierungen
