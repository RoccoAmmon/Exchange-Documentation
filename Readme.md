# 📊 Exchange Server Dokumentations-Tool

> **Automatische Erfassung und Dokumentation von Microsoft Exchange On-Premises Umgebungen (Exchange 2013, 2016, 2019 & Subscription Edition)**

[![Status](https://img.shields.io/badge/Status-Release-brightgreen?style=flat-square)](./CHANGELOG.md)
[![Version](https://img.shields.io/badge/Version-1.5-blue?style=flat-square)](./CHANGELOG.md)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?style=flat-square)](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](./LICENSE)
[![GitHub Issues](https://img.shields.io/github/issues/RoccoAmmon/Exchange-Documentation?style=flat-square)](https://github.com/RoccoAmmon/Exchange-Documentation/issues)
[![Downloads](https://img.shields.io/github/downloads/RoccoAmmon/Exchange-Documentation/total?style=flat-square&label=Downloads&color=orange)](https://github.com/RoccoAmmon/Exchange-Documentation/releases)
[![Latest Release Downloads](https://img.shields.io/github/downloads/RoccoAmmon/Exchange-Documentation/latest/total?style=flat-square&label=Latest)](https://github.com/RoccoAmmon/Exchange-Documentation/releases/latest)

---

## 🎯 Übersicht

Das **Exchange Dokumentations-Tool** ist ein umfassendes PowerShell-Skript, das automatisch eine detaillierte HTML-Dokumentation von Microsoft Exchange On-Premises Umgebungen erstellt. Es erfasst alle relevanten Konfigurationen, Systemeinstellungen und Sicherheitsaspekte – ideal für **Administratoren, Consultants, Auditoren und Compliance-Teams**.

Das Tool unterstützt **Exchange Server 2013, 2016, 2019** sowie **Exchange Server Subscription Edition (SE)** und bietet intelligente Fallback-Mechanismen für robuste Netzwerk-Kommunikation. Die Exchange-Edition wird zuverlässig über die Versionsnummer (15.0/15.1/15.2) erkannt.

---

## ✨ Features

### 🖥️ Hardware & Systeminfos
- CPU, RAM, Festplattenspeicher, Betriebssystem
- Virtualisierungs-Erkennung (VMware, Hyper-V, etc.)
- Windows Hotfixes und Update-Status
- Netzwerk-Konfiguration (IP-Adressen, DNS, DHCP)
- Pagefile-Einstellungen

### 📡 Exchange Konfiguration
- Exchange Server Übersicht (Edition, Build, Role)
- Alle virtuellen Verzeichnisse (OWA, ECP, EWS, MAPI, ActiveSync, OAB)
- Authentifizierungsmechanismen pro Verzeichnis
- Outlook Anywhere Konfiguration
- Autodiscover Service

### 📊 Infrastruktur & Hochverfügbarkeit
- Database Availability Group (DAG) Status
- Mailbox-Datenbanken mit Copy-Status
- Backup-Status und -Historie
- Datenbank-Quotas und Aufbewahrungsrichtlinien
- Lagged Copies und Replay Queue

### 🔒 Sicherheit & Compliance
- SSL/TLS Zertifikate mit Ablauf-Tracking
- Transport-Regeln (Mail Flow Rules)
- TLS Send/Receive Domain Secure Lists
- Anti-Spam & Anti-Malware Konfiguration
- Compliance & Data Loss Prevention (DLP)
- Litigation Hold und In-Place Hold Status
- RBAC-Rollengruppen und Berechtigungen

### �️ Modernes GUI-Redesign (v1.5)
- **Komplett überarbeitete WPF-Oberfläche** mit modernem, sauberem Design
- **Farbverlauf-Header** mit Firmenlogo-Bereich und Version-Badge
- **Moderne Karten (Cards)** mit dezenten Schatteneffekten für alle Bereiche
- **Kategorisierte Dokumentationsbereiche** mit Icons und分组 visueller Trennung
- **Farbige Status-Anzeige** mit Icons (Error/Success/Info)
- **Elegante Toggle-Buttons** für Ausgabeformat-Auswahl
- **Benutzerdefinierte Steuerelemente** (ScrollViewer, CheckBox, TextBox, Buttons)
- **Verbesserte Statusmeldungen** während der Dokumentation
- **Automatische Erstellung des Ausgabeverzeichnisses**

### �🛡️ Erweiterte Funktionen
- **Exchange Emergency Mitigation Service (EEMS)** - Status und Telemetrie
- **Transportkomponenten - Physische Speicherorte** - Queue-DB, Logs, Message-Tracking, SMTP-Protokolle, Safety-Net
- Active Directory Integration (FSMO-Rollen, Schema-Version)
- Remote Registry Access (ohne WinRM-Abhängigkeit)
- Automatischer CIM → DCOM Fallback
- Mail-Tips & MailTips-Konfiguration
- Journal Rules und Message Tracking

### 📄 Export & Output
- **HTML-Export** mit formatiertem Inhaltsverzeichnis
- **PDF-Export** (optional)
- **Markdown-Export** (optional)
- Word-kompatible HTML-Struktur
- Detailliertes Logging

---

## 🚀 Schnellstart

### Voraussetzungen

| Komponente | Anforderung |
|---|---|
| **PowerShell** | 5.1+ oder PowerShell 7.x |
| **Exchange Tools** | Exchange Management Shell oder Exchange SE Module |
| **AD Module** | Active Directory PowerShell-Modul |
| **Netzwerk** | CIM/RPC oder WinRM zu den Zielservern |
| **Rechte** | Administrative Privilegien auf Exchange-Servern |

### Installation

```powershell
# 1. Repository klonen
git clone https://github.com/RoccoAmmon/Exchange-Documentation.git
cd Exchange-Documentation

# 2. Ausführungsrichtlinie (falls nötig) anpassen
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

# 3. Skript ausführen
.\Exchange_Dokumentation.ps1
```

---

## � Downloads

### Direkte Download-Optionen

| Methode | Link | Info |
|---|---|---|
| **🔗 Raw-Download (Main)** | [Exchange_Documentation.ps1](https://raw.githubusercontent.com/RoccoAmmon/Exchange-Documentation/main/Exchange_Documentation.ps1) | Neueste Version direkt vom Repository |
| **📦 Git Clone** | `git clone https://github.com/RoccoAmmon/Exchange-Documentation.git` | Komplettes Repository mit Versionskontrolle |
| **📋 Releases** | [Alle Releases](https://github.com/RoccoAmmon/Exchange-Documentation/releases) | Veröffentlichte Versionen (wenn verfügbar) |
| **⬇️ ZIP Download** | [Download ZIP](https://github.com/RoccoAmmon/Exchange-Documentation/archive/refs/heads/main.zip) | Komplettes Repository als ZIP |

> **💡 Empfehlung:** Nutze `git clone` für regelmäßige Updates und Best Practices!

### Schneller Download der PS1-Datei

```powershell
# Direkt herunterladen und speichern
$ScriptUrl = "https://raw.githubusercontent.com/RoccoAmmon/Exchange-Documentation/main/Exchange_Documentation.ps1"
$OutFile = "C:\Scripts\Exchange_Documentation.ps1"
Invoke-WebRequest -Uri $ScriptUrl -OutFile $OutFile
Write-Host "✓ Heruntergeladen: $OutFile"
```

### 📊 Community-Nutzung

Die aktuellen Zugriffs- und Download-Statistiken finden Sie:
- **Badges oben:** Live-Statistiken für den gesamten Download
- **[GitHub Insights](https://github.com/RoccoAmmon/Exchange-Documentation/graphs/traffic):** Detaillierte Traffic- und Download-Analysen

---

## �📖 Verwendungsbeispiele

### Einfaches Beispiel (GUI-Auswahl)
```powershell
# Startet das Skript mit grafischer Oberfläche zur Server-Auswahl
.\Exchange_Dokumentation.ps1
```

### Mehrere Server dokumentieren
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03') `
  -CompanyName 'Contoso GmbH' `
  -OutputPath 'D:\ExchangeInventory'
```

### Ohne GUI, mit allen Formaten
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Meine Firma' `
  -OutputPath 'C:\Reports' `
  -OutputFormats @('HTML','PDF','Markdown') `
  -NoGui
```

### Nur spezifische Dokumentationsbereiche
```powershell
$sections = @(
    'Hardware',
    'ExchangeServer',
    'Databases',
    'Certificates'
)

.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -Sections $sections `
  -OutputPath 'C:\Reports'
```

---

## ⚙️ Parameter-Referenz

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `ExchangeServers` | string[] | (GUI) | Zu dokumentierende Exchange-Server (Array) |
| `CompanyName` | string | "Meine Organisation" | Firmenname im Bericht |
| `OutputPath` | string | C:\ExchangeDoku | Ausgabeverzeichnis für Dokumente |
| `OutputFormats` | string[] | @('HTML') | Ausgabeformate: HTML, PDF, Markdown |
| `Sections` | string[] | (alle) | Zu dokumentierende Abschnitte |
| `ShowGui` | switch | - | GUI erzwingen |
| `NoGui` | switch | - | GUI unterdrücken |

---

## 📋 Dokumentierte Inhalte

Das generierte Dokument enthält automatisch:

✅ Hardwaredetails (CPU, RAM, Festplatte, Netzwerk)  
✅ Windows-Betriebssystem Info  
✅ Exchange Server Edition & Build  
✅ Dienste-Status  
✅ Zertifikate (mit Ablauf-Tracking)  
✅ Virtuelle Verzeichnisse & URLs  
✅ Authentifizierungsmethoden  
✅ DAG-Konfiguration  
✅ Datenbanken & Backups  
✅ Transport-Regeln  
✅ **Transportkomponenten - Speicherorte** (Queue, Logs, Message-Tracking, SMTP-Protokoll, Safety-Net)  
✅ Accepted Domains & Remote Domains  
✅ FSMO-Rollen & AD-Info  
✅ RBAC-Konfiguration  
✅ Event Logs (7 Tage)  
✅ Hybrid-Konfiguration  
✅ DLP & Compliance  
✅ EEMS-Status  

---

## 🔧 Konfigurationsoptionen

Im Skript sind folgende Variablen konfigurierbar (Zeile ~100-120):

```powershell
$script:WarningDiskSpaceGB      = 20      # Warnung bei weniger Speicher
$script:WarningCertDaysExpiry   = 30      # Zertifikat-Ablauf-Warnung (Tage)
$script:MaxMailboxesForStats    = 500     # Limit für Mailbox-Statistiken
$script:DNSServer               = ""      # DNS-Server für MX-Abfragen (leer = Standard)
```

---

## 🌐 Netzwerkverbindung

Das Tool verwendet intelligente Verbindungsmechanismen:

1. **WsMan (WinRM)** - Bevorzugter Standard
2. **DCOM RPC Fallback** - Funktioniert auch ohne WinRM
3. **Remote Registry** - .NET-basiert, kein WinRM nötig
4. **Invoke-Command** - Für lokal verfügbare Befehle

✅ **Vorteil:** Funktioniert auch in Umgebungen mit deaktiviertem WinRM!

---

## 📝 Output-Beispiel

Das generierte HTML-Dokument hat diese Struktur:

```
Exchange Server Dokumentation
├─ Inhaltsverzeichnis
├─ Zusammenfassung
│  ├─ Dokumentierte Server
│  ├─ Erstellungsdatum
│  ├─ Fehler- und Warnung-Zähler
│
├─ Hardware-Informationen & Server-Details
│  ├─ System-Übersicht
│  ├─ Betriebssystem
│  ├─ Prozessor
│  ├─ Speicher & Festplatten
│  └─ Netzwerk
│
├─ Exchange Konfiguration
│  ├─ Server-Übersicht
│  ├─ Virtuelle Verzeichnisse
│  ├─ Zertifikate
│  ├─ Datenbanken & DAG
│  └─ Transport-Einstellungen
│
└─ [weitere Sektionen...]
```

---

## 🐛 Troubleshooting

### Problem: "Exchange Management Tools nicht gefunden"
```powershell
# Sicherstellen, dass auf einem Exchange-Server ausgeführt wird
# ODER Exchange Management Shell manuell laden:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

### Problem: "Keine Verbindung zu Server XY"
```powershell
# Netzwerk-Erreichbarkeit prüfen:
Test-Connection -ComputerName 'EX01' -Count 2

# DCOM/RPC aktivieren:
Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)"
```

### Problem: "Zugriff verweigert"
```powershell
# Mit Admin-Rechten ausführen
# oder UAC-Bypass:
Start-Process powershell -ArgumentList "& '.\Exchange_Dokumentation.ps1'" -Verb RunAs
```

### Problem: "Keine Sektionen werden dokumentiert"
```powershell
# Alle verfügbaren Sektionen auflisten und ohne Filter ausführen
.\Exchange_Dokumentation.ps1 -ExchangeServers @('EX01') -NoGui
```

---

## 📚 Wiki & Dokumentation

Ausführliche Dokumentation im **GitHub Wiki**:
- 📖 **[Installation](https://github.com/RoccoAmmon/Exchange-Documentation/wiki/Installation)** - Schritt-für-Schritt Anleitung
- 🎮 **[Bedienung](https://github.com/RoccoAmmon/Exchange-Documentation/wiki/Bedienung)** - Parameter & Beispiele
- ⚙️ **[Konfiguration](https://github.com/RoccoAmmon/Exchange-Documentation/wiki/Konfiguration)** - Erweiterte Einstellungen

---

## 🤝 Mitarbeit / Contributing

Beiträge sind willkommen! Bitte:

1. **Fork** das Repository
2. **Feature Branch** erstellen (`git checkout -b feature/MyFeature`)
3. **Änderungen committen** (`git commit -am 'Add MyFeature'`)
4. **Branch pushen** (`git push origin feature/MyFeature`)
5. **Pull Request** öffnen

---

## ❓ FAQ

**F: Funktioniert das Skript auch auf non-Administrator-Konten?**  
A: Nein, administrative Rechte sind erforderlich.

**F: Kann ich das Dokument in Word bearbeiten?**  
A: Ja, die HTML-Datei kann direkt in Word geöffnet werden.

**F: Wie lange dauert die Dokumentation?**  
A: Je nach Anzahl Server und Umgebungsgröße 5-30 Minuten.

**F: Welche Daten werden übertragen?**  
A: Nur Lesevorgänge auf die Exchange-Server. Keine Daten verlassen die Umgebung.

**F: Wird Exchange Online (Hybrid) unterstützt?**  
A: Ja, Hybrid-Konfigurationen werden dokumentiert.

---

## 📄 Lizenz

Dieses Projekt ist unter der **MIT-Lizenz** lizenziert – siehe [LICENSE](LICENSE) für Details.

---

## 👤 Autor

**Rocco Ammon**

- 🔗 GitHub: [@RoccoAmmon](https://github.com/RoccoAmmon)
- 📧 Kontakt via GitHub Issues
- 🌍 Repository: [Exchange-Documentation](https://github.com/RoccoAmmon/Exchange-Documentation)

---

## 📞 Support

- 📋 **Issues**: [GitHub Issues](https://github.com/RoccoAmmon/Exchange-Documentation/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/RoccoAmmon/Exchange-Documentation/discussions)
- 📖 **Wiki**: [GitHub Wiki](https://github.com/RoccoAmmon/Exchange-Documentation/wiki)

---

**Hinweis:** Vor dem Einsatz in **produktiven Umgebungen** die Skripte in einer **Testumgebung** prüfen.
