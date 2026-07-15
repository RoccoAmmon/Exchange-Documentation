# 📊 Exchange Server Dokumentations-Tool / Exchange Documentation Tool

> **DE:** Automatische Erfassung und Dokumentation von Microsoft Exchange On-Premises Umgebungen (Exchange 2013, 2016, 2019 & Subscription Edition)
> **EN:** Automated inventory and documentation of Microsoft Exchange On-Premises environments (Exchange 2013, 2016, 2019 & Subscription Edition)

[![Status](https://img.shields.io/badge/Status-Release-brightgreen?style=flat-square)](./CHANGELOG.md)
[![Version](https://img.shields.io/badge/Version-1.9-blue?style=flat-square)](./CHANGELOG.md)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?style=flat-square)](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](./LICENSE)
[![GitHub Issues](https://img.shields.io/github/issues/RoccoAmmon/Exchange-Documentation?style=flat-square)](https://github.com/RoccoAmmon/Exchange-Documentation/issues)
[![Downloads](https://img.shields.io/github/downloads/RoccoAmmon/Exchange-Documentation/total?style=flat-square&label=Downloads&color=orange)](https://github.com/RoccoAmmon/Exchange-Documentation/releases)
[![Latest Release Downloads](https://img.shields.io/github/downloads/RoccoAmmon/Exchange-Documentation/latest/total?style=flat-square&label=Latest)](https://github.com/RoccoAmmon/Exchange-Documentation/releases/latest)

---

## 🎯 Übersicht / Overview

**DE**
Das **Exchange Dokumentations-Tool** ist ein umfassendes PowerShell-Skript, das automatisch eine detaillierte HTML-Dokumentation von Microsoft Exchange On-Premises Umgebungen erstellt. Es erfasst alle relevanten Konfigurationen, Systemeinstellungen und Sicherheitsaspekte – ideal für **Administratoren, Consultants, Auditoren und Compliance-Teams**.

Das Tool unterstützt **Exchange Server 2013, 2016, 2019** sowie **Exchange Server Subscription Edition (SE)** und bietet intelligente Fallback-Mechanismen für robuste Netzwerk-Kommunikation. Die Exchange-Edition wird zuverlässig über die Versionsnummer (15.0/15.1/15.2) erkannt.

**EN**
The **Exchange Documentation Tool** is a comprehensive PowerShell script that automatically generates detailed HTML documentation of Microsoft Exchange On-Premises environments. It captures all relevant configurations, system settings, and security aspects – ideal for **administrators, consultants, auditors, and compliance teams**.

The tool supports **Exchange Server 2013, 2016, 2019** and **Exchange Server Subscription Edition (SE)** with intelligent fallback mechanisms for robust network communication. The Exchange edition is reliably detected via the version number (15.0/15.1/15.2).

---

## ✨ Features / Funktionen

### 🖥️ Hardware & Systeminfos / System Information
- CPU, RAM, disk space, operating system
- Virtualization detection (VMware, Hyper-V, etc.)
- Windows Hotfixes and update status
- Network configuration (IP addresses, DNS, DHCP)
- Pagefile settings

### 📡 Exchange Konfiguration / Exchange Configuration
- Exchange Server overview (edition, build, role)
- All virtual directories (OWA, ECP, EWS, MAPI, ActiveSync, OAB)
- Authentication mechanisms per directory
- Outlook Anywhere configuration
- Autodiscover service

### 📊 Infrastruktur & Hochverfügbarkeit / Infrastructure & High Availability
- Database Availability Group (DAG) status
- Mailbox databases with copy status
- Backup status and history
- Database quotas and retention policies
- Lagged copies and replay queue

### 🔒 Sicherheit & Compliance / Security & Compliance
- SSL/TLS certificates with expiry tracking
- Transport rules (Mail Flow Rules)
- TLS Send/Receive Domain Secure Lists
- Anti-Spam & Anti-Malware configuration
- Compliance & Data Loss Prevention (DLP)
- Litigation Hold and In-Place Hold status
- RBAC role groups and permissions

### 🖥️ Modernes GUI-Redesign v1.5 / Modern GUI Redesign
- **Complete WPF redesign** with modern, clean design
- **Gradient header** with company logo area and version badge
- **Modern cards** with subtle shadow effects
- **Categorized documentation areas** with icons and visual separation
- **Color-coded status display** with icons (Error/Success/Info)
- **Elegant toggle buttons** for output format selection
- **Custom controls** (ScrollViewer, CheckBox, TextBox, Buttons)
- **Improved status messages** during documentation
- **Automatic output directory creation**

### 🛡️ Erweiterte Funktionen / Advanced Features
- **Exchange Emergency Mitigation Service (EEMS)** – status and telemetry
- **Transport Components - Physical Locations** – Queue DB, Logs, Message Tracking, SMTP protocols, Safety-Net
- Active Directory integration (FSMO roles, schema version)
- Remote Registry access (no WinRM dependency)
- Automatic CIM → DCOM fallback
- MailTips configuration
- Journal Rules and Message Tracking

### 📦 Neue Funktionen & Fixes / New Features & Fixes

**v1.9**
- **Exchange Build-Versionskatalog** – Umfassende Lookup-Tabelle / Build version catalog
- **Online-Versionsprüfung** – Automatischer Abgleich mit Microsoft Learn / Online version check
- **Produktname-Anzeige** – Freundlicher Name (z.B. "Exchange Server SE RTM Jul26SU") / Product name display
- **Update-Warnung** – ⚠️ bei veralteter Version + Link zum aktuellsten Build / Update warning
- **Support-End-Erkennung** – 🔴 Kritisch bei beendetem Support / Support end detection
- **Neue Spalten** in Server-Übersicht und Build-Informationen / New columns

**v1.8**
- **AutoReseed-Konfiguration** – AutoDagVolumesRootFolderPath, AutoDagDatabasesRootFolderPath, Spare Volumes
- **Security CVE Prüfung** – Live-BSI RSS-Feed für Exchange & Windows / Live BSI RSS feed
- **FIP-FS Scan Engine** – 4-stufige Fallback-Suche / 4-stage fallback search
- **ConvertTo-HTMLTable** – Link-Spalte korrigiert / Link column fixed
- **Anti-Malware Status** – Logik korrigiert / Logic corrected

**v1.7 – 15 neue System- & Sicherheitschecks / 15 new system & security checks**
- Windows Features & Rollen / Windows Features & Roles
- .NET Framework Version & DLLs
- Ausstehende Neustarts (6 Methoden) / Pending Reboots (6 methods)
- CPU Throttling Analyse / CPU Throttling analysis
- Visual C++ Redistributable (32/64-Bit)
- Credential Guard Status
- Lokale Administratoren / Local Administrators
- Domain Trusts & Verschlüsselung / Domain Trusts & encryption
- FIP-FS Scan Engine Version
- Exchange Setting Overrides
- Exchange Server Component State
- Security CVE Prüfung / CVE check
- HTTP Proxy Konfiguration / HTTP Proxy configuration
- Installierte Antivirenlösung / Installed antivirus
- NIC Receive Buffer Analyse / NIC Receive Buffer analysis

**v1.6 – 7 neue Dokumentationsbereiche / 7 new documentation areas**
- Message Queue Analyse / Analysis
- Calendar & Resource Mailbox
- Exchange Archive
- Exchange Message Size Limits
- Partner Applications (SharePoint, Skype, CRM)
- Federated Sharing
- OAuth / Certificate Based Auth

### 📄 Export & Output / Ausgabe
- **HTML Export** with formatted table of contents
- **PDF Export** (optional)
- **Markdown Export** (optional)
- Word-compatible HTML structure
- Detailed logging

---

## 🚀 Schnellstart / Quick Start

### Voraussetzungen / Requirements

| Komponente / Component | Anforderung / Requirement |
|---|---|
| **PowerShell** | 5.1+ or PowerShell 7.x |
| **Exchange Tools** | Exchange Management Shell or Exchange SE Module |
| **AD Module** | Active Directory PowerShell module |
| **Netzwerk / Network** | CIM/RPC or WinRM to target servers |
| **Rechte / Permissions** | Administrative privileges on Exchange servers |

### Installation

```powershell
# 1. Repository klonen / clone repository
git clone https://github.com/RoccoAmmon/Exchange-Documentation.git
cd Exchange-Documentation

# 2. Execution policy anpassen (falls nötig) / adjust if needed
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

# 3. Skript ausführen / run script
.\Exchange_Dokumentation.ps1
```

---

## 📥 Downloads

### Direkte Download-Optionen / Direct Download Options

| Methode / Method | Link | Info |
|---|---|---|
| **🔗 Raw-Download (Main)** | [Exchange_Documentation.ps1](https://raw.githubusercontent.com/RoccoAmmon/Exchange-Documentation/main/Exchange_Documentation.ps1) | Latest version directly from repository |
| **📦 Git Clone** | `git clone https://github.com/RoccoAmmon/Exchange-Documentation.git` | Full repository with version control |
| **📋 Releases** | [All Releases](https://github.com/RoccoAmmon/Exchange-Documentation/releases) | Published versions |
| **⬇️ ZIP Download** | [Download ZIP](https://github.com/RoccoAmmon/Exchange-Documentation/archive/refs/heads/main.zip) | Full repository as ZIP |

> **💡 Recommendation:** Use `git clone` for regular updates and best practices!

### Schneller Download der PS1-Datei / Quick PS1 Download

```powershell
# Download and save directly
$ScriptUrl = "https://raw.githubusercontent.com/RoccoAmmon/Exchange-Documentation/main/Exchange_Documentation.ps1"
$OutFile = "C:\Scripts\Exchange_Documentation.ps1"
Invoke-WebRequest -Uri $ScriptUrl -OutFile $OutFile
Write-Host "✓ Downloaded: $OutFile"
```

### 📊 Community-Nutzung / Community Usage

**DE:** Die aktuellen Zugriffs- und Download-Statistiken finden Sie:
**EN:** Current traffic and download statistics:
- **Badges above:** Live statistics for total downloads
- **[GitHub Insights](https://github.com/RoccoAmmon/Exchange-Documentation/graphs/traffic):** Detailed traffic and download analytics

---

## 📖 Verwendungsbeispiele / Usage Examples

### Einfaches Beispiel (GUI-Auswahl) / Basic Example (GUI Selection)
```powershell
# Startet das Skript mit grafischer Oberfläche zur Server-Auswahl
# Starts the script with GUI for server selection
.\Exchange_Dokumentation.ps1
```

### Mehrere Server dokumentieren / Document Multiple Servers
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02','EX03') `
  -CompanyName 'Contoso GmbH' `
  -OutputPath 'D:\ExchangeInventory'
```

### Ohne GUI, mit allen Formaten / Without GUI, All Formats
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'My Company' `
  -OutputPath 'C:\Reports' `
  -OutputFormats @('HTML','PDF','Markdown') `
  -NoGui
```

### Nur spezifische Dokumentationsbereiche / Specific Documentation Sections
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

## ⚙️ Parameter-Referenz / Parameter Reference

| Parameter | Typ / Type | Default | Beschreibung / Description |
|---|---|---|---|
| `ExchangeServers` | string[] | (GUI) | Exchange servers to document (array) |
| `CompanyName` | string | "Meine Organisation" | Company name in report |
| `OutputPath` | string | C:\ExchangeDoku | Output directory for documents |
| `OutputFormats` | string[] | @('HTML') | Output formats: HTML, PDF, Markdown |
| `Sections` | string[] | (all) | Documentation sections to include |
| `ShowGui` | switch | - | Force GUI |
| `NoGui` | switch | - | Suppress GUI |

---

## 📋 Dokumentierte Inhalte / Documented Contents

**DE:** Das generierte Dokument enthält automatisch:
**EN:** The generated document automatically includes:

✅ Hardware details (CPU, RAM, disk, network)  
✅ Windows OS information  
✅ Exchange Server edition & build  
✅ Service status  
✅ Certificates (with expiry tracking)  
✅ Virtual directories & URLs  
✅ Authentication methods  
✅ DAG configuration  
✅ Databases & backups  
✅ Transport rules  
✅ **Transport component locations** (Queue, Logs, Message Tracking, SMTP, Safety-Net)  
✅ Accepted Domains & Remote Domains  
✅ FSMO roles & AD info  
✅ RBAC configuration  
✅ Event logs (7 days)  
✅ Hybrid configuration  
✅ DLP & Compliance  
✅ EEMS status  
✅ **Windows Features & Roles** (v1.7)  
✅ **.NET Framework Version & DLLs** (v1.7)  
✅ **Pending reboots** (v1.7)  
✅ **CPU Throttling analysis** (v1.7)  
✅ **Visual C++ Redistributable** (v1.7)  
✅ **Credential Guard status** (v1.7)  
✅ **Local administrators** (v1.7)  
✅ **Domain Trusts & encryption** (v1.7)  
✅ **FIP-FS Scan Engine version** (v1.7)  
✅ **Exchange Setting Overrides** (v1.7)  
✅ **Server Component State** (v1.7)  
✅ **Security CVE check** (v1.7)  
✅ **HTTP Proxy configuration** (v1.7)  
✅ **Installed antivirus** (v1.7)  
✅ **Automatic Database Reseed** (v1.8)  
✅ **Security CVE via BSI RSS-Feed** (v1.8)  
✅ **FIP-FS Scan Engine fallback** (v1.8)  

---

## 🔧 Konfigurationsoptionen / Configuration Options

**DE:** Im Skript sind folgende Variablen konfigurierbar (Zeile ~100-120):
**EN:** The following variables can be configured in the script (line ~100-120):

```powershell
$script:WarningDiskSpaceGB      = 20      # Warning when disk space is low
$script:WarningCertDaysExpiry   = 30      # Certificate expiry warning (days)
$script:MaxMailboxesForStats    = 500     # Limit for mailbox statistics
$script:DNSServer               = ""      # DNS server for MX queries (empty = default)
```

---

## 🌐 Netzwerkverbindung / Network Connection

**DE:** Das Tool verwendet intelligente Verbindungsmechanismen:
**EN:** The tool uses intelligent connection mechanisms:

1. **WsMan (WinRM)** – Preferred standard
2. **DCOM RPC Fallback** – Works without WinRM
3. **Remote Registry** – .NET-based, no WinRM required
4. **Invoke-Command** – For locally available commands

✅ **Vorteil / Advantage:** Works even in environments with WinRM disabled!

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
