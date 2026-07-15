# Changelog / Versionshistorie

**DE:** Alle signifikanten Änderungen am Projekt werden hier dokumentiert.
**EN:** All notable changes to this project are documented here.

---

## [1.9] - 2026-07-15

**DE**
- ✨ **Exchange Build-Versionskatalog & Online-Abgleich**
  - Umfassende statische Lookup-Tabelle (SE, 2019 CU15-CU12, 2016 CU23, 2013 CU23)
  - Automatischer Online-Abgleich mit Microsoft Learn Build-Tabelle
  - Freundlicher Produktname (z.B. "Exchange Server SE RTM Jul26SU")
  - Update-Status: ⚠️ Warnung bei veralteter Version + Anzeige der aktuellsten Version
  - Link zum aktuellsten Build direkt in der HTML-Tabelle
- 🔴 **Support-End-Erkennung** – Markiert beendeten Support (2013/2016/2019) als kritisch
- 🔄 **Online-Versionscheck** – Einmaliger Abruf von MS Learn beim Skriptstart
- 📊 **Neue Spalten in Exchange Server Übersicht:** ProductName, LatestBuild, Status
- 📊 **Neue Spalten in Exchange Version & Build:** ProductName, LatestBuild, Status
- ♻️ **Fallback auf statischen Katalog** bei fehlender Internetverbindung

**EN**
- ✨ **Exchange Build Version Catalog & Online Sync**
  - Comprehensive static lookup table (SE, 2019 CU15-CU12, 2016 CU23, 2013 CU23)
  - Automatic online sync with Microsoft Learn build table
  - Friendly product name display (e.g. "Exchange Server SE RTM Jul26SU")
  - ⚠️ Update warning for outdated versions with link to latest build
  - Link to latest build directly in HTML table
- 🔴 **Support End Detection** – Marks ended support (2013/2016/2019) as critical
- 🔄 **Online Version Check** – Fetches MS Learn build table once at script startup
- 📊 **New columns in Exchange Server Overview:** ProductName, LatestBuild, Status
- 📊 **New columns in Exchange Version & Build:** ProductName, LatestBuild, Status
- ♻️ **Static catalog fallback** when no internet connection available

## [1.8] - 2026-07-10

**DE**
- ✨ **Neue Funktion: AutoReseed-Konfiguration**
  - Automatic Database Reseed (AutoReseed) Prüfung für DAG-Mitglieder
  - AutoDagVolumesRootFolderPath und AutoDagDatabasesRootFolderPath
  - Spare Volumes / Recovery-Datenbanken erkennen
- 🔒 **Security CVE Prüfung auf BSI RSS-Feed umgestellt**
  - Live-Abfrage von https://wid.cert-bund.de/content/public/securityAdvisory/rss
  - Filter erweitert auf Exchange + Windows Server (Microsoft-Produkte)
  - Korrektes XML-Parsing via Invoke-WebRequest + [xml] (RSS 2.0)
- 🛠️ **FIP-FS Scan Engine – 4-stufige Fallback-Suche**
  - Stufe 1: Rekursive Suche nach ScanEngine.ini
  - Stufe 2: .ini-Dateien mit EngineVersion/EngineName
  - Stufe 3: Breitsuche im ProgramFiles
  - Stufe 4: Registry-Fallback
- 🐛 **ConvertTo-HTMLTable: Link-Spalte korrigiert**
  - Link-Spalte wird nicht mehr HTML-encoded (Links sind jetzt klickbar)
- 🐛 **Anti-Malware Status Logik korrigiert**
  - Zeigt nur aktiv, wenn FIP-FS Engine tatsächlich vorhanden
- ♻️ **CVE-Filter auf Microsoft-Produkte erweitert**
  - Findet jetzt alle relevanten Microsoft-Sicherheitsmeldungen

**EN**
- ✨ **New Feature: AutoReseed Configuration**
  - Automatic Database Reseed check for DAG members
  - AutoDagVolumesRootFolderPath and AutoDagDatabasesRootFolderPath
  - Spare volume / Recovery database detection
- 🔒 **Security CVE Check switched to BSI RSS feed**
  - Live query from wid.cert-bund.de
  - Filter expanded to Exchange + Windows Server (Microsoft products)
  - Correct XML parsing via Invoke-WebRequest + [xml] (RSS 2.0)
- 🛠️ **FIP-FS Scan Engine – 4-stage fallback search**
  - Stage 1: Recursive search for ScanEngine.ini
  - Stage 2: .ini files with EngineVersion/EngineName
  - Stage 3: Broad search in ProgramFiles
  - Stage 4: Registry fallback
- 🐛 **ConvertTo-HTMLTable: Link column fixed**
  - Link column no longer HTML-encoded (links are now clickable)
- 🐛 **Anti-Malware status logic corrected**
  - Only shows active when FIP-FS engine is actually present
- ♻️ **CVE filter expanded to Microsoft products**
  - Now finds all relevant Microsoft security advisories

## [1.7] - 2026-07-10

**DE**
- ✨ **15 neue System- & Sicherheitschecks hinzugefügt**
  - Windows Features & Rollen (installierte Server Rollen per CIM/DISM)
  - .NET Framework Version & DLLs (Release-Key, Assembly-Versionen)
  - Ausstehende Neustarts (6 Prüfmethoden: PendingFileRename, WU, CBS, ServerMgr, CIM, SCCM)
  - CPU Throttling Analyse (CurrentClockSpeed vs MaxClockSpeed)
  - Visual C++ Redistributable Versionen (32/64-Bit via Registry)
  - Credential Guard Status (LsaCfgFlags, Virtualization Based Security)
  - Lokale Administratoren (Mitglieder via CIM/ADSI)
  - Domain Trusts & Verschlüsselung (Trust-Typen, SupportedEncryptionTypes)
  - FIP-FS Scan Engine Version (Anti-Malware Engine, Pattern-Update)
  - Exchange Setting Overrides (alle aktiven Overrides dokumentieren)
  - Exchange Server Component State (Maintenance Mode Erkennung)
  - Security CVE Prüfung (CVE-2021-34470, CVE-2022-21978)
  - HTTP Proxy Konfiguration (WinHTTP, Registry, netsh)
  - Installierte Antivirenlösung (SecurityCenter WMI, Registry, Defender-Status)
  - NIC Receive Buffer Analyse (10/25/40 Gbit/s Check, StalledDueTo-Prävention)

**EN**
- ✨ **15 new system & security checks added**
  - Windows Features & Roles (installed server roles via CIM/DISM)
  - .NET Framework Version & DLLs (Release-Key, Assembly versions)
  - Pending reboots (6 check methods: PendingFileRename, WU, CBS, ServerMgr, CIM, SCCM)
  - CPU Throttling analysis (CurrentClockSpeed vs MaxClockSpeed)
  - Visual C++ Redistributable versions (32/64-bit via Registry)
  - Credential Guard status (LsaCfgFlags, Virtualization Based Security)
  - Local administrators (members via CIM/ADSI)
  - Domain Trusts & encryption (Trust types, SupportedEncryptionTypes)
  - FIP-FS Scan Engine version (Anti-Malware Engine, Pattern-Update)
  - Exchange Setting Overrides (document all active overrides)
  - Exchange Server Component State (Maintenance Mode detection)
  - Security CVE check (CVE-2021-34470, CVE-2022-21978)
  - HTTP Proxy configuration (WinHTTP, Registry, netsh)
  - Installed antivirus solution (SecurityCenter WMI, Registry, Defender status)
  - NIC Receive Buffer analysis (10/25/40 Gbit/s check, StalledDueTo prevention)

## [1.6] - 2026-07-10

**DE**
- ✨ **7 neue Dokumentationsbereiche hinzugefügt**
  - Message Queue Analyse (Warteschlangen, Status, Nachrichtenanzahl)
  - Calendar & Resource Mailbox (Raum-/Ressourcenpostfächer, Buchungsoptionen)
  - Exchange Archive (Archivpostfächer, Quotas, Auto-Expanding Archive)
  - Exchange Message Size Limits (Org, Connector, Remote Domain, Benutzer)
  - Exchange Partner Applications (SharePoint, Skype, CRM)
  - Exchange Federated Sharing (Federation Trust, Organization Relationships)
  - OAuth / Certificate Based Auth (Auth Server, Zertifikate, CBA)

**EN**
- ✨ **7 new documentation areas added**
  - Message Queue analysis (queues, status, message count)
  - Calendar & Resource Mailbox (room/resource mailboxes, booking options)
  - Exchange Archive (archive mailboxes, quotas, Auto-Expanding Archive)
  - Exchange Message Size Limits (Org, Connector, Remote Domain, user)
  - Exchange Partner Applications (SharePoint, Skype, CRM)
  - Exchange Federated Sharing (Federation Trust, Organization Relationships)
  - OAuth / Certificate Based Auth (Auth Server, certificates, CBA)

## [1.5] - 2026-06-30

**DE**
- 🎨 **Komplettes GUI-Redesign**
  - Moderne WPF-Oberfläche mit Farbverlauf-Header und Version-Badge
  - Karten (Cards) mit dezenten Schatteneffekten für alle Bereiche
  - Kategorisierte Dokumentationsbereiche mit Icons und visueller Trennung
  - Farbige Status-Anzeige mit Icons (Error/Success/Info)
  - Elegante Toggle-Buttons für Ausgabeformat-Auswahl
  - Benutzerdefinierte Steuerelemente (ScrollViewer, CheckBox, TextBox, Buttons)
- 🐛 **Bugfixes**
  - `[System.Windows.GridLength]::Star` → korrekte `::new(1, Star)` Syntax
  - `[System.Windows.Media.Brushes]::FromArgb` → `SolidColorBrush::new(Color::FromArgb)` (8 Stellen)
  - Doppelte `Thumb`-Zuweisung im ScrollBar-Template entfernt
  - `LetterSpacing` aus XAML entfernt (nicht in WPF unterstützt)
  - Automatische Erstellung des Ausgabeverzeichnisses hinzugefügt
  - Fallback-Verzeichniserstellung in `Write-Log` Funktion
- ✅ Verbesserte Fehlerbehandlung und Statusmeldungen in der GUI

**EN**
- 🎨 **Complete GUI Redesign**
  - Modern WPF interface with gradient header and version badge
  - Cards with subtle shadow effects for all areas
  - Categorized documentation areas with icons and visual separation
  - Color-coded status display with icons (Error/Success/Info)
  - Elegant toggle buttons for output format selection
  - Custom controls (ScrollViewer, CheckBox, TextBox, Buttons)
- 🐛 **Bugfixes**
  - `[System.Windows.GridLength]::Star` → correct `::new(1, Star)` syntax
  - `[System.Windows.Media.Brushes]::FromArgb` → `SolidColorBrush::new(Color::FromArgb)` (8 digits)
  - Removed duplicate `Thumb` assignment in ScrollBar template
  - Removed `LetterSpacing` from XAML (not supported in WPF)
  - Added automatic output directory creation
  - Fallback directory creation in `Write-Log` function
- ✅ Improved error handling and status messages in GUI

## [1.4] - 2026-06-29

**DE**
- ✨ **17 HealthChecker Checks integriert**
  - Processor Core Analyse (4-Kern Minimum-Check)
  - RAM Requirements Check (Exchange-spezifisch: 64GB/128GB)
  - Certificate Expiration Status (Ablauf-Ampel)
  - Exchange Service Status (Kritische Services Überwachung)
  - IIS Application Pool Konfiguration (Recycling-Zeiten aus ApplicationHost.config)
  - NIC Speed & Performance Checks (1 Gbps Best-Practice-Warnung)
  - Power Plan Konfiguration (Höchste Leistung Best-Practice)
  - SMBv1 Status & Security Check (Sicherheits-Best-Practice)
  - DAG Replication Health (mit Prüfung ob DAG konfiguriert)
- 🔧 **IIS AppPool XML-Auslesen korrigiert**
  - Behobenes Problem: Recycling-Zeiten werden jetzt korrekt aus ApplicationHost.config gelesen
  - Struktur war `applicationPools` statt `applicationPool` (mit s)
  - Recycling Element ist `recycling.periodicRestart.time` statt `recycleConfig`
  - Bessere Bewertungs-Logik (konfiguriert vs. Standard vs. Warnung)
- 🎨 **PDF-Export Optimierung**
  - Entfernte aggressive `page-break-before: always` Rules
  - Intelligente Seitenumbrüche: `page-break-inside: avoid` auf Sektionen/Tabellen
  - Automatische Wiederholung von Tabellen-Headern auf jeder Seite
  - Kompaktere Font-Größen und Margins für Print
- ✅ Automatischer Neustart mit Administrator-Rechten
- ✅ Auflistung installierter Software pro Server (32-Bit & 64-Bit)
- ✅ Integrierte Local/Remote-Erkennung (`$isLocal` Pattern)

**EN**
- ✨ **17 HealthChecker checks integrated**
  - Processor Core analysis (4-core minimum check)
  - RAM Requirements check (Exchange-specific: 64GB/128GB)
  - Certificate Expiration status (traffic light indicator)
  - Exchange Service status (critical services monitoring)
  - IIS Application Pool configuration (recycling times from ApplicationHost.config)
  - NIC Speed & Performance checks (1 Gbps best practice warning)
  - Power Plan configuration (High Performance best practice)
  - SMBv1 Status & Security check (security best practice)
  - DAG Replication Health (with check if DAG is configured)
- 🔧 **IIS AppPool XML reading fixed**
  - Fixed: Recycling times now correctly read from ApplicationHost.config
  - Structure was `applicationPools` instead of `applicationPool` (with s)
  - Recycling element is `recycling.periodicRestart.time` instead of `recycleConfig`
  - Improved evaluation logic (configured vs. default vs. warning)
- 🎨 **PDF Export optimization**
  - Removed aggressive `page-break-before: always` rules
  - Smart page breaks: `page-break-inside: avoid` on sections/tables
  - Automatic table header repeat on each page
  - Compact font sizes and margins for print
- ✅ Automatic restart with administrator privileges
- ✅ Installed software listing per server (32-bit & 64-bit)
- ✅ Integrated local/remote detection (`$isLocal` pattern)

## [1.3] - 2026-06-10

**DE**
- 🐛 **Verbesserte Exchange-Edition-Erkennung**
  - Erkennung jetzt primär über Versionsnummer (15.0/15.1/15.2) statt nur über Build-Ranges
  - Zusätzliche Unterstützung für Exchange 2013 (Version 15.0) und Exchange 2016 (Version 15.1)
  - Sauberere Fallback-Logik (kein hartes "2019"-Fallback mehr)
  - Konsistente Logik zwischen `Get-ExchangeEdition` und `Get-ExchangeServerOverview`
- 📚 Versionsinformationen in Header und NOTES aktualisiert

_Hinweis: Die Version 1.2 (TLS/SSL Konfiguration) wurde übersprungen und nicht als offizielles Release veröffentlicht._

**EN**
- 🐛 **Improved Exchange edition detection**
  - Detection now primarily via version number (15.0/15.1/15.2) instead of build ranges only
  - Added support for Exchange 2013 (Version 15.0) and Exchange 2016 (Version 15.1)
  - Cleaner fallback logic (no hard "2019" fallback anymore)
  - Consistent logic between `Get-ExchangeEdition` and `Get-ExchangeServerOverview`
- 📚 Updated version information in header and NOTES

_Note: Version 1.2 (TLS/SSL configuration) was skipped and not released officially._

## [1.1] - 2026-06-04

**DE**
- ✨ **Neue Funktion: Transportkomponenten - Physische Speicherorte**
  - Queue-Datenbank und Queue-Logs Dokumentation
  - Message-Tracking-Logs Pfade und Konfiguration
  - SMTP-Protokolllogs für Sende- und Empfangsverbindungen
  - Safety-Net Konfiguration und Datenbereiche
  - Standard-Speicherorte Übersichtstabelle
- 🐛 Abschnittsnummern korrigiert und aktualisiert
- 📚 Dokumentation und Versionsinformationen aktualisiert

**EN**
- ✨ **New Feature: Transport Components - Physical Locations**
  - Queue database and queue logs documentation
  - Message tracking log paths and configuration
  - SMTP protocol logs for send and receive connectors
  - Safety-Net configuration and data areas
  - Default locations overview table
- 🐛 Fixed and updated section numbers
- 📚 Updated documentation and version information

## [1.0] - 2026-06-03

**DE**
- Erstveröffentlichung: Initial release mit vollständiger Dokumentationserstellung für Exchange On-Premises (Exchange 2019 & SE)
- Unterstützung für Exchange Server 2019 und Exchange Server Subscription Edition (SE)
- CIM/DCOM Fallback für robuste Netzwerk-Kommunikation
- Exchange Emergency Mitigation Service (EEMS) Monitoring
- Sicherheits- und Compliance-Dokumentation
- Multiple Export-Formate (HTML, PDF, Markdown)
- Grafische Benutzeroberfläche zur Server-Auswahl

**EN**
- Initial release with complete documentation generation for Exchange On-Premises (Exchange 2019 & SE)
- Support for Exchange Server 2019 and Exchange Server Subscription Edition (SE)
- CIM/DCOM fallback for robust network communication
- Exchange Emergency Mitigation Service (EEMS) monitoring
- Security and compliance documentation
- Multiple export formats (HTML, PDF, Markdown)
- Graphical user interface for server selection
