# Changelog / Versionshistorie

**DE:** Alle signifikanten Änderungen am Projekt werden hier dokumentiert.
**EN:** All notable changes to this project are documented here.

---

## [2.1] - 2026-11-01

**DE**
- 🔄 **Automatisches Self-Update**
  - Versionsprüfung gegen die GitLab-Releases-API (asynchron, blockiert den Start nicht)
  - Update-Badge in der GUI: **Klick = Update installieren**, **Rechtsklick = Release-Seite**
  - Konsolen-Update mit `j`-Bestätigung vor dem Datenabgleich
  - Sicherheitsverifikation des Update-Pakets: SHA256, Dateigröße, PowerShell-Syntax, Versionsnummer
  - Automatisches Backup der aktuellen Datei + Rollback-Hilfsprozess bei Fehlern
  - Neustart mit identischen Parametern nach dem Update
  - Neue Parameter `-NoUpdateCheck` und `-NoSelfUpdate` zum Deaktivieren
- 🌐 **Vollständige DE/EN-Lokalisierung abgeschlossen**
  - 957 Übersetzungspaare (vorher 685), 0 verbleibende deutsche Stellen im EN-Modus
  - Log- und Dokument-Ausgaben zu 100 % übersetzt (356/356 Log, 314/314 DOC)
- 🚀 **Erreichbarkeits-Check überarbeitet**
  - Parallele Prüfung aller Server (Jobs) statt sequenzieller ICMP-Pings
  - TCP-Fallback (Port 135/443) wenn ICMP blockiert ist
  - Hartes 15s-Gesamt-Timeout – kein Hängenbleiben mehr
  - Ergebnis mit Prüfmethode (ICMP/TCP/135/TCP/443)
- 📦 **Exchange Build-Katalog aktuell** (Stand Oct26SU)
  - SE 15.2.2562.56, 2019 CU15 15.2.1748.55, 2019 CU14 15.2.1544.50, 2016 CU23 15.1.2507.77
  - Korrekte Microsoft-Download-Links (KB-Prüfung aller Einträge)
  - Fehlende Builds Aug26/Sep26/Sep26v2/Oct26 ergänzt (114 Einträge)
- 🔧 **MS-Learn-Online-Parser repariert** (war nie funktionsfähig)
  - Abschnittsweise Auswertung über h2-IDs statt Header-Text-Matching
  - Sprachunabhängig (de/en), `&nbsp;`- und `<br>`-tolerant
  - Regressionsschutz: Online-Wert wird nur übernommen wenn er wirklich neuer ist
- 🐛 **Fehlerbehebungen**
  - `PSRemotingJob.Remove()`-Fehler beim Erreichbarkeits-Check behoben
  - ADSI-Admin-Gruppen-Lookup über SID `S-1-5-32-544` mit Fallbacks
  - BSI RSS-Feed robust (`-UseBasicParsing -TimeoutSec`)
  - PDF-Export hart gemacht (Edge/Chrome-Headless-Fallback, Retry, Polling)

**EN**
- 🔄 **Automatic Self-Update**
  - Version check against GitLab Releases API (async, non-blocking)
  - GUI update badge: **click = install update**, **right-click = release page**
  - Console update with `j` confirmation before data collection
  - Update package verification: SHA256, file size, PowerShell syntax, version number
  - Automatic backup of current file + rollback helper process on failure
  - Restart with identical parameters after update
  - New `-NoUpdateCheck` and `-NoSelfUpdate` parameters to disable
- 🌐 **Complete DE/EN localization finished**
  - 957 translation pairs (was 685), zero remaining German strings in EN mode
  - Log and document output 100% translated (356/356 log, 314/314 doc)
- 🚀 **Reachability check reworked**
  - Parallel check of all servers (jobs) instead of sequential ICMP pings
  - TCP fallback (port 135/443) when ICMP is blocked
  - Hard 15s total timeout – no more hanging
  - Result includes check method (ICMP/TCP 135/TCP 443)
- 📦 **Exchange build catalog up to date** (as of Oct26SU)
  - SE 15.2.2562.56, 2019 CU15 15.2.1748.55, 2019 CU14 15.2.1544.50, 2016 CU23 15.1.2507.77
  - Correct Microsoft download links (KB verified for all entries)
  - Missing Aug26/Sep26/Sep26v2/Oct26 builds added (114 entries)
- 🔧 **MS Learn online parser fixed** (was never functional)
  - Section-based evaluation via h2 IDs instead of header text matching
  - Language independent (de/en), `&nbsp;` and `<br>` tolerant
  - Regression guard: online value only accepted if it is actually newer
- 🐛 **Bug fixes**
  - `PSRemotingJob.Remove()` error during reachability check fixed
  - ADSI admin group lookup via SID `S-1-5-32-544` with fallbacks
  - BSI RSS feed hardened (`-UseBasicParsing -TimeoutSec`)
  - PDF export hardened (Edge/Chrome headless fallback, retry, polling)

## [2.0] - 2026-07-16

**DE**
- 🌐 **Sprachumschaltung Deutsch/Englisch (DE/EN)**
  - Neuer Parameter `-Language` mit `[ValidateSet("DE","EN")]` (Standard: DE)
  - Komplette Lokalisierung des Skripts: Konsole, Logs, Fehlermeldungen
  - Lokalisierte Reports: HTML, PDF und Markdown in Deutsch oder Englisch
  - GUI-Sprachumschalter (Deutsch/Englisch) direkt im Startdialog
  - Automatischer Neustart der GUI bei Sprachwechsel
- 🧩 **Lokalisierungs-Engine**
  - `$script:LocUI` – Wörterbuch für feste UI-/Report-Strings (DE/EN)
  - `$script:LocContentPairs` – 685 DE|EN-Übersetzungspaare für dynamische Inhalte
  - `Get-LocText` – Lokalisierte feste Strings mit Format-Argumenten
  - `ConvertTo-LocalizedText` – Übersetzt deutschen Klartext (no-op bei DE)
  - `ConvertTo-LocalizedHtml` – Übersetzt nur HTML-Textknoten, lässt Tags/Attribute intakt
  - `Initialize-LocalizationMap` – Baut Wörterbuch + kombinierten Regex (längste Begriffe zuerst)
- 📁 **Lokalisierte Dateinamen** – `Exchange_Dokumentation`/`Exchange_Documentation` je nach Sprache
- 🖥️ **GUI** – Sprach-Karte mit Deutsch/Englisch-Toggle, lokalisierte Beschriftungen
- 🐛 **XAML-Validierung** – Platzhalter-basierte Lokalisierung, XML-gültig nach Ersetzung

**EN**
- 🌐 **German/English Language Switch (DE/EN)**
  - New `-Language` parameter with `[ValidateSet("DE","EN")]` (default: DE)
  - Full script localization: console, logs, error messages
  - Localized reports: HTML, PDF and Markdown in German or English
  - GUI language toggle (German/English) directly in the start dialog
  - Automatic GUI restart on language switch
- 🧩 **Localization Engine**
  - `$script:LocUI` – dictionary for fixed UI/report strings (DE/EN)
  - `$script:LocContentPairs` – 685 DE|EN translation pairs for dynamic content
  - `Get-LocText` – localized fixed strings with format arguments
  - `ConvertTo-LocalizedText` – translates plain German text (no-op for DE)
  - `ConvertTo-LocalizedHtml` – translates only HTML text nodes, keeps tags/attributes intact
  - `Initialize-LocalizationMap` – builds dictionary + combined regex (longest terms first)
- 📁 **Localized filenames** – `Exchange_Dokumentation`/`Exchange_Documentation` depending on language
- 🖥️ **GUI** – Language card with German/English toggle, localized labels
- 🐛 **XAML validation** – placeholder-based localization, XML-valid after replacement

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
