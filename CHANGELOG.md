# Changelog

Alle signifikanten Änderungen am Projekt werden hier dokumentiert.

## [1.9] - 2026-07-15
- ✨ **Exchange Build-Versionskatalog & Online-Abgleich**
  - Umfassende statische Lookup-Tabelle (SE, 2019 CU15-CU12, 2016 CU23, 2013 CU23)
  - Automatischer Online-Abgleich mit Microsoft Learn Build-Tabelle
  - Freundlicher Produktname (z.B. "Exchange Server SE RTM Jul26SU")
  - Update-Status: ⚠️ Warnung bei veralteter Version + Anzeige der aktuellsten Version
  - Link zum aktuellsten Build direkt in der HTML-Tabelle
- � **Support-End-Erkennung** – Markiert beendeten Support (2013/2016/2019) als kritisch
- �🔄 **Online-Versionscheck** – Einmaliger Abruf von MS Learn beim Skriptstart
- 📊 **Neue Spalten in Exchange Server Übersicht:** ProductName, LatestBuild, Status
- 📊 **Neue Spalten in Exchange Version & Build:** ProductName, LatestBuild, Status
- ♻️ **Fallback auf statischen Katalog** bei fehlender Internetverbindung

## [1.8] - 2026-07-10
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

## [1.7] - 2026-07-10
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

## [1.6] - 2026-07-10
- ✨ **7 neue Dokumentationsbereiche hinzugefügt**
  - Message Queue Analyse (Warteschlangen, Status, Nachrichtenanzahl)
  - Calendar & Resource Mailbox (Raum-/Ressourcenpostfächer, Buchungsoptionen)
  - Exchange Archive (Archivpostfächer, Quotas, Auto-Expanding Archive)
  - Exchange Message Size Limits (Org, Connector, Remote Domain, Benutzer)
  - Exchange Partner Applications (SharePoint, Skype, CRM)
  - Exchange Federated Sharing (Federation Trust, Organization Relationships)
  - OAuth / Certificate Based Auth (Auth Server, Zertifikate, CBA)

## [1.5] - 2026-06-30
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

## [1.4] - 2026-06-29
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

## [1.3] - 2026-06-10
- 🐛 **Verbesserte Exchange-Edition-Erkennung**
  - Erkennung jetzt primär über Versionsnummer (15.0/15.1/15.2) statt nur über Build-Ranges
  - Zusätzliche Unterstützung für Exchange 2013 (Version 15.0) und Exchange 2016 (Version 15.1)
  - Sauberere Fallback-Logik (kein hartes "2019"-Fallback mehr)
  - Konsistente Logik zwischen `Get-ExchangeEdition` und `Get-ExchangeServerOverview`
- 📚 Versionsinformationen in Header und NOTES aktualisiert

_Hinweis: Die Version 1.2 (TLS/SSL Konfiguration) wurde übersprungen und nicht als offizielles Release veröffentlicht._

## [1.1] - 2026-06-04
- ✨ **Neue Funktion: Transportkomponenten - Physische Speicherorte**
  - Queue-Datenbank und Queue-Logs Dokumentation
  - Message-Tracking-Logs Pfade und Konfiguration
  - SMTP-Protokolllogs für Sende- und Empfangsverbindungen
  - Safety-Net Konfiguration und Datenbereiche
  - Standard-Speicherorte Übersichtstabelle
- 🐛 Abschnittsnummern korrigiert und aktualisiert
- 📚 Dokumentation und Versionsinformationen aktualisiert

## [1.0] - 2026-06-03
- Erstveröffentlichung: Initial release mit vollständiger Dokumentationserstellung für Exchange On-Premises (Exchange 2019 & SE)
- Unterstützung für Exchange Server 2019 und Exchange Server Subscription Edition (SE)
- CIM/DCOM Fallback für robuste Netzwerk-Kommunikation
- Exchange Emergency Mitigation Service (EEMS) Monitoring
- Sicherheits- und Compliance-Dokumentation
- Multiple Export-Formate (HTML, PDF, Markdown)
- Grafische Benutzeroberfläche zur Server-Auswahl
