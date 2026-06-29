# Changelog

Alle signifikanten Änderungen am Projekt werden hier dokumentiert.

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
