# Changelog

Alle signifikanten Änderungen am Projekt werden hier dokumentiert.

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
