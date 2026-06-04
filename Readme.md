![Status](https://img.shields.io/badge/status-Release-blue)
![Version](https://img.shields.io/badge/version-1.0-blue)

# Exchange Dokumentations-Tool (On-Premises)

Dieses Repository enthält ein PowerShell-Skript zur automatisierten Erstellung einer umfassenden HTML-Dokumentation von Microsoft Exchange On-Premises Umgebungen (Exchange 2019 & Subscription Edition).

**Zielgruppe:** Exchange-Administratoren, Consultants und Auditoren, die eine vollständige Bestandsaufnahme und Dokumentation ihrer Exchange-Umgebung benötigen.

**Version:** 1.0

**Wichtig:** Vor dem Einsatz in produktiven Umgebungen die Skripte in einer Testumgebung prüfen.

## Features
- Ermittelt Hardware- und Systeminformationen per CIM (mit DCOM-Fallback)
- Liefert Exchange-spezifische Konfigurationen: Server, Dienste, Datenbanken, DAG, virtuelle Verzeichnisse
- Erfasst Zertifikate, Transport-Regeln, Mailflow, Connectoren und Policies
- Exchange Emergency Mitigation Service (EEMS) Monitoring
- HTML-Ausgabe mit Inhaltsverzeichnis für einfache Weiterverarbeitung

## Voraussetzungen
- PowerShell 5.1 oder PowerShell 7.x
- Exchange Management Shell / ExchangeManagementTools oder Exchange SE Modul
- Active Directory PowerShell-Modul (für AD-bezogene Abfragen)
- Administratorrechte auf den Zielservern

## Schnellstart
1. Repository klonen

```powershell
git clone <your-repo-url>
cd Exchange-Documentation
```

2. Skript lokal ausführen (Beispiel)

```powershell
.\Exchange_Dokumentation.ps1 -ExchangeServers @('EX01','EX02') -CompanyName 'Contoso GmbH' -OutputPath 'C:\Temp\ExchangeDoku'
```

3. Ergebnis: HTML-Datei im angegebenen `OutputPath` mit vollständiger Dokumentation.

## Changelog
- v1.0 – Erstveröffentlichung (Initial Release)

## Lizenz
Dieses Projekt ist lizenziert unter der MIT-Lizenz.

## Kontakt
Autor: Rocco Ammon
Repository: https://github.com/RoccoAmmon/Exchange-Documentation
