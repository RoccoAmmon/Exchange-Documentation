# 📋 Changelog

Versionshistorie und Update-Informationen.

---

## 📌 Version 1.0 (Release) – 2026-06-04

**🎉 Erste öffentliche Veröffentlichung**

### ✨ Neue Features
- ✅ Vollständige Exchange 2019 Unterstützung
- ✅ Exchange Subscription Edition (SE) Unterstützung
- ✅ Automatische Edition-Erkennung (2019 vs. SE)
- ✅ CIM-Session mit DCOM-Fallback (kein WinRM-Abhängigkeit)
- ✅ Remote-Registry-Zugriff (.NET-basiert)
- ✅ Grafische Benutzeroberfläche (WPF)
- ✅ Multi-Format Export: HTML, PDF, Markdown
- ✅ Detaillierte Hardware-Erfassung
- ✅ Exchange-Konfiguration dokumentieren
- ✅ Zertifikat-Ablauf-Tracking
- ✅ Active Directory Integration
- ✅ FSMO-Rollen-Erfassung
- ✅ DAG-Status & Datenbanken
- ✅ Transport-Regeln & Policies
- ✅ Sicherheits- & Compliance-Details
- ✅ Exchange Emergency Mitigation Service (EEMS) Monitoring
- ✅ Detailliertes Logging
- ✅ Umfassendes Inhaltsverzeichnis (HTML)

### 🔧 Komponenten
- `Exchange_Dokumentation.ps1` – Hauptskript (v1.0)
- `README.md` – Projekt-Dokumentation
- `.github/workflows/release.yml` – GitHub Actions Workflow
- `scripts/create-release.ps1` – Release-Helper
- `CHANGELOG.md` – Versionshistorie

### 📚 Dokumentation
- **Wiki mit 9 Seiten:**
  - Home – Übersicht & Quick Links
  - Installation – Schritt-für-Schritt Setup
  - Bedienung – Parameter & Optionen
  - Konfiguration – Erweiterte Einstellungen
  - Beispiele – 10+ Szenarien
  - Best Practices – Optimierungen & Tipps
  - FAQ – Häufige Fragen
  - Troubleshooting – Probleme & Lösungen
  - Output-Struktur – Report-Aufbau

### 📊 Unterstützte Dokumentationsinhalte
```
Hardware & System:
✅ CPU, RAM, Festplatte, Netzwerk
✅ Windows-Version & Hotfixes
✅ Virtualisierungs-Erkennung

Exchange:
✅ Server-Übersicht
✅ Virtuelle Verzeichnisse (OWA, ECP, EWS, etc)
✅ Zertifikate
✅ Datenbanken & DAG
✅ Transport-Regeln
✅ Compliance & DLP

Active Directory:
✅ FSMO-Rollen
✅ Schema-Version
✅ Domänen & Sites
✅ Domänencontroller
```

### 🛠️ Technische Details
- **Sprache:** PowerShell 5.1+
- **Abhängigkeiten:** Exchange Management Shell, AD Module
- **Output:** HTML (mit CSS), PDF (optional), Markdown
- **Plattformen:** Windows Server, Client OS (mit Exchange Snap-In)

### 📦 Release-Assets
- Source Code auf GitHub
- Release-Tag: `v1.0`
- Workflow-Automation für Releases

### 🐛 Bekannte Einschränkungen
- Mailbox-Inhalte werden nicht gelesen (nur Statistiken)
- Passwörter nicht enthalten
- Remote-Ausführung komplexer als lokal
- PDF-Export benötigt zusätzliche Tools

### 🔒 Sicherheit
- ✅ Nur Read-Only Operationen
- ✅ Keine Daten außer Haus
- ✅ Open Source Code
- ✅ Lokal ausgeführt

---

## 📅 Version-Timeline

```
2026-03-05: Interne Entwicklung (v0.1 - v0.8)
2026-04-01: Alpha-Getestet
2026-05-15: Beta mit limitiertem Feedback
2026-06-04: 🎉 v1.0 Release
```

---

## 🔜 Geplante Features (Roadmap)

### Version 1.1 (Q3 2026)
- [ ] Verbessertes GUI-Design
- [ ] Export in Excel (.xlsx)
- [ ] Comparative Reports (Diff zwischen zwei Reports)
- [ ] Automatische Probleme-Erkennung
- [ ] Performance-Optimierungen

### Version 1.2 (Q4 2026)
- [ ] Multilingual Support (Deutsch/English/Französisch)
- [ ] Cloud-Backup Integration
- [ ] Compliance-Reports (ISO 27001, SOC2)
- [ ] API für Integration mit anderen Tools
- [ ] Docker-Container

### Version 2.0 (2027)
- [ ] Exchange Online Support
- [ ] Hybrid-Mode (On-Prem + Cloud)
- [ ] Web-Dashboard für Reports
- [ ] Automatische Trend-Analyse
- [ ] Predictive Maintenance Alerts

---

## 📝 Version Numbering

Das Projekt folgt **Semantic Versioning**:

```
v[MAJOR].[MINOR].[PATCH]
  ↓        ↓         ↓
  |        |         └─ Bugfixes, kleine Verbesserungen
  |        └────────── Neue Features (rückwärts-kompatibel)
  └───────────────── Breaking Changes
```

**Beispiele:**
- `v1.0.0` – Erste Release
- `v1.1.0` – Neue Features hinzugefügt
- `v1.1.5` – Bugfixes in v1.1
- `v2.0.0` – Major Update (Breaking Changes)

---

## 🔄 Update-Prozess

### Automatische Updates prüfen
```powershell
cd Exchange-Documentation
git fetch origin
git log --oneline -n 5   # Prüfe neue Commits
```

### Auf neue Version updaten
```powershell
# Latest fetchen
git fetch origin

# Zu Version updaten
git checkout v1.1

# oder latest main
git checkout main
git pull origin main
```

### Release-Seite
https://github.com/RoccoAmmon/Exchange-Documentation/releases

---

## 🐛 Bug-Berichte & Fixes

### Bekannte Bugs (v1.0)
- Keine derzeit bekannt

### Fehlerbehobene Bugs in 1.0
- (Keine Vorgängerversion)

### Reporting
Fehler melden unter:
https://github.com/RoccoAmmon/Exchange-Documentation/issues

---

## 📞 Support & Feedback

### Feedback geben
- 💬 GitHub Discussions
- 🐛 GitHub Issues (für Bugs)
- 🚀 Feature-Requests via Issues

### Community
- 🔗 GitHub Repository
- 📖 Wiki & Documentation
- 💡 Best Practices Forum

---

## 🙏 Danksagungen

**Version 1.0 entstanden durch:**
- Exchange Server Community-Feedback
- Testing durch Early Adopters
- Microsoft Exchange Documentation
- PowerShell Best Practices

---

## 📄 Lizenz & Nutzungsbedingungen

- **Lizenz:** MIT
- **Frei verwendbar:** Persönlich & Kommerziell
- **Open Source:** Code einsehbar & änderbar
- **Keine Garantie:** "Wie vorhanden" (As-Is)

Siehe [LICENSE](https://github.com/RoccoAmmon/Exchange-Documentation/blob/main/LICENSE)

---

## 🔗 Verwandte Projekte

Von demselben Autor:

- 🔗 [Active Directory Health Check](https://github.com/RoccoAmmon/Active-Directory-Health-Check)
- 🔗 [GPO Dokumentation](https://github.com/RoccoAmmon/GPO-Dokumentation)

---

## 📅 Unterstützungszeitraum

| Version | Release | Support bis | Status |
|---|---|---|---|
| v1.0 | 2026-06-04 | 2027-06-04 | 🟢 Active |
| v1.1 | geplant | 2027-06-04 | 🔵 Planned |
| v2.0 | 2027 | TBD | 🔵 Future |

---

## 🎯 Nächste Schritte

- 📖 [Dokumentation](https://github.com/RoccoAmmon/Exchange-Documentation/wiki) lesen
- 🚀 [Installation](Installation) durchführen
- 💡 [Beispiele](Beispiele) ausprobieren
- 🐛 [Issues](https://github.com/RoccoAmmon/Exchange-Documentation/issues) melden
