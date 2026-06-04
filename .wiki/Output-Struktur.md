# 📊 Output-Struktur

Detaillierter Aufbau und Struktur der generierten Dokumentationen.

---

## 📋 HTML-Struktur

Das HTML-Dokument folgt folgende Struktur:

```
Exchange_Dokumentation_20260604_120000.html
│
├─ COVER PAGE
│  ├─ Titel & Untertitel
│  ├─ Firmenname
│  ├─ Erstellungsdatum
│  ├─ Dokumentierte Server
│  └─ Version
│
├─ TABLE OF CONTENTS
│  └─ Automatisch generierte Inhaltslinks
│
├─ SUMMARY
│  ├─ Dokumentierte Server
│  ├─ Erstellungsdatum & -zeit
│  ├─ Erstellter von & Computer
│  ├─ Gesamtzahl Sektionen
│  ├─ Fehleranzahl
│  └─ Warnungsanzahl
│
├─ HARDWARE-INFORMATIONEN
│  ├─ Server: EX01
│  │  ├─ System-Übersicht
│  │  ├─ Betriebssystem
│  │  ├─ Prozessor(en)
│  │  ├─ Logische Laufwerke
│  │  ├─ Physische Festplatten
│  │  ├─ Pagefile
│  │  ├─ Netzwerkkonfiguration
│  │  ├─ Exchange Version (Registry)
│  │  ├─ Windows Hotfixes
│  │  └─ Exchange Dienste
│  │
│  └─ Server: EX02
│     └─ (wie oben)
│
├─ EXCHANGE PATCHES & BUILD
│  └─ Patch-Informationen pro Server
│
├─ FSMO-ROLLEN
│  ├─ Gesamtstruktur-FSMO-Rollen
│  ├─ Domänen-FSMO-Rollen
│  └─ (weitere Rollen)
│
├─ AD INFORMATIONEN & SCHEMA
│  ├─ Gesamtstruktur
│  ├─ Domäne
│  ├─ Schema-Version
│  ├─ Exchange Organisation
│  ├─ Domänencontroller
│  ├─ AD Sites
│  └─ AD Site Links
│
├─ EXCHANGE SERVER ÜBERSICHT
│  └─ Alle Exchange Server mit
│     ├─ Edition
│     ├─ Build
│     ├─ Role
│     ├─ Site
│     └─ Status
│
├─ ORGANISATIONSKONFIGURATION
│  ├─ Exchange Organisationskonfiguration
│  └─ Transport-Konfiguration
│
├─ EXCHANGE URLS
│  ├─ OWA Virtual Directory
│  ├─ ECP Virtual Directory
│  ├─ ActiveSync Virtual Directory
│  ├─ EWS (Exchange Web Services)
│  ├─ MAPI Virtual Directory
│  ├─ OAB Virtual Directory
│  ├─ Autodiscover
│  ├─ Outlook Anywhere
│  └─ EEMS (Exchange Emergency Mitigation)
│
├─ DATENBANKEN & DAG
│  ├─ Mailbox-Datenbanken
│  ├─ Datenbankkopie-Status (DAG)
│  └─ DAG-Konfiguration
│
├─ [weitere Sektionen]
│  └─ ...
│
└─ FOOTER
   ├─ Dokument-Info
   ├─ Erstellungsdatum & -zeit
   ├─ Seite X von Y
   └─ Version des Skripts
```

---

## 🎨 HTML-Styling

Das HTML-Dokument enthält:

### CSS-Klassen
```css
.cover-page      /* Titelseite */
.toc             /* Inhaltsverzeichnis */
.summary-box     /* Zusammenfassung */
.section         /* Abschnitte */
.section h2      /* Überschriften */
.section table   /* Tabellen */
.footer          /* Fußzeile */

.server-break    /* Server-Trennlinie */
.even / .odd     /* Zebra-Striping in Tabellen */
.no-data         /* Keine Daten Nachricht */
.error           /* Fehlermeldung */
.warning         /* Warnung */
```

### Farben & Formatierung
```
Primär: #0078D4 (Microsoft Blau)
Header: #333333
Text: #000000
Tabellen: Alternierend (weiß/grau)
```

---

## 📄 Sektionen-Übersicht

### 1. Hardware-Informationen
```
Für JEDEN dokumentierten Server:
- System-Übersicht (Hersteller, Modell, VM-Erkennung)
- Betriebssystem (Windows-Version, Build, Install-Datum)
- Prozessor (Kerne, Takt, Hyperthreading)
- RAM (Gesamt, Frei, Belegt %)
- Festplatten (Größe, Frei, Belegt %)
- Netzwerk (IP, DNS, DHCP)
- Pagefile
- Hotfixes (letzte 30)
- Exchange Dienste Status
```

### 2. Patch-Informationen
```
- Exchange Server Details (Edition, Build, Role)
- Windows-Patches (Hotfix-Liste)
- Exchange Build-Nummer
- Schema-Version
```

### 3. FSMO-Rollen
```
Gesamtstruktur:
- Schema Master
- Domain Naming Master

Domäne:
- PDC Emulator
- RID Master
- Infrastructure Master
```

### 4. AD-Informationen
```
- Gesamtstruktur-Name & Funktionsebene
- Domäne & NetBIOS-Name
- Global Catalogs
- FSMO-Rollen
- Domänencontroller
- Schema-Version
- Exchange Schema-Version
- Sites & Site Links
```

### 5. Exchange Server Übersicht
```
Pro Server:
- Name
- Edition (2019/SE)
- Build/Version
- Rollen (Mailbox/CAS/Edge)
- Site
- Erstellt/Geändert Datum
```

### 6. Organisationskonfiguration
```
- Exchange Org Config
- Transport-Config
- Max Send/Receive Size
- Externe Postmaster-Adresse
- TLS-Sicherheitslisten
- Shadow Redundancy
- Journaling Config
```

### 7. Exchange URLs
```
Pro virtualem Verzeichnis:
- Server
- Intern-URL
- Extern-URL
- Authentifizierungsmethoden
- (für jedes VDirectory: OWA, ECP, EAS, EWS, MAPI, OAB)

Plus:
- Autodiscover Service
- Outlook Anywhere
- EEMS-Status
```

### 8. Datenbanken & DAG
```
Mailbox-Datenbanken:
- Name
- Server
- Größe
- Mailbox-Anzahl
- Backup-Status
- Quotas
- Aufbewahrung

DAG-Status:
- DAG-Name
- Mitglieder
- Witness-Server
- Copy-Status pro Datenbank
- Content Index Status
```

### 9. Öffentliche Ordner
```
- Public Folder Konfiguration
- Statistiken
- Replikation
```

### 10. Transport & Routing
```
- Sende-Connectoren
- Empfangs-Connectoren
- Remote Domains
- Accepted Domains
- Transport-Regeln
- Journal Rules
- Message Tracking Config
```

### 11. Zertifikate
```
Pro Zertifikat:
- Thumbprint
- Subject
- Issuer
- Gültig von/bis
- Ablauf-Warnung
- Bindings (HTTP/SMTP)
- Verwendete Dienste
```

### 12. Sicherheit & Compliance
```
- TLS-Einstellungen
- Auth-Methoden
- DLP-Richtlinien
- Litigation Hold
- Retention Policies
- Journal Rules
- RBAC-Rollen & Mitglieder
```

### 13. E-Mail-Policies
```
- E-Mail-Adressrichtlinien
- Adresslisten
- Globale Adressliste (GAL)
- Offline-Adressbuch
```

### 14. Sonstige
```
- Mobile Device Policies
- Throttling Policies
- Mailbox-Statistiken
- Event Logs (7 Tage)
- Hybrid-Konfiguration
- SMTP-Relay
```

---

## 📊 Tabellen-Formate

### Standard-Tabelle
```
| Spalte 1 | Spalte 2 | Spalte 3 |
|----------|----------|----------|
| Wert A   | Wert B   | Wert C   |
| Wert D   | Wert E   | Wert F   |
```

### Server-Sektion-Header
```
═ Server: EX01 ═
  System-Übersicht
  ├─ Hersteller: Dell Inc.
  ├─ Modell: PowerEdge R750
  ├─ RAM: 256 GB
  └─ CPU: Xeon Gold 5320
```

### Status-Indikatoren
```
✅ OK / aktiv
⚠️  WARNUNG / zu beachten
❌ FEHLER / Problem
◎ UNBEKANNT / nicht verfügbar
```

---

## 📄 PDF-Struktur

PDF hat **identische Struktur** wie HTML mit:

✅ Seitennummerierung (unten)  
✅ Kopf-/Fußzeile mit Firma & Datum  
✅ Inhaltsverzeichnis mit Seitennummerierung  
✅ Seitenumbrüche nach Hauptsektionen  
✅ Druckoptimierte Farben & Abstände  

---

## 🔤 Markdown-Struktur

Markdown hat **vereinfachte Struktur**:

```markdown
# Dokumenttitel

## Inhaltsverzeichnis
- [Sektion 1](#sektion-1)
- [Sektion 2](#sektion-2)

## Sektion 1
### Server EX01
#### System-Übersicht
| Feld | Wert |
| --- | --- |
| Hersteller | Dell |

### Server EX02
...
```

**Vorteil:** GitHub-kompatibel, Wiki-freundlich

---

## 📏 Dateigrößen (typisch)

| Kombination | HTML | PDF | MD |
|---|---|---|---|
| 1 Server, Basis-Sektionen | 2-5 MB | 5-10 MB | 300-500 KB |
| 1 Server, alle Sektionen | 5-10 MB | 10-20 MB | 1-2 MB |
| 3 Server, alle Sektionen | 15-30 MB | 30-60 MB | 3-6 MB |
| 10+ Server, alle Sektionen | 50+ MB | 100+ MB | 10+ MB |

---

## 🔍 In Word öffnen

### Schritte:
1. HTML-Datei mit Doppelklick öffnen
2. **Oder:** Word → Datei → Öffnen → HTML-Datei

### Anpassungen in Word:
```
- ✅ Kopf-/Fußzeilen hinzufügen
- ✅ Seitenzahlen anpassen
- ✅ Firmenllogo einfügen
- ✅ Schriftart ändern
- ✅ Seite formatieren
```

### Speichern:
- Datei → Speichern unter → **Word-Format (.docx)**
- oder **PDF** für Versand

---

## 📊 Beispiel-Output

### Minimales HTML (gekürzt):
```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Exchange Dokumentation</title>
  <style>...</style>
</head>
<body>
  <div class="cover-page">
    <h1>Exchange Server Dokumentation</h1>
    ...
  </div>
  
  <div class="toc">
    <h2>Inhaltsverzeichnis</h2>
    ...
  </div>
  
  <div class="section">
    <h2>Hardware-Informationen</h2>
    <table>...</table>
  </div>
  
  ...
</body>
</html>
```

---

## 🎯 Nächste Schritte

- ➡️ [Beispiele](Beispiele) – Praktische Szenarien
- ➡️ [Best Practices](Best-Practices) – Optimierungen
- ➡️ [FAQ](FAQ) – Häufige Fragen
