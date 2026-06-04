# 📋 Installation & Setup

Ausführliche Anleitung zur Installation und Einrichtung des Exchange Dokumentations-Tools.

---

## ✅ Systemanforderungen

### Hardware
- **Prozessor:** Beliebig (Mind. 1 GHz empfohlen)
- **RAM:** Min. 2 GB, 4+ GB empfohlen
- **Festplatte:** Min. 500 MB für Installation + Ausgabedateien
- **Netzwerk:** Verbindung zu Exchange-Servern erforderlich

### Software
| Komponente | Version | Bemerkung |
|---|---|---|
| **PowerShell** | 5.1+ oder 7.x | Windows PowerShell oder PowerShell Core |
| **Windows** | Server 2016+ | Auf Client oder Server-OS läufbar |
| **Exchange Snap-In** | 2019 oder SE | Exchange Management Shell |
| **AD Modul** | Beliebig | Remote AD-Modul möglich |

### Netzwerk & Zugriff
- ✅ Netzwerk-Zugriff zu allen zu dokumentierenden Exchange-Servern
- ✅ CIM/RPC (DCOM) oder WinRM aktiviert auf Zielservern
- ✅ Administratorrechte auf den Exchange-Servern
- ✅ Firewall erlaubt Port 135 (RPC/DCOM) oder 5985/5986 (WinRM)

### Berechtigungen
```
- Exchange Organization Management (mindestens View-Only)
- Active Directory Reader (für AD-Abfragen)
- Lokale Administrator-Rechte auf dem ausführenden Server
```

---

## 🚀 Schritt-für-Schritt Installation

### Schritt 1: Repository klonen

```powershell
# Navigiere zum Zielverzeichnis
cd C:\Scripts

# Klone das Repository
git clone https://github.com/RoccoAmmon/Exchange-Documentation.git
cd Exchange-Documentation

# Optional: Zum neuesten Release wechseln
git checkout v1.0
```

### Schritt 2: Ausführungsrichtlinie setzen (falls nötig)

```powershell
# Nur für aktuelle PowerShell-Session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# ODER dauerhaft für den Benutzer
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Schritt 3: Exchange Module laden

**Auf einem Exchange Server:**
```powershell
# Exchange Management Shell öffnen
# Oder manuell laden:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

**Remote (falls nötig):**
```powershell
# Remote-Verbindung zu Exchange aufbauen
$Session = New-PSSession -ComputerName EX01 -ConfigurationName Microsoft.Exchange
Import-PSSession $Session
```

### Schritt 4: Skript ausführen

```powershell
# Mit grafischer Oberfläche
.\Exchange_Dokumentation.ps1

# ODER direkt mit Parametern
.\Exchange_Dokumentation.ps1 -ExchangeServers @('EX01','EX02') -CompanyName "Meine Firma"
```

---

## ⚡ Schnellstart-Befehle

### Einfachste Variante (GUI)
```powershell
.\Exchange_Dokumentation.ps1
```

### Automatisch (ohne GUI)
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01') `
  -CompanyName 'Contoso' `
  -OutputPath 'D:\Reports' `
  -NoGui
```

### Mit allen Formaten
```powershell
.\Exchange_Dokumentation.ps1 `
  -ExchangeServers @('EX01','EX02') `
  -OutputFormats @('HTML','PDF','Markdown') `
  -OutputPath 'C:\Reports'
```

---

## 🔧 Umgebungsvariablen (optional)

```powershell
# Ausgabepfad als Variable setzen
$env:EXDOC_OUTPUT = 'D:\Reports'

# In Skript verwenden
.\Exchange_Dokumentation.ps1 -OutputPath $env:EXDOC_OUTPUT
```

---

## ❌ Häufige Installationsprobleme

### Problem: "Snap-In nicht gefunden"
```
❌ Exception: "Snap-In nicht registriert"
```

**Lösung:**
```powershell
# Auf einem Exchange-Server ausführen
# Oder prüfen, ob Exchange installiert ist:
Test-Path 'C:\Program Files\Microsoft\Exchange Server\V15\bin\ExchangeManagementShell.psc1'
```

### Problem: "Execution Policy verhindert Ausführung"
```
❌ Exception: "Datei konnte nicht geladen werden. Ausführung von Skripten ist auf diesem System deaktiviert."
```

**Lösung:**
```powershell
# Temporär für diese Session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Oder mit Admin-Rechten starten und ausführen
Start-Process powershell -ArgumentList "Set-ExecutionPolicy Bypass; & '.\Exchange_Dokumentation.ps1'" -Verb RunAs
```

### Problem: "Keine Verbindung zu Servern"
```
❌ Exception: "Keine Verbindung zu EX01 möglich"
```

**Lösung:**
```powershell
# Netzwerk prüfen
Test-Connection -ComputerName 'EX01' -Count 2

# Firewall prüfen (DCOM/RPC)
Test-NetConnection -ComputerName 'EX01' -Port 135

# WinRM Status prüfen
Test-WSMan -ComputerName 'EX01'
```

---

## 📦 Git-Integration (Optional)

### Repository aktualisieren
```powershell
cd Exchange-Documentation
git pull origin main
```

### Release-Version installieren
```powershell
git fetch origin
git checkout v1.0
```

---

## ✨ Installationsverifikation

Nach der Installation prüfen:

```powershell
# 1. Skript existiert?
Test-Path '.\Exchange_Dokumentation.ps1'

# 2. Exchange-Snap-In verfügbar?
Get-PSSnapin -Registered | Select-Object Name

# 3. AD-Modul verfügbar?
Get-Module -ListAvailable -Name ActiveDirectory

# 4. Ausgabeverzeichnis erstellen?
Test-Path 'C:\ExchangeDoku' -PathType Container
```

---

## 🎯 Nächste Schritte

Nach erfolgreicher Installation:

1. ➡️ Gehe zur [**Bedienung**](Bedienung)-Seite
2. ➡️ Schaue dir [**Beispiele**](Beispiele) an
3. ➡️ Starte deine erste Dokumentation!

---

## 🆘 Weitere Hilfe

- 🔍 [Troubleshooting Guide](Troubleshooting)
- ❓ [FAQ](FAQ)
- 🐛 [GitHub Issues](https://github.com/RoccoAmmon/Exchange-Documentation/issues)
