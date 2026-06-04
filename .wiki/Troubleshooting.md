# 🐛 Troubleshooting

Lösungen für häufige Probleme und Fehlermeldungen.

---

## ⚠️ Häufige Fehlermeldungen

### Fehler: "Snap-In nicht registriert"

```
Exception: System.Management.Automation.PSSnapInException: 
Das angeforderte Snap-In 'Microsoft.Exchange.Management.PowerShell.SnapIn' ist nicht installiert.
```

**Ursachen & Lösungen:**

| Ursache | Lösung |
|---|---|
| **Exchange nicht installiert** | Auf einem Exchange-Server ausführen |
| **Exchange beschädigt** | Exchange reparieren/neu installieren |
| **PowerShell nicht korrekt konfiguriert** | PowerShell neu starten oder als Admin |

**Schritt-für-Schritt:**

```powershell
# 1. Überprüfen, ob auf Exchange-Server
Test-Path 'C:\Program Files\Microsoft\Exchange Server\V15'

# 2. Falls ja, prüfen ob Snap-In registriert
Get-PSSnapin -Registered

# 3. Falls nicht sichtbar, Exchange reparieren:
# Gehe zu Control Panel → Programs → Exchange Server → Repair
```

---

### Fehler: "Zugriff verweigert"

```
Exception: Access Denied
ou cannot access 'REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\...'
```

**Ursachen & Lösungen:**

```powershell
# ✅ Lösung 1: Mit Admin-Rechten ausführen
Start-Process powershell -Verb RunAs
# Danach Skript ausführen

# ✅ Lösung 2: Execution-Policy ändern
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force
.\Exchange_Dokumentation.ps1

# ✅ Lösung 3: Berechtigungen auf Servern prüfen
# Benötigt: Exchange View-Only Admin-Rechte
```

---

### Fehler: "Keine Verbindung zu Server XY"

```
Exception: The network connection is unavailable.
Could not connect to 'EX01'
```

**Checkliste:**

```powershell
# 1. Server online?
ping EX01
# Falls: "Host nicht erreichbar" → Server ist offline

# 2. Firewall blockiert DCOM/RPC (Port 135)?
Test-NetConnection -ComputerName EX01 -Port 135
# Falls: "TCPTestSucceeded : False" → Firewall blockiert

# 3. WinRM aktiviert?
Test-WSMan -ComputerName EX01
# Falls: "Der WSMan Service ist nicht erreichbar" → WinRM aus

# 4. Netzwerk-Problem?
tracert EX01
# Prüfe Hops und Latenz
```

**Lösungen:**

```powershell
# A) Firewall-Regel auf Exchange-Server
# Gehe zu Windows Firewall with Advanced Security
# → Erlauben Sie "Windows Management Instrumentation (WMI)"
# → Erlauben Sie Port 135 (DCOM)

# B) WinRM aktivieren
Enable-PSRemoting -Force

# C) Lokal testen (auf dem Server selbst)
.\Exchange_Dokumentation.ps1 -ExchangeServers @($env:COMPUTERNAME)
```

---

### Fehler: "Exchange Management Tools nicht verfügbar"

```
Exception: Exchange Management Tools not found.
Please run this script on an Exchange server.
```

**Lösung:**

```powershell
# Das Skript MUSS auf einem Exchange-Server ausgeführt werden
# oder mit Remote-Zugriff

# Option 1: Auf einem Exchange-Server ausführen
# Direkter Login zu EX01
# PowerShell öffnen
# Skript ausführen

# Option 2: Remote-Session
$session = New-PSSession -ComputerName EX01
Enter-PSSession $session
# Skript ausführen
```

---

## 🔍 Fehlerdiagnose

### Debugging aktivieren

```powershell
# Verbose-Modus
.\Exchange_Dokumentation.ps1 -Verbose -NoGui

# Debug-Modus
$DebugPreference = "Continue"
.\Exchange_Dokumentation.ps1 -Debug -NoGui

# Full ErrorAction
$ErrorActionPreference = "Stop"
.\Exchange_Dokumentation.ps1 -NoGui
```

### Log-Dateien prüfen

```powershell
# Letzte Log-Datei anzeigen
$logFile = Get-ChildItem "C:\ExchangeDoku\*.log" | 
  Sort-Object LastWriteTime -Descending | Select-Object -First 1

Get-Content $logFile.FullName | Select-String "ERROR" -Context 2

# In Editor öffnen
notepad $logFile.FullName
```

---

## 🌐 Netzwerk-Probleme

### CIM-Session schlägt fehl

```
Exception: CIM Session failed for server 'EX01'
```

**Lösungen:**

```powershell
# 1. Manuell CIM-Session testen
$session = New-CimSession -ComputerName 'EX01' -ErrorAction Stop
# Falls erfolgreich: OK

# 2. Mit DCOM-Fallback testen
$dcomOption = New-CimSessionOption -Protocol Dcom
$session = New-CimSession -ComputerName 'EX01' -SessionOption $dcomOption

# 3. RPC-Port öffnen
# Firewall: Port 135 erlauben
# oder WinRM nutzen (Port 5985/5986)
```

### WinRM ist deaktiviert

```
Exception: The WinRM service is not running
```

**Lösung:**

```powershell
# Auf dem Exchange-Server:
Enable-PSRemoting -Force

# Verifies WinRM service
Get-Service WinRM

# Start service
Start-Service WinRM
```

---

## 💾 Festplatte & Speicher

### "Nicht genug Speicherplatz"

```
Exception: Not enough free disk space
Only X GB available, but Y GB required
```

**Lösungen:**

```powershell
# 1. Speicherplatz prüfen
Get-PSDrive C | Select-Object Used, Free

# 2. Ausgabepfad auf anderesLaufwerk
.\Exchange_Dokumentation.ps1 `
  -OutputPath 'D:\Reports'  # Statt C:\

# 3. Alte Reports löschen
Remove-Item 'C:\ExchangeDoku\*.pdf' -Older Than 30 days
```

---

## 🗄️ Datenbank-Probleme

### "Mailbox-Statistiken nicht verfügbar"

```
Exception: Could not retrieve mailbox statistics
```

**Ursachen & Lösungen:**

```powershell
# 1. Zu viele Mailboxen? → MaxMailboxesForStats anpassen
# Im Skript Line ~115:
$script:MaxMailboxesForStats = 1000  # Erhöhen

# 2. Timeout? → DAG-Problem?
# Datenbank-Status prüfen:
Get-MailboxDatabaseCopyStatus

# 3. Mailbox-Server nicht erreichbar?
Get-MailboxServer | Test-Connection
```

---

## 📄 Output-Probleme

### "HTML-Datei wird nicht korrekt angezeigt"

```
HTML-Datei sieht beschädigt aus oder CSS fehlt
```

**Lösungen:**

```powershell
# 1. Datei-Encoding prüfen
Get-Content 'Report.html' -Encoding UTF8

# 2. In Browser öffnen statt Word
Start-Process "Report.html"

# 3. Browser-Cache löschen
# Browser neuladen (Ctrl+F5)
```

### "PDF-Export funktioniert nicht"

```
Exception: PDF export failed
```

**Lösungen:**

```powershell
# 1. HTML → PDF Konverter prüfen
# Das Skript benötigt einen PDF-Konverter
# (Details im Skript bei Export-Dokument​ToPdf)

# 2. Alternative: Nur HTML exportieren
-OutputFormats @('HTML')
# Danach manuell in Word → Als PDF speichern

# 3. LibreOffice nutzen zum Konvertieren
& 'C:\Program Files\LibreOffice\soffice.exe' --headless --convert-to pdf Report.html
```

---

## 🔗 Remote-Session-Probleme

### "Remote-Session konnte nicht erstellt werden"

```
Exception: New-PSSession failed
Could not create a remote session to 'EX01'
```

**Checkliste:**

```powershell
# 1. Zielserver erreichbar?
Test-Connection -ComputerName 'EX01'

# 2. WinRM aktiviert auf Zielserver?
Test-WSMan -ComputerName 'EX01'

# 3. Firewall erlaubt Port 5985?
Test-NetConnection -ComputerName 'EX01' -Port 5985

# 4. Anmeldeinformationen korrekt?
$cred = Get-Credential
New-PSSession -ComputerName 'EX01' -Credential $cred
```

---

## 🆘 Support & Zusätzliche Hilfe

### Wo bekomme ich Hilfe?

1. **Dieses Wiki** – Volle Dokumentation durchsuchen
2. **Logs prüfen** – `C:\ExchangeDoku\*.log`
3. **GitHub Issues** – Fehler melden
4. **Troubleshooting Chat** – Communities

### Fehler melden

Wenn Sie einen Fehler finden:

1. **Log-Datei** exportieren
2. **Schritte** reproduzieren
3. **GitHub Issue** öffnen: https://github.com/RoccoAmmon/Exchange-Documentation/issues

**Angaben für Issue:**

```
- PowerShell-Version
- Exchange-Version
- OS-Version
- Fehlermeldung (vollständig)
- Log-Datei-Inhalt
- Reproduktionsschritte
```

---

### Performance-Optimierung

Falls das Skript langsam ist:

```powershell
# 1. Sektionen filtern
-Sections @('Hardware','ExchangeServer','Certificates')

# 2. Weniger Server
-ExchangeServers @('EX01')  # Statt 10 Server

# 3. Nur HTML (kein PDF)
-OutputFormats @('HTML')

# 4. Off-Peak ausführen
# Nachts, wenn weniger Last

# 5. Auf lokalem Server ausführen (nicht remote)
```

---

## 📞 Notfall-Support

**Kritischer Fehler?**

```powershell
# 1. Sofort stoppen: Ctrl+C
# 2. Log-Datei speichern
# 3. Fehler reproduzieren
# 4. Issue auf GitHub öffnen
# 5. Kontakt: GitHub Issues
```

---

## 🎯 Nächste Schritte

- ➡️ [FAQ](FAQ) – Häufige Fragen
- ➡️ [Best Practices](Best-Practices) – Optimierungen
- ➡️ [Beispiele](Beispiele) – Weitere Szenarien
