# ❓ FAQ – Häufig gestellte Fragen

Antworten zu häufigen Fragen zum Exchange Dokumentations-Tool.

---

## 🚀 Installation & Setup

### Muss ich auf einem Exchange-Server arbeiten?

**Nein**, aber **empfohlen** für vollständigen Zugriff.

- ✅ **Exchange Management Shell** wird benötigt
- ✅ Dies ist auf Exchange-Servern **installiert**
- ❌ Remote-Ausführung möglich, aber komplizierter
- ✅ **Empfehlung:** Von Admin-Server oder lokal auf Exchange ausführen

---

### Kann ich das Skript von einem Remote-Server ausführen?

**Ja**, aber mit Einschränkungen:

```powershell
# Remote-Session zu Exchange-Server
$session = New-PSSession -ComputerName EX01

# Skript remote ausführen
Invoke-Command -Session $session -ScriptBlock {
    & 'C:\Scripts\Exchange_Dokumentation.ps1' ...
}
```

**Problem:** Komplexer und anfälliger für Fehler. **Empfehlung:** Lokal ausführen.

---

### Welche PowerShell-Version wird benötigt?

**PowerShell 5.1+** oder **PowerShell 7.x**

```powershell
# Version prüfen
$PSVersionTable.PSVersion

# PowerShell 5.1 = OK
# PowerShell 7.x = OK
# PowerShell < 5.0 = NICHT OK
```

---

## ⏱️ Laufzeit & Performance

### Wie lange dauert die Dokumentation?

**Abhängig von Umgebung:**

| Umgebung | Dauer |
|---|---|
| 1 Server (schnelle Sektionen) | 5-10 Min |
| 1 Server (alle Sektionen) | 15-30 Min |
| 3+ Server | 45-120 Min |
| Große Umgebung (10+ Server) | 2-4 Stunden |

**Tipp:** Mit filtrierten Sektionen viel schneller!

---

### Kann ich die Dokumentation abbrechen?

**Ja**: `Ctrl+C` drücken

```powershell
.\Exchange_Dokumentation.ps1 ...
# Während Ausführung
[Ctrl+C]  # Zum Abbrechen
```

**Warnung:** Laufende Berichte werden verworfen.

---

### Warum ist die Dokumentation langsam?

**Häufige Gründe:**

1. **Viele Server** → Reduzieren Sie die Anzahl
2. **Alle Sektionen** → Nutzen Sie Sektionsfilter
3. **Netzwerk-Latenz** → Prüfen Sie Verbindung
4. **Große Mailbox-Datenbanken** → `MaxMailboxesForStats` anpassen
5. **PDF-Export** → Dauert länger, nutzen Sie nur HTML

---

## 📊 Output & Dateien

### Wo werden die Reports gespeichert?

**Standard:** `C:\ExchangeDoku`

```powershell
# Oder angeben:
-OutputPath 'D:\Reports'

# Dateiname automatisch generiert:
Exchange_Dokumentation_20260604_120000.html
Exchange_Dokumentation_20260604_120000.log
```

---

### Welche Dateiformate werden unterstützt?

**3 Formate:**

| Format | Use Case | Größe |
|---|---|---|
| **HTML** | Anzeigen, Word-Import, Web | ~5-10 MB |
| **PDF** | Druck, Versand, Archiv | ~10-20 MB |
| **Markdown** | Wiki, GitHub, Dokumentation | ~1-2 MB |

---

### Kann ich die HTML-Datei in Word bearbeiten?

**Ja!** Einfach öffnen:

```powershell
# Datei in Word öffnen
Start-Process "D:\Reports\Exchange_Dokumentation.html"
# oder manuell: Datei → Öffnen → HTML-Datei wählen
```

**Achtung:** Formatierung kann sich verändern.

---

### Wie groß werden die Report-Dateien?

**Typische Größen:**

- 1 Server, alle Sektionen: **5-10 MB (HTML)**
- 3 Server, alle Sektionen: **15-30 MB (HTML)**
- Mit PDF-Export: **+50%**
- Mit Markdown: **-80%**

**Tipp:** Nutzen Sie Markdown für Archivierung.

---

## 🔍 Inhalte & Daten

### Was genau wird dokumentiert?

**Hauptkategorien:**

✅ Hardware & OS  
✅ Exchange-Konfiguration  
✅ Zertifikate  
✅ Datenbanken & DAG  
✅ Transport & Routing  
✅ Security & Compliance  
✅ Active Directory  

**Keine sensiblen Passwörter!**

---

### Was passiert mit meinen Daten?

**Wichtig:** 

- ✅ **Nur lokale Verarbeitung** – Keine Cloud/Upload
- ✅ **Read-Only** – Skript verändert nichts
- ✅ **Lokal gespeichert** – Im angegebenen `OutputPath`
- ✅ **Sie kontrollieren Zugriff** – Sharing nach Bedarf
- ❌ **Keine Authentifizierung** – Nicht verschlüsselt

---

### Können Mailbox-Inhalte gelesen werden?

**Nein**, das Skript liest:

- ✅ Mailbox-Statistiken (Größe, Item-Count)
- ✅ Mailbox-Eigenschaften (Alias, Archive-Status)
- ❌ **NICHT** E-Mail-Inhalte
- ❌ **NICHT** Passwörter

---

## 🔐 Sicherheit & Berechtigungen

### Welche Rechte werden benötigt?

**Erforderlich:**

- 🔑 Exchange Organization Management (Read-Only OK)
- 🔑 Active Directory Reader
- 🔑 Lokale Administrator-Rechte (optional, aber empfohlen)

```powershell
# Mit Standard-User (keine Admin-Rechte):
# Funktioniert, aber einige Infos fehlen (Registry-Abfragen)

# Mit Admin-Rechten:
# Vollständige Informationen
```

---

### Ist das Skript sicher zu verwenden?

**Ja:**

✅ **Read-Only** – Ändert nichts  
✅ **Open Source** – Code einsehbar  
✅ **Signiert?** – Nein, aus GitHub  
✅ **Zertifikat?** – Nein  

**Empfehlung:** 

```powershell
# Vor Ausführung prüfen:
Get-Content '.\Exchange_Dokumentation.ps1' -Head 100
```

---

### Kann ich das in einer Test-Umgebung testen?

**Ja, absolut!** **Empfohlen:**

```powershell
# 1. In Test-Umgebung testen
# 2. Mit vertrauten Servern beginnen (z.B. 1 Server)
# 3. Schrittweise ausbauen (→ mehr Server)
# 4. Dann in Produktion nutzen
```

---

## 🛠️ Troubleshooting

### Ich bekomme "Exchange-Snap-In nicht gefunden"

**Lösung:**

```powershell
# 1. Prüfen, ob auf Exchange-Server
Get-PSSnapin -Registered

# 2. Falls nicht, manuell laden:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# 3. Falls Fehler, ist Exchange nicht installiert
# → Auf einem Exchange-Server ausführen
```

---

### Das Skript kann sich nicht verbinden

**Checklisten:**

- [ ] Server sind erreichbar (`ping EX01`)
- [ ] Firewall öffnet Port 135 (DCOM) oder 5985 (WinRM)
- [ ] Mit Admin-Rechten ausgeführt?
- [ ] Netzwerk-Konnektivität?

```powershell
# Verbindung testen
Test-Connection -ComputerName 'EX01'
Test-NetConnection -ComputerName 'EX01' -Port 135
```

---

### Ich erhalte "Zugriff verweigert"

**Lösungen:**

```powershell
# 1. Mit Admin-Rechten ausführen
Start-Process powershell -Verb RunAs

# 2. Execution-Policy ändern
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

# 3. Berechtigungen auf Servern prüfen
# → Sie brauchen View-Only Exchange-Rechte
```

---

## 📋 Versioning & Updates

### Wie prüfe ich die Version?

**Im Skript:**

```powershell
# Datei öffnen und suchen nach:
# .VERSION
# oder:
# Version: X.X

# Oder:
Select-String -Path '.\Exchange_Dokumentation.ps1' -Pattern 'Version:'
```

---

### Gibt es neue Versionen?

**Ja!** Prüfen auf GitHub:

```powershell
cd Exchange-Documentation
git fetch origin
git log --oneline -n 5   # Letzte 5 Commits

# Aktualisieren
git pull origin main
```

---

### Kann ich ein Feature-Request einreichen?

**Ja!** Bitte auf GitHub:

🔗 https://github.com/RoccoAmmon/Exchange-Documentation/issues

- Bug-Reports
- Feature-Anfragen
- Verbesserungsvorschläge

---

## 📞 Support

### Wo kann ich Hilfe bekommen?

1. **Dieses Wiki** – Umfassende Dokumentation
2. **Troubleshooting** – [Siehe hier](Troubleshooting)
3. **GitHub Issues** – Probleme melden
4. **Best Practices** – [Siehe hier](Best-Practices)

---

### Kann ich das Skript erweitern?

**Ja!** Es ist Open Source.

- 🔓 Forken Sie das Repository
- 🛠️ Machen Sie Ihre Änderungen
- 🔄 Reichen Sie Pull Request ein

---

## 🎯 Nächste Schritte

- ➡️ [Troubleshooting](Troubleshooting) – Mehr Hilfe
- ➡️ [Best Practices](Best-Practices) – Tipps & Tricks
- ➡️ [Beispiele](Beispiele) – Praktische Szenarien
