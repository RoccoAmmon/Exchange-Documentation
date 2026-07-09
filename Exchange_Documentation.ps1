<#
.SYNOPSIS
    Exchange On-Premises Dokumentations-Skript (v1.7 - Exchange 2013/2016/2019 & SE Support)
.DESCRIPTION
    Erstellt eine umfassende HTML-Dokumentation der gesamten Exchange On-Premises Umgebung.
    Unterstützt Exchange Server 2019 und Exchange Server Subscription Edition (SE).
    Das HTML-Dokument kann direkt in Microsoft Word importiert werden.
    
    WICHTIG: Diese Version verwendet CIM-Sessions mit automatischem DCOM-Fallback,
    sodass das Skript auch funktioniert, wenn WinRM (PowerShell Remoting) nicht
    korrekt konfiguriert ist.

    NEU in v1.7 (Erweiterte System- & Sicherheitschecks):
    - Windows Features & Rollen (installierte Server Rollen)
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
    - Installierte Antivirenlösung (SecurityCenter, Registry, Defender-Status)
    - NIC Receive Buffer Analyse (10/25/40 Gbit/s, Intel/Microsoft-Empfehlung)

    NEU in v1.6 (Erweiterte Dokumentationsbereiche):
    - Message Queue Analyse (Warteschlangen, Status, Nachrichtenanzahl)
    - Calendar & Resource Mailbox Konfiguration (Raum-/Ressourcenpostfächer, Buchungsoptionen)
    - Exchange Archive Konfiguration (Archivpostfächer, Quotas, Auto-Expanding Archive)
    - Exchange Message Size Limits (Org, Connector, Remote Domain, Benutzer)
    - Exchange Partner Applications (SharePoint, Skype, CRM)
    - Exchange Federated Sharing (Federation Trust, Organization Relationships)
    - OAuth / Certificate Based Auth (Auth Server, Zertifikate, CBA)

    NEU in v1.5 (Modernes GUI-Redesign & Bugfixes):
    - Komplett überarbeitete WPF-Oberfläche mit modernem Design
    - Neue Header-Grafik mit Farbverlauf und Version-Badge
    - Moderne Karten (Cards) für alle Bereiche mit Schatteneffekten
    - Kategorisierte Server-Auswahl mit Icons und Farbcodierung
    - Neue Status-Anzeige mit farbigen Hinweisen (Error/Success/Info)
    - Elegante Toggle-Buttons für Ausgabeformate
    - Custom ScrollViewer, CheckBox, TextBox und Button-Styles
    - Verbesserte Fehlerbehandlung und Logging-Mechanismen
    - Automatische Erstellung des Ausgabeverzeichnisses

    NEU in v1.4 (17 Health Checks):
    - Automatischer Neustart mit Administrator-Rechten
    - Auflistung installierter Software pro Server
    - NIC Speed & Performance Checks (Gbps-Warnung)
    - Power Plan Konfiguration (Best Practice: Höchste Leistung)
    - SMBv1 Status & Security Check
    - Processor Core Analyse (Mindestanforderung 4 Kerne)
    - RAM Requirements Check (Exchange-spezifisch)
    - Certificate Expiration Status (Ablauf-Überwachung mit Ampel)
    - Exchange Service Status (Kritische Services)
    - IIS Application Pool Konfiguration
    - DAG Replication Health (nur wenn konfiguriert)
    
    NEU in v1.3:
    - Zuverlässigere Exchange-Edition-Erkennung über Versionsnummer (15.0/15.1/15.2)
    - Zusätzliche Unterstützung für Exchange 2013 (Version 15.0) und Exchange 2016 (Version 15.1)
    - Sauberere Fallback-Logik ohne harten "2019"-Fallback
    - Konsistente Edition-Erkennung in Get-ExchangeEdition und Get-ExchangeServerOverview

    NEU in v1.2:
    - TLS/SSL Konfiguration mit Best Practice Bewertung (Registry, Ciphers, Zertifikate)
    - Windows Service TLS/TLS 1.2 Erzwingung
    - Exchange Connector TLS-Settings Audit
    
    NEU in v1.1:
    - Transportkomponenten - Physische Speicherorte (Queue-DB, Logs, Message-Tracking, SMTP-Protokoll, Safety-Net)
    
    NEU in v1.0:
    - Vollständige Exchange Server SE Unterstützung (Modul-basiert)
    - Automatische Erkennung Exchange 2019 vs. Exchange SE (Build-Nummer + Registry-Check)
    - Erweiterte Sicherheits-Dokumentation (TLS, Auth, SMTP-Relay)
    - Backup/Recovery-Status und High Availability Details
    - Anti-Spam/Anti-Malware Konfiguration
    - Compliance & DLP Richtlinien
    - Quota-Konfiguration und Warnungs-Schwellwerte
    - Exchange Emergency Mitigation Service (EEMS) Monitoring
    - Detaillierte Webservice-Authentifizierung für alle VirtualDirectories

    Dokumentiert werden u.a.:
    - Hardware-Informationen (OS, RAM, CPU, Festplatten, Netzwerk)
    - Installierte Software (Programme, Versionen, Hersteller, Installationsdatum)
    - Netzwerk-Performance (NIC Speed, Duplex-Mode) - mit 1 Gbps Best-Practice-Warnung
    - Power Plan Status (Sollte auf "Höchste Leistung" eingestellt sein)
    - SMBv1 Status (Security Best Practice: deaktiviert)
    - DAG Replication Health (mit Prüfung ob DAG konfiguriert ist)
    - Windows- und Exchange-Patches
    - FSMO-Rollen und AD-Informationen
    - Schema-Version (inkl. Exchange Schema Mapping für 2019 & SE)
    - Exchange Server-Konfiguration und Dienste-Health-Check
    - Virtuelle Verzeichnisse / URLs / Namespace mit Authentication:
      • OWA, ECP, ActiveSync, EWS, MAPI, OAB (Basic/Windows/Forms/IIS-Auth)
      • Outlook Anywhere Konfiguration
      • Autodiscover Service
      • Exchange Emergency Mitigation Service (EEMS)
    - Exchange Emergency Mitigation Service (EEMS) - detaillierte Informationen:
      • Windows Service Status (MSExchangeServiceHost)
      • Pattern Service Erreichbarkeit
      • Applied Mitigations Registry-Werte
      • Telemetry Status
      • IIS Module Anomalies Detection
    - Datenbanken und DAG-Konfiguration (inkl. Lagged Copies, Replay Queue)
    - Public Folder
    - Sende- und Empfangsconnectoren
    - Remote Domains, Accepted Domains
    - MX-Records, SPF, DMARC, DKIM, Autodiscover-DNS
    - E-Mail-Adressrichtlinien
    - Adresslisten und Offline-Adressbuch
    - Mobile Device Policies
    - Zertifikate inkl. Bindungen und Ablauf-Ampel
    - Transport-Regeln
    - Transportkomponenten - Physische Speicherorte (Queue-DB, Logs, Message-Tracking, SMTP-Protokoll, Safety-Net)
    - Mailbox-Statistiken und Empfänger-Übersicht
    - Throttling Policies
    - Retention Policies und Journal Rules
    - RBAC-Rollengruppen und Mitglieder
    - Organisationskonfiguration und Transport-Config
    - Event Logs (Fehler/Kritisch der letzten 7 Tage)
    - Hybrid-Konfiguration mit Exchange Online (Connectors, OAuth, Remote Mailboxes, Migration)
    - Sicherheit & Authentifizierung (TLS, OAuth, SMTP-Auth-Einstellungen)
    - TLS/SSL Konfiguration (Windows Registry, Cipher Suites, Zertifikate mit Best Practice Bewertung)
    - Anti-Spam/Anti-Malware Konfiguration
    - Compliance (DLP, Litigation Hold, In-Place Hold)
    - Mailbox-Quotas und Warnungs-Schwellwerte
    - SMTP-Relay Konfiguration (Geräte/Applikationen)
    - Message Queue Analyse (Warteschlangen, Status, Nachrichtenanzahl)
    - Calendar & Resource Mailbox (Raum-/Ressourcenpostfächer, Buchungsoptionen)
    - Exchange Archive (Archivpostfächer, Quotas, Auto-Expanding Archive)
    - Exchange Message Size Limits (Org, Connector, Remote Domain, Benutzer)
    - Exchange Partner Applications (SharePoint, Skype, CRM)
    - Exchange Federated Sharing (Federation Trust, Organization Relationships)
    - OAuth / Certificate Based Auth (Auth Server, Zertifikate, CBA)

.PARAMETER ExchangeServers
    Array von Exchange-Servernamen, die dokumentiert werden sollen.
    Beispiel: @("EX01","EX02")

.PARAMETER OutputPath
    Pfad für die Ausgabe-Datei. Standard: C:\ExchangeMigrationLog

.PARAMETER CompanyName
    Name der Firma für die Dokumentation.

.EXAMPLE
    .\Exchange_Dokumentation.ps1 -ExchangeServers @("EX01","EX02") -CompanyName "Contoso GmbH"

.EXAMPLE
    .\Exchange_Dokumentation.ps1 -ExchangeServers @("EX01") -CompanyName "Meine Firma" -OutputPath "D:\Doku"

.NOTES
    Autor:           Rocco Ammon
    Version:         1.7
    Erstellt:        2026-03-05
    Letzte Änderung: 2026-07-10
    Änderungen:      v1.7 - Erweiterte System- & Sicherheitschecks:
                          • Windows Features & Rollen (installierte Server Rollen)
                          • .NET Framework Version & DLLs (Release-Key, Assembly-Versionen)
                          • Ausstehende Neustarts (6 Prüfmethoden)
                          • CPU Throttling Analyse (CurrentClockSpeed vs MaxClockSpeed)
                          • Visual C++ Redistributable Versionen (32/64-Bit)
                          • Credential Guard Status (LsaCfgFlags, VBS)
                          • Lokale Administratoren (Mitglieder via CIM/ADSI)
                          • Domain Trusts & Verschlüsselung (Trust-Typen)
                          • FIP-FS Scan Engine Version (Anti-Malware Engine)
                          • Exchange Setting Overrides (alle aktiven Overrides)
                          • Exchange Server Component State (Maintenance Mode)
                          • Security CVE Prüfung (CVE-2021-34470, CVE-2022-21978)
                          • HTTP Proxy Konfiguration (WinHTTP, Registry, netsh)
                          • Installierte Antivirenlösung (SecurityCenter WMI, Registry, Defender)
                          • NIC Receive Buffer Analyse (10/25/40 Gbit/s, Intel/Microsoft-Empfehlung)
                     v1.6 - Erweiterte Dokumentationsbereiche:
                          • Message Queue Analyse (Warteschlangen, Status, Nachrichtenanzahl)
                          • Calendar & Resource Mailbox (Raum-/Ressourcenpostfächer, Buchungsoptionen)
                          • Exchange Archive (Archivpostfächer, Quotas, Auto-Expanding Archive)
                          • Exchange Message Size Limits (Org, Connector, Remote Domain, Benutzer)
                          • Exchange Partner Applications (SharePoint, Skype, CRM)
                          • Exchange Federated Sharing (Federation Trust, Organization Relationships)
                          • OAuth / Certificate Based Auth (Auth Server, Zertifikate, CBA)
                     v1.5 - Modernes GUI-Redesign & Bugfixes:
                          • Komplett überarbeitete WPF-Oberfläche mit modernem Design
                          • Neue Header-Grafik mit Farbverlauf und Version-Badge
                          • Moderne Karten (Cards) für alle Bereiche mit Schatteneffekten
                          • Kategorisierte Server-Auswahl mit Icons und Farbcodierung
                          • Neue Status-Anzeige mit farbigen Hinweisen (Error/Success/Info)
                          • Elegante Toggle-Buttons für Ausgabeformate
                          • Custom ScrollViewer, CheckBox, TextBox und Button-Styles
                          • Verbesserte Fehlerbehandlung und Logging-Mechanismen
                          • Automatische Erstellung des Ausgabeverzeichnisses
                     v1.3 - Verbesserte Exchange-Edition-Erkennung (Unterstützung 2013/2016/2019/SE via Versionsnummer)
                     v1.2 - TLS/SSL Konfiguration mit Best Practice Bewertung
                     v1.1 - Erweiterte Transportkomponenten-Dokumentation (Speicherorte, Queue-DB, Message-Tracking, SMTP-Logs, Safety-Net)
                     v1.0 - Erstveröffentlichung mit Exchange SE Unterstützung, EEMS Monitoring und erweiterten Sicherheits-/Compliance-Dokumentationen
                     v0.1 - Vorversionen / interne Tests (historisch)
    Voraussetzungen: - Exchange 2019 Management Shell/Snap-In ODER Exchange SE PowerShell-Modul
                     - Active Directory PowerShell-Modul
                     - Administratorrechte auf den Exchange-Servern
                     - RPC/DCOM oder WinRM muss erreichbar sein
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Geben Sie die Exchange-Servernamen als Array an, z.B. @('EX01','EX02')")]
    [string[]]$ExchangeServers,

    [Parameter(Mandatory = $false, HelpMessage = "Ausgabepfad für die Dokumentation")]
    [string]$OutputPath = "C:\ExchangeDoku",

    [Parameter(Mandatory = $false, HelpMessage = "Firmenname für die Dokumentation")]
    [string]$CompanyName = "Meine Organisation",

    [Parameter(Mandatory = $false, HelpMessage = "Ausgabeformate: HTML, PDF, Markdown")]
    [ValidateSet("HTML", "PDF", "Markdown")]
    [string[]]$OutputFormats = @("HTML"),

    [Parameter(Mandatory = $false, HelpMessage = "Liste der zu erstellenden Sektions-Schlüssel. Leer = alle")]
    [string[]]$Sections,

    [Parameter(Mandatory = $false, HelpMessage = "GUI zur Auswahl anzeigen (Standard, wenn keine Server angegeben)")]
    [switch]$ShowGui,

    [Parameter(Mandatory = $false, HelpMessage = "GUI unterdrücken und direkt mit Parametern starten")]
    [switch]$NoGui
)

#region ============================================================
# ADMINISTRATOR-PRÜFUNG & AUTO-RESTART
#endregion ============================================================

# Prüfe, ob das Skript als Administrator läuft
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $isAdmin) {
    Write-Host "Das Skript erfordert Administrator-Rechte. Starte neu mit erhöhten Rechten..." -ForegroundColor Yellow
    
    # Sammle alle übergebenen Parameter
    $argumentList = @()
    
    if ($ExchangeServers) {
        $argumentList += "-ExchangeServers @($(($ExchangeServers | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($OutputPath -ne "C:\ExchangeDoku") {
        $argumentList += "-OutputPath '$OutputPath'"
    }
    if ($CompanyName -ne "Meine Organisation") {
        $argumentList += "-CompanyName '$CompanyName'"
    }
    if ($OutputFormats.Count -gt 0) {
        $argumentList += "-OutputFormats @($(($OutputFormats | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($Sections) {
        $argumentList += "-Sections @($(($Sections | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($ShowGui) {
        $argumentList += "-ShowGui"
    }
    if ($NoGui) {
        $argumentList += "-NoGui"
    }
    
    # Starte das Skript mit Administrator-Rechten
    $scriptPath = $MyInvocation.MyCommand.Path
    $command = "& '$scriptPath' $($argumentList -join ' ')"
    
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$command`"" -Verb RunAs -Wait
    exit
}

Write-Host "Administrator-Rechte bestätigt. Skript wird ausgeführt..." -ForegroundColor Green

#region ============================================================
# VARIABLEN-DEFINITION
#endregion ============================================================

# --- Pfade und Dateien ---
$script:LogPath                 = $OutputPath
$script:LogFile                 = Join-Path -Path $LogPath -ChildPath "Exchange-Dokumentation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:HTMLOutputFile          = Join-Path -Path $LogPath -ChildPath "Exchange_Dokumentation_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$script:Timestamp               = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$script:DateOnly                = Get-Date -Format "dd.MM.yyyy"

# --- Dokumentations-Variablen ---
$script:DocTitle                = "Exchange Server - Umgebungsdokumentation"
$script:DocSubTitle             = "$CompanyName"
$script:DocAuthor               = $env:USERNAME
$script:DocComputerName         = $env:COMPUTERNAME

# --- Ausgabeverzeichnis erstellen ---
try {
    if (-not (Test-Path $script:LogPath)) {
        New-Item -Path $script:LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Host "FEHLER: Ausgabeverzeichnis konnte nicht erstellt werden: $_" -ForegroundColor Red
}

# --- Sammelvariablen für HTML-Sektionen ---
$script:HTMLSections            = [System.Collections.ArrayList]::new()
$script:TOCEntries              = [System.Collections.ArrayList]::new()
$script:SectionCounter          = 0
$script:ErrorCount              = 0
$script:WarningCount            = 0

# --- Exchange Snap-In / Modul Name ---
$script:ExchangeSnapIn          = "Microsoft.Exchange.Management.PowerShell.SnapIn"
$script:ExchangeSEModule        = "ExchangeManagementTools"    # Exchange SE verwendet ein PS-Modul
$script:ExchangeEdition         = "Unknown"                    # Wird zur Laufzeit ermittelt (2019 oder SE)

# --- DNS-Server für MX-Abfragen (leer = Standard-DNS) ---
$script:DNSServer               = ""

# --- Grenzwerte für Warnungen ---
$script:WarningDiskSpaceGB      = 20      # Warnung bei weniger als 20 GB freiem Speicher
$script:WarningCertDaysExpiry   = 30      # Warnung bei Zertifikaten, die in 30 Tagen ablaufen
$script:MaxMailboxesForStats    = 500     # Maximale Anzahl Mailboxen für detaillierte Statistiken

#endregion

#region ============================================================
# HILFSFUNKTIONEN
#endregion ============================================================

function Write-Log {
    <#
    .SYNOPSIS
        Schreibt eine Nachricht in die Log-Datei und auf die Konsole.
    .PARAMETER Message
        Die zu protokollierende Nachricht.
    .PARAMETER Level
        Log-Level: INFO, WARNING, ERROR
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    try {
        $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$logTimestamp] [$Level] $Message"

        # Sicherstellen, dass das Verzeichnis existiert
        $logDir = Split-Path $script:LogFile -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        # In Datei schreiben
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8

        # Konsolenausgabe mit Farbe
        switch ($Level) {
            "INFO"    { Write-Host $logEntry -ForegroundColor Green }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow; $script:WarningCount++ }
            "ERROR"   { Write-Host $logEntry -ForegroundColor Red; $script:ErrorCount++ }
        }
    }
    catch {
        Write-Host "FEHLER beim Schreiben der Log-Datei: $_" -ForegroundColor Red
    }
}

function New-HTMLSection {
    <#
    .SYNOPSIS
        Erstellt eine neue HTML-Sektion mit Überschrift und fügt sie dem Inhaltsverzeichnis hinzu.
    .PARAMETER Title
        Titel der Sektion.
    .PARAMETER Content
        HTML-Inhalt der Sektion.
    .PARAMETER Level
        Überschriften-Level (1-3). Standard: 2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $false)]
        [int]$Level = 2
    )

    $script:SectionCounter++
    $anchorId = "section_$($script:SectionCounter)"

    # Inhaltsverzeichnis-Eintrag
    [void]$script:TOCEntries.Add(@{
        Title  = $Title
        Anchor = $anchorId
        Level  = $Level
    })

    # HTML-Sektion erstellen
    $sectionHTML = @"
    <div class="section">
        <h$Level id="$anchorId">$($script:SectionCounter). $Title</h$Level>
        $Content
    </div>
"@

    [void]$script:HTMLSections.Add($sectionHTML)
}

function ConvertTo-HTMLTable {
    <#
    .SYNOPSIS
        Konvertiert ein Array von Objekten in eine formatierte HTML-Tabelle.
    .PARAMETER Data
        Die zu konvertierenden Daten.
    .PARAMETER Properties
        Optionale Eigenschaftsauswahl.
    .PARAMETER NoDataMessage
        Nachricht, wenn keine Daten vorhanden sind.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$Data,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties,

        [Parameter(Mandatory = $false)]
        [string]$NoDataMessage = "Keine Daten verfügbar."
    )

    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>$NoDataMessage</p>"
    }

    try {
        if ($Properties) {
            $Data = $Data | Select-Object -Property $Properties
        }

        $html = "<table>`n<thead>`n<tr>"

        # Header ermitteln
        $headers = $Data[0].PSObject.Properties.Name
        foreach ($header in $headers) {
            $html += "<th>$header</th>"
        }
        $html += "</tr>`n</thead>`n<tbody>`n"

        # Datenzeilen
        $rowIndex = 0
        foreach ($row in $Data) {
            $rowClass = if ($rowIndex % 2 -eq 0) { "even" } else { "odd" }
            $html += "<tr class='$rowClass'>"
            foreach ($header in $headers) {
                $value = $row.$header
                if ($null -eq $value) { $value = "-" }
                $html += "<td>$([System.Web.HttpUtility]::HtmlEncode($value.ToString()))</td>"
            }
            $html += "</tr>`n"
            $rowIndex++
        }

        $html += "</tbody>`n</table>"
        return $html
    }
    catch {
        Write-Log -Message "Fehler bei HTML-Tabellen-Konvertierung: $_" -Level "ERROR"
        return "<p class='error'>Fehler bei der Datenkonvertierung: $_</p>"
    }
}

function New-ServerCimSession {
    <#
    .SYNOPSIS
        Erstellt eine CIM-Session zu einem Server mit automatischem WsMan→DCOM Fallback.
    .PARAMETER ComputerName
        Name des Zielservers.
    .OUTPUTS
        CimSession-Objekt oder $null bei Fehler.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    # Methode 1: WsMan (Standard, nutzt WinRM)
    try {
        Write-Log -Message "CIM-Session über WsMan für $ComputerName..." -Level "INFO"
        $session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
        Write-Log -Message "CIM-Session (WsMan) zu $ComputerName erfolgreich." -Level "INFO"
        return $session
    }
    catch {
        Write-Log -Message "WsMan fehlgeschlagen für ${ComputerName}: $($_.Exception.Message)" -Level "WARNING"
    }

    # Methode 2: DCOM Fallback (kein WinRM nötig, nutzt RPC)
    try {
        Write-Log -Message "CIM-Session über DCOM für $ComputerName..." -Level "INFO"
        $dcomOption = New-CimSessionOption -Protocol Dcom
        $session = New-CimSession -ComputerName $ComputerName -SessionOption $dcomOption -ErrorAction Stop
        Write-Log -Message "CIM-Session (DCOM) zu $ComputerName erfolgreich." -Level "INFO"
        return $session
    }
    catch {
        Write-Log -Message "Auch DCOM fehlgeschlagen für ${ComputerName}: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Get-RemoteRegistryValue {
    <#
    .SYNOPSIS
        Liest Registry-Werte remote über .NET-Methoden (kein WinRM nötig).
    .PARAMETER ComputerName
        Name des Zielcomputers.
    .PARAMETER RegistryPath
        Pfad innerhalb der Registry (HKLM).
    .PARAMETER ValueName
        Name des zu lesenden Wertes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$RegistryPath,

        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName
        )
        $subKey = $regKey.OpenSubKey($RegistryPath)
        if ($subKey) {
            $value = $subKey.GetValue($ValueName)
            $subKey.Close()
            $regKey.Close()
            return $value
        }
        $regKey.Close()
        return $null
    }
    catch {
        Write-Log -Message "Remote Registry auf ${ComputerName} fehlgeschlagen ($ValueName): $_" -Level "WARNING"
        return $null
    }
}

#endregion

#region ============================================================
# EXCHANGE SNAP-IN LADEN
#endregion ============================================================

function Initialize-ExchangeEnvironment {
    <#
    .SYNOPSIS
        Lädt das Exchange PowerShell Snap-In für On-Premises.
        Die Edition (2019/SE) wird NACH dem Laden via Get-ExchangeServer ermittelt.
    #>
    Write-Log -Message "=== Exchange Management Tools werden geladen ===" -Level "INFO"

    try {
        # --- Methode 1: Snap-In bereits geladen? ---
        if (Get-PSSnapin -Name $script:ExchangeSnapIn -ErrorAction SilentlyContinue) {
            Write-Log -Message "Exchange Snap-In ist bereits geladen." -Level "INFO"
            return $true
        }

        # --- Methode 2: Snap-In registriert? ---
        if (Get-PSSnapin -Registered -Name $script:ExchangeSnapIn -ErrorAction SilentlyContinue) {
            Add-PSSnapin -Name $script:ExchangeSnapIn -ErrorAction Stop
            Write-Log -Message "Exchange Snap-In erfolgreich geladen." -Level "INFO"
            return $true
        }

        # --- Methode 3: Exchange Management Shell via RemoteExchange.ps1 ---
        Write-Log -Message "Snap-In nicht direkt verfügbar. Versuche Exchange Management Shell zu laden..." -Level "WARNING"

        $exchangeInstallPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ErrorAction SilentlyContinue).MsiInstallPath
        if ($exchangeInstallPath) {
            $remoteExchangeScript = Join-Path -Path $exchangeInstallPath -ChildPath "bin\RemoteExchange.ps1"
            if (Test-Path $remoteExchangeScript) {
                . $remoteExchangeScript
                Connect-ExchangeServer -auto -AllowClobber
                Write-Log -Message "Exchange Management Shell über RemoteExchange.ps1 geladen." -Level "INFO"
                return $true
            }
        }

        Write-Log -Message "Exchange Management Tools nicht gefunden. Bitte auf einem Exchange-Server ausführen." -Level "ERROR"
        return $false
    }
    catch {
        Write-Log -Message "Fehler beim Laden der Exchange Management Tools: $_" -Level "ERROR"
        return $false
    }
}

function Get-ExchangeEdition {
    <#
    .SYNOPSIS
        Ermittelt die Exchange Edition (2013, 2016, 2019 oder SE) anhand der Version und Build.
        Version-erste Erkennung (zuverlässiger):
        - Version 15.0 → 2013
        - Version 15.1 → 2016
        - Version 15.2 mit Build >= 2000 → SE
        - Version 15.2 mit Build < 2000 → 2019
        Fallback auf Build-Ranges wenn Version nicht zu extrahieren ist.
    #>
    Write-Log -Message "=== Ermittle Exchange Edition ===" -Level "INFO"

    try {
        $exServer = Get-ExchangeServer -ErrorAction Stop | Select-Object -First 1
        if ($exServer) {
            $adminVersion = $exServer.AdminDisplayVersion.ToString()
            Write-Log -Message "AdminDisplayVersion: $adminVersion" -Level "INFO"

            # Versionsstring prüfen (zuverlässiger als Build-Ranges)
            if ($adminVersion -match 'Version\s+(\d+\.\d+)') {
                $version = $Matches[1]
                Write-Log -Message "Exchange Version erkannt: $version" -Level "INFO"

                if ($version -eq "15.0") {
                    $script:ExchangeEdition = "2013"
                    Write-Log -Message "Exchange Edition erkannt: 2013 (Version 15.0)" -Level "INFO"
                }
                elseif ($version -eq "15.1") {
                    $script:ExchangeEdition = "2016"
                    Write-Log -Message "Exchange Edition erkannt: 2016 (Version 15.1)" -Level "INFO"
                }
                elseif ($version -eq "15.2") {
                    # Version 15.2 könnte 2019 oder SE sein - Build prüfen
                    if ($adminVersion -match 'Build\s+(\d+)') {
                        $build = [int]$Matches[1]
                        if ($build -ge 2000) {
                            $script:ExchangeEdition = "SE"
                            Write-Log -Message "Exchange Edition erkannt: SE (Version 15.2, Build $build >= 2000)" -Level "INFO"
                        }
                        else {
                            $script:ExchangeEdition = "2019"
                            Write-Log -Message "Exchange Edition erkannt: 2019 (Version 15.2, Build $build < 2000)" -Level "INFO"
                        }
                    }
                    else {
                        # Build nicht zu extrahieren, Fallback auf 2019
                        $script:ExchangeEdition = "2019"
                        Write-Log -Message "Exchange Edition erkannt: 2019 (Version 15.2, Build nicht erkannt)" -Level "INFO"
                    }
                }
                else {
                    # Unbekannte Version
                    $script:ExchangeEdition = "Unbekannt"
                    Write-Log -Message "Exchange Edition unbekannt - Version $version nicht erkannt" -Level "WARNING"
                }
            }
            else {
                Write-Log -Message "Version konnte nicht extrahiert werden - Fallback auf Build-Ranges" -Level "WARNING"
                $script:ExchangeEdition = "Unbekannt"
            }
        }
        else {
            Write-Log -Message "Kein Exchange Server gefunden" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Fehler bei Edition-Erkennung: $_" -Level "WARNING"
        $script:ExchangeEdition = "Unbekannt"
    }
}

#endregion

#region ============================================================
# DATENSAMMLUNGS-FUNKTIONEN
#endregion ============================================================

# ---------------------------------------------------------------
# 1. HARDWARE-INFORMATIONEN (CIM/DCOM FALLBACK)
# ---------------------------------------------------------------
function Get-HardwareInformation {
    <#
    .SYNOPSIS
        Sammelt Hardware-Informationen über CIM-Sessions mit automatischem
        WsMan→DCOM Fallback. Funktioniert auch ohne WinRM.
    #>
    Write-Log -Message "=== Sammle Hardware-Informationen (CIM/DCOM Fallback) ===" -Level "INFO"

    $allHardwareHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Hardware-Info für Server: $server" -Level "INFO"

            # --- CIM-Session aufbauen (WsMan → DCOM Fallback) ---
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $allHardwareHTML += "<h3 class='server-break'>Server: $server</h3>"
                $allHardwareHTML += "<p class='error'>Keine Verbindung möglich (weder WsMan noch DCOM). Bitte Netzwerk/Firewall prüfen.</p>"
                continue
            }

            # --- Betriebssystem ---
            $osInfo = $null
            try {
                $osRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction Stop
                $osInfo = [PSCustomObject]@{
                    Computername    = $osRaw.CSName
                    Betriebssystem  = $osRaw.Caption
                    Version         = $osRaw.Version
                    BuildNumber     = $osRaw.BuildNumber
                    Architektur     = $osRaw.OSArchitecture
                    InstallDatum    = $osRaw.InstallDate.ToString("dd.MM.yyyy")
                    LetzterBoot     = $osRaw.LastBootUpTime.ToString("dd.MM.yyyy HH:mm")
                    "RAM_Gesamt_GB" = [math]::Round($osRaw.TotalVisibleMemorySize / 1MB, 2)
                    "RAM_Frei_GB"   = [math]::Round($osRaw.FreePhysicalMemory / 1MB, 2)
                    "RAM_Belegt_%"  = [math]::Round((1 - ($osRaw.FreePhysicalMemory / $osRaw.TotalVisibleMemorySize)) * 100, 1)
                }
            }
            catch {
                Write-Log -Message "OS-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- CPU ---
            $cpuInfo = $null
            try {
                $cpuRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor -ErrorAction Stop
                $cpuInfo = foreach ($cpu in $cpuRaw) {
                    [PSCustomObject]@{
                        Prozessor        = $cpu.Name
                        Kerne            = $cpu.NumberOfCores
                        "Logische Proz." = $cpu.NumberOfLogicalProcessors
                        "Max Takt (MHz)" = $cpu.MaxClockSpeed
                        Sockel           = $cpu.SocketDesignation
                        Hyperthreading   = if ($cpu.NumberOfLogicalProcessors -gt $cpu.NumberOfCores) { "Ja" } else { "Nein" }
                    }
                }
            }
            catch {
                Write-Log -Message "CPU-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Logische Laufwerke ---
            $diskInfo = $null
            try {
                $diskRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
                $diskInfo = foreach ($disk in $diskRaw) {
                    $belegtProzent = if ($disk.Size -gt 0) { [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 1) } else { 0 }
                    $freiGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                    [PSCustomObject]@{
                        Laufwerk    = $disk.DeviceID
                        Volumenname = $disk.VolumeName
                        "Gesamt_GB" = [math]::Round($disk.Size / 1GB, 2)
                        "Frei_GB"   = $freiGB
                        "Belegt_%"  = $belegtProzent
                        Dateisystem = $disk.FileSystem
                        Status      = if ($freiGB -lt $script:WarningDiskSpaceGB) { "⚠️ WENIG PLATZ!" } else { "✅ OK" }
                    }
                }
            }
            catch {
                Write-Log -Message "Festplatten-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Physische Festplatten ---
            $physDiskInfo = $null
            try {
                $physDiskRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_DiskDrive -ErrorAction Stop
                $physDiskInfo = foreach ($pd in $physDiskRaw) {
                    [PSCustomObject]@{
                        Modell       = $pd.Model
                        "Größe_GB"   = [math]::Round($pd.Size / 1GB, 2)
                        InterfaceTyp = $pd.InterfaceType
                        MediaType    = $pd.MediaType
                        Partitionen  = $pd.Partitions
                        Status       = $pd.Status
                    }
                }
            }
            catch {
                Write-Log -Message "Phys. Festplatten-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Pagefile ---
            $pageFileInfo = $null
            try {
                $pfRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_PageFileSetting -ErrorAction SilentlyContinue
                if ($pfRaw) {
                    $pageFileInfo = foreach ($pf in $pfRaw) {
                        [PSCustomObject]@{
                            Pfad           = $pf.Name
                            "InitGröße_MB" = $pf.InitialSize
                            "MaxGröße_MB"  = $pf.MaximumSize
                        }
                    }
                }
                else {
                    $pageFileInfo = [PSCustomObject]@{
                        Pfad           = "Automatisch verwaltet"
                        "InitGröße_MB" = "Auto"
                        "MaxGröße_MB"  = "Auto"
                    }
                }
            }
            catch {
                Write-Log -Message "Pagefile-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Netzwerk ---
            $nicInfo = $null
            try {
                $nicRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop
                $nicInfo = foreach ($nic in $nicRaw) {
                    [PSCustomObject]@{
                        Adapter    = $nic.Description
                        IPAdresse  = ($nic.IPAddress -join ', ')
                        Subnetz    = ($nic.IPSubnet -join ', ')
                        Gateway    = ($nic.DefaultIPGateway -join ', ')
                        DNS_Server = ($nic.DNSServerSearchOrder -join ', ')
                        DHCP       = if ($nic.DHCPEnabled) { "Ja" } else { "Nein" }
                        MAC        = $nic.MACAddress
                    }
                }
            }
            catch {
                Write-Log -Message "Netzwerk-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- System-Übersicht / Virtualisierungs-Check ---
            $vmInfo = $null
            try {
                $csRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_ComputerSystem -ErrorAction Stop
                $vmInfo = [PSCustomObject]@{
                    Hersteller      = $csRaw.Manufacturer
                    Modell          = $csRaw.Model
                    Domäne          = $csRaw.Domain
                    Virtuell        = if ($csRaw.Model -match "Virtual|VMware|KVM|Xen|HVM|VRTUAL") { "Ja" } else { "Nein" }
                    "RAM_Gesamt_GB" = [math]::Round($csRaw.TotalPhysicalMemory / 1GB, 2)
                }
            }
            catch {
                Write-Log -Message "System-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Windows Hotfixes über CIM ---
            $hotfixInfo = $null
            try {
                $hotfixRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_QuickFixEngineering -ErrorAction Stop |
                    Sort-Object InstalledOn -Descending | Select-Object -First 30
                $hotfixInfo = foreach ($hf in $hotfixRaw) {
                    [PSCustomObject]@{
                        HotfixID       = $hf.HotFixID
                        Beschreibung   = $hf.Description
                        InstalliertAm  = if ($hf.InstalledOn) { $hf.InstalledOn.ToString("dd.MM.yyyy") } else { "Unbekannt" }
                        InstalliertVon = $hf.InstalledBy
                    }
                }
            }
            catch {
                Write-Log -Message "Hotfix-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Exchange Version über Remote Registry (kein WinRM nötig!) ---
            $exVersionInfo = $null
            try {
                $regPath = "SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
                $exBuildMajor  = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "MsiProductMajor"
                $exBuildMinor  = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "MsiProductMinor"
                $exBuildNum    = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "MsiBuildMajor"
                $exBuildRev    = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "MsiBuildMinor"
                $exInstallPath = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "MsiInstallPath"

                if ($exBuildMajor) {
                    $exVersionInfo = [PSCustomObject]@{
                        "Exchange Build"    = "$exBuildMajor.$exBuildMinor.$exBuildNum.$exBuildRev"
                        "Installationspfad" = $exInstallPath
                    }
                }
            }
            catch {
                Write-Log -Message "Exchange Registry-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Exchange Dienste über CIM ---
            $exchangeServices = $null
            try {
                $servicesRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -Filter "Name LIKE 'MSExchange%'" -ErrorAction Stop |
                    Sort-Object Name
                $exchangeServices = foreach ($svc in $servicesRaw) {
                    [PSCustomObject]@{
                        Dienstname  = $svc.Name
                        Anzeigename = $svc.DisplayName
                        Status      = $svc.State
                        Starttyp    = $svc.StartMode
                        Warnung     = if ($svc.State -ne "Running" -and $svc.StartMode -eq "Auto") { "⚠️ GESTOPPT!" } else { "✅ OK" }
                    }
                }
            }
            catch {
                Write-Log -Message "Dienste-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- CIM-Session aufräumen ---
            if ($cimSession) {
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }

            # --- HTML zusammenbauen ---
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $serverHTML += "<h4>System-Übersicht</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($vmInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Betriebssystem</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($osInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Prozessor(en)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $cpuInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Logische Laufwerke</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $diskInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Physische Festplatten</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $physDiskInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Pagefile</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($pageFileInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Netzwerkkonfiguration</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $nicInfo -NoDataMessage "Nicht verfügbar")

            if ($exVersionInfo) {
                $serverHTML += "<h4>Exchange Version (Registry)</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data @($exVersionInfo))
            }

            $serverHTML += "<h4>Windows Hotfixes (letzte 30)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $hotfixInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Exchange Dienste</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $exchangeServices -NoDataMessage "Nicht verfügbar")

            $allHardwareHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Allgemeiner Fehler bei Hardware-Info für ${server}: $_" -Level "ERROR"
            $allHardwareHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Hardware-Informationen & Server-Details" -Content $allHardwareHTML
}

# ---------------------------------------------------------------
# 1.5 INSTALLIERTE SOFTWARE
# ---------------------------------------------------------------
function Get-InstalledSoftwareInfo {
    <#
    .SYNOPSIS
        Sammelt alle installierten Software-Programme pro Server über die Registry.
        Unterstützt 32-Bit und 64-Bit Programme.
    #>
    Write-Log -Message "=== Sammle installierte Software ===" -Level "INFO"

    $allSoftwareHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Installierte Software für Server: $server" -Level "INFO"
            
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $softwareList = @()

            # Registry-Pfade für 64-Bit und 32-Bit Software
            $regPaths = @(
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )

            try {
                $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                    [Microsoft.Win32.RegistryHive]::LocalMachine, $server
                )

                foreach ($regPath in $regPaths) {
                    try {
                        $subKey = $regKey.OpenSubKey($regPath)
                        if ($subKey) {
                            foreach ($appGuid in $subKey.GetSubKeyNames()) {
                                try {
                                    $appKey = $subKey.OpenSubKey($appGuid)
                                    if ($appKey) {
                                        $displayName = $appKey.GetValue("DisplayName")
                                        $displayVersion = $appKey.GetValue("DisplayVersion")
                                        $installDate = $appKey.GetValue("InstallDate")
                                        $publisher = $appKey.GetValue("Publisher")

                                        # Nur Programme mit DisplayName auflisten (echte Software, nicht Registry-Reste)
                                        if ($displayName) {
                                            # InstallDate formatieren wenn vorhanden
                                            $formattedDate = if ($installDate -and $installDate -match "^\d{8}$") {
                                                [datetime]::ParseExact($installDate, "yyyyMMdd", $null).ToString("dd.MM.yyyy")
                                            }
                                            elseif ($installDate) {
                                                $installDate
                                            }
                                            else {
                                                "Unbekannt"
                                            }

                                            $softwareList += [PSCustomObject]@{
                                                "Softwarename"     = $displayName
                                                "Version"          = if ($displayVersion) { $displayVersion } else { "-" }
                                                "Hersteller"       = if ($publisher) { $publisher } else { "-" }
                                                "Installationsdatum" = $formattedDate
                                            }
                                        }
                                        $appKey.Close()
                                    }
                                }
                                catch {
                                    # Fehlerhafte Einträge überspringen
                                }
                            }
                            $subKey.Close()
                        }
                    }
                    catch {
                        Write-Log -Message "Registry-Pfad nicht lesbar (${regPath}): $_" -Level "WARNING"
                    }
                }

                $regKey.Close()
            }
            catch {
                Write-Log -Message "Remote Registry auf ${server} nicht erreichbar: $_" -Level "ERROR"
                $serverHTML += "<p class='error'>Keine Verbindung zur Remote Registry möglich.</p>"
                $allSoftwareHTML += $serverHTML
                continue
            }

            # Software sortieren und deduplizieren
            $softwareList = $softwareList | 
                Sort-Object -Property "Softwarename" -Unique |
                Select-Object -Property "Softwarename", "Version", "Hersteller", "Installationsdatum"

            if ($softwareList.Count -gt 0) {
                $serverHTML += "<p><strong>Gesamt: $($softwareList.Count) Software-Programme installiert</strong></p>"
                $serverHTML += (ConvertTo-HTMLTable -Data $softwareList -NoDataMessage "Keine Software gefunden")
            }
            else {
                $serverHTML += "<p class='no-data'>Keine Software gefunden.</p>"
            }

            $allSoftwareHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Software-Abfrage für ${server}: $_" -Level "ERROR"
            $allSoftwareHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler bei der Abfrage: $_</p>"
        }
    }

    New-HTMLSection -Title "Installierte Software" -Content $allSoftwareHTML
}

# ---------------------------------------------------------------
# 1.6 NETZWERK-KONFIGURATION (NIC SPEED & PERFORMANCE)
# ---------------------------------------------------------------
function Get-NetworkConfigurationInfo {
    <#
    .SYNOPSIS
        Prüft Netzwerk-Geschwindigkeit und Performance-Einstellungen via Get-NetAdapter.
    #>
    Write-Log -Message "=== Sammle Netzwerk-Konfiguration (NIC Speed) ===" -Level "INFO"

    $allNetworkHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Netzwerk-Config für Server: $server (lokaler Hostname: $env:COMPUTERNAME)" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $nicList = @()
            $adapters = @()

            try {
                # Prüfe ob es der lokale Server ist (case-insensitive)
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                Write-Log -Message "isLocal=$isLocal für Server=$server" -Level "INFO"
                
                if ($isLocal) {
                    # Lokal: Get-NetAdapter direkt
                    Write-Log -Message "Versuche Get-NetAdapter lokal..." -Level "INFO"
                    $adapters = @(Get-NetAdapter -ErrorAction Stop)
                    Write-Log -Message "Get-NetAdapter gefunden: $($adapters.Count) Adapter" -Level "INFO"
                } else {
                    # Remote: Invoke-Command mit Get-NetAdapter
                    Write-Log -Message "Versuche Get-NetAdapter via Remoting auf $server..." -Level "INFO"
                    $adapters = @(Invoke-Command -ComputerName $server -ScriptBlock {
                        Get-NetAdapter -ErrorAction Stop
                    } -ErrorAction Stop)
                    Write-Log -Message "Remoting Get-NetAdapter gefunden: $($adapters.Count) Adapter" -Level "INFO"
                }
                
                # Verarbeite Adapter
                if ($adapters -and $adapters.Count -gt 0) {
                    Write-Log -Message "Verarbeite $($adapters.Count) Adapter..." -Level "INFO"
                    
                    $nicList = @()
                    foreach ($nic in $adapters) {
                        try {
                            # Extrahiere Speed aus LinkSpeed
                            $speedMatch = $null
                            if ($nic.LinkSpeed -match '(\d+)\s*Gbps') {
                                $speedMatch = [double]$Matches[1]
                            }
                            
                            $speedStatus = if ($null -eq $speedMatch) { 
                                "⚠️ Speed unbekannt" 
                            }
                            elseif ($speedMatch -ge 1) { 
                                "✅ OK (>= 1 Gbps)" 
                            } else { 
                                "⚠️ WARNUNG (< 1 Gbps!)" 
                            }
                            
                            $nicList += [PSCustomObject]@{
                                "Adapter Name"      = $nic.Name
                                "Beschreibung"      = $nic.InterfaceDescription
                                "Status"            = $nic.Status
                                "Link Speed"        = $nic.LinkSpeed
                                "MAC-Adresse"       = $nic.MacAddress
                                "Speed-Status"      = $speedStatus
                            }
                        }
                        catch {
                            Write-Log -Message "Fehler bei Adapter-Verarbeitung: $_" -Level "WARNING"
                        }
                    }
                    Write-Log -Message "Erfolgreich $($nicList.Count) Adapter verarbeitet" -Level "INFO"
                }
                else {
                    Write-Log -Message "Keine Adapter gefunden" -Level "WARNING"
                }
            }
            catch {
                Write-Log -Message "NIC-Abfrage für $server fehlgeschlagen: $_" -Level "ERROR"
            }

            # Ausgabe
            if ($nicList -and @($nicList).Count -gt 0) {
                $serverHTML += "<h4>Netzwerkadapter</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data $nicList)
            }
            else {
                $serverHTML += "<p class='no-data'>Keine Netzwerkadapter gefunden.</p>"
            }

            $allNetworkHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Netzwerk-Info für ${server}: $_" -Level "ERROR"
            $allNetworkHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Netzwerk-Konfiguration (NIC Speed & Performance)" -Content $allNetworkHTML
}

# ---------------------------------------------------------------
# 1.7 POWER PLAN & PERFORMANCE
# ---------------------------------------------------------------
function Get-PowerPlanInfo {
    <#
    .SYNOPSIS
        Prüft den aktuellen Power Plan. Sollte auf "Hohe Leistung" eingestellt sein für Produktions-Exchange-Server.
    #>
    Write-Log -Message "=== Sammle Power Plan Informationen ===" -Level "INFO"

    $allPowerHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Power Plan für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            # Registry: Aktuell aktiver Power Plan
            $activePlanGuid = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes" -ValueName "ActivePowerScheme"
            $activePlanName = "Unbekannt"
            $powerStatus = "⚠️ WARNUNG"
            
            if ($activePlanGuid) {
                # Bekannte GUIDs
                $planMap = @{
                    "381b4222-f694-41f0-9685-ff5bb260df2e" = "Höchste Leistung"
                    "1537a863-7a1c-4bea-b8a0-5e755d1e55f0" = "Ausgewogen"
                    "a1841308-3541-4fab-bc81-f71556f20b4a" = "Energiesparen"
                }
                
                # Versuche den Namen aus Registry zu lesen
                try {
                    $planPath = "SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\$activePlanGuid"
                    $activePlanName = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $planPath -ValueName "FriendlyName"
                    if (-not $activePlanName) {
                        $activePlanName = $planMap[$activePlanGuid] -replace "^({.*})", "" 
                    }
                    
                    # Bewertung
                    if ($activePlanName -match "Höchste Leistung|High Performance") {
                        $powerStatus = "✅ OK - Höchste Leistung"
                    }
                    elseif ($activePlanName -match "Ausgewogen|Balanced") {
                        $powerStatus = "⚠️ WARNUNG - Sollte 'Höchste Leistung' sein"
                    }
                }
                catch {
                    Write-Log -Message "Power Plan Name nicht ermittelt: $_" -Level "WARNING"
                }
            }

            $powerInfo = [PSCustomObject]@{
                "Aktiver Power Plan"    = $activePlanName
                "GUID"                  = $activePlanGuid
                "Empfohlener Status"    = "Höchste Leistung"
                "Bewertung"             = $powerStatus
            }

            $serverHTML += "<h4>Aktuelle Power Plan Einstellung</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($powerInfo))
            $serverHTML += "<p><strong>Hinweis:</strong> Für Produktions-Exchange-Server sollte der Power Plan auf 'Höchste Leistung' oder 'High Performance' eingestellt sein.</p>"

            $allPowerHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Power Plan für ${server}: $_" -Level "ERROR"
            $allPowerHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Power Plan & Performance-Einstellungen" -Content $allPowerHTML
}

# ---------------------------------------------------------------
# 1.8 SMBv1 STATUS (SICHERHEIT)
# ---------------------------------------------------------------
function Get-SMBv1StatusInfo {
    <#
    .SYNOPSIS
        Prüft ob SMBv1 deaktiviert ist (Sicherheits- und Performance-Best-Practice).
    #>
    Write-Log -Message "=== Sammle SMBv1 Status ===" -Level "INFO"

    $allSMBHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "SMBv1 Status für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            # Prüfe ob es der lokale Server ist
            $isLocal = ($server -eq $env:COMPUTERNAME) -or ($server -eq 'localhost') -or ($server -eq '.')
            $smbStatus = $null
            
            try {
                if ($isLocal) {
                    # Direkt lokal prüfen mit Get-SmbServerConfiguration
                    try {
                        $smbConfig = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
                        if ($null -ne $smbConfig) {
                            $smbStatus = @{
                                EnableSMB1 = $smbConfig.EnableSMB1Protocol
                                IsDisabled = -not $smbConfig.EnableSMB1Protocol
                                Source = "Get-SmbServerConfiguration"
                            }
                        }
                    } catch {}
                    
                    # Fallback Registry
                    if ($null -eq $smbStatus) {
                        try {
                            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
                            $smb1Reg = Get-ItemProperty -Path $regPath -Name "SMB1" -ErrorAction SilentlyContinue
                            if ($null -ne $smb1Reg) {
                                $smbStatus = @{
                                    SMB1Value = $smb1Reg.SMB1
                                    IsDisabled = ($smb1Reg.SMB1 -eq 0)
                                    Source = "Registry"
                                }
                            }
                        } catch {}
                    }
                } else {
                    # Remoting für entfernte Server
                    $smbStatus = Invoke-Command -ComputerName $server -ScriptBlock {
                        # Methode 1: Get-SmbServerConfiguration
                        try {
                            $smbConfig = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
                            if ($null -ne $smbConfig) {
                                return @{
                                    EnableSMB1 = $smbConfig.EnableSMB1Protocol
                                    IsDisabled = -not $smbConfig.EnableSMB1Protocol
                                    Source = "Get-SmbServerConfiguration"
                                }
                            }
                        } catch {}
                        
                        # Methode 2: Registry Fallback
                        try {
                            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
                            $smb1Reg = Get-ItemProperty -Path $regPath -Name "SMB1" -ErrorAction SilentlyContinue
                            if ($null -ne $smb1Reg) {
                                return @{
                                    SMB1Value = $smb1Reg.SMB1
                                    IsDisabled = ($smb1Reg.SMB1 -eq 0)
                                    Source = "Registry"
                                }
                            }
                        } catch {}
                        
                        return $null
                    } -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Log -Message "SMBv1-Abfrage für $server fehlgeschlagen: $_" -Level "WARNING"
            }

            # Bewertung
            if ($smbStatus) {
                $isDisabled = $smbStatus.IsDisabled
                
                if ($isDisabled -eq $true) {
                    $smbDisplay = "✅ DEAKTIVIERT (Sicher)"
                    $statusColor = "green"
                }
                elseif ($isDisabled -eq $false) {
                    $smbDisplay = "🔴 AKTIVIERT (Unsicher!)"
                    $statusColor = "red"
                }
                else {
                    $smbDisplay = "⚠️ Status konnte nicht eindeutig ermittelt werden"
                    $statusColor = "yellow"
                }
                
                $smbInfo = [PSCustomObject]@{
                    "SMBv1 Status"          = $smbDisplay
                    "Quelle"                = $smbStatus.Source
                    "Sicherheitsbewertung"  = if ($isDisabled -eq $true) { "✅ SEHR GUT" } elseif ($isDisabled -eq $false) { "🔴 KRITISCH" } else { "⚠️ Unbekannt" }
                }
            }
            else {
                # Kein SMBv1 gefunden - wahrscheinlich moderne Windows Version
                $smbInfo = [PSCustomObject]@{
                    "SMBv1 Status"          = "ℹ️ Nicht vorhanden oder Moderne Windows-Version"
                    "Quelle"                = "Abfrage fehlgeschlagen"
                    "Sicherheitsbewertung"  = "✅ SEHR GUT (wahrscheinlich nicht aktiviert)"
                }
            }

            $serverHTML += "<h4>SMBv1 Konfiguration</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($smbInfo))
            $serverHTML += "<p><strong>Best Practice:</strong> SMBv1 sollte deaktiviert sein. Dies bietet Schutz gegen WannaCry und verbessert die Performance. In modernen Windows-Versionen ist SMBv1 standardmäßig deaktiviert.</p>"

            $allSMBHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei SMBv1 Status für ${server}: $_" -Level "ERROR"
            $allSMBHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>SMBv1-Status konnte nicht ermittelt werden: $_</p>"
        }
    }

    New-HTMLSection -Title "SMBv1 Status (Sicherheit & Performance)" -Content $allSMBHTML
}

# ---------------------------------------------------------------
# 1.9 DAG REPLICATION HEALTH
# ---------------------------------------------------------------
function Get-DAGReplicationHealthInfo {
    <#
    .SYNOPSIS
        Prüft DAG Replication Health (nur wenn DAG konfiguriert ist).
    #>
    Write-Log -Message "=== Sammle DAG Replication Health ===" -Level "INFO"

    $dagHTML = ""

    try {
        # Prüfe ob DAGs in der Organisation vorhanden sind
        $dags = Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue
        
        if (-not $dags -or $dags.Count -eq 0) {
            Write-Log -Message "Keine DAG(s) in der Organisation konfiguriert" -Level "INFO"
            $dagHTML = "<p class='no-data'>Keine Database Availability Groups in dieser Organisation konfiguriert.</p>"
            New-HTMLSection -Title "DAG Replication Health" -Content $dagHTML
            return
        }

        Write-Log -Message "Gefundene DAGs: $($dags.Count)" -Level "INFO"

        foreach ($dag in $dags) {
            try {
                $dagName = $dag.Name
                Write-Log -Message "Prüfe DAG: $dagName" -Level "INFO"

                $serverHTML = "<h3>DAG: $dagName</h3>"
                
                # DAG Status
                $dagInfo = [PSCustomObject]@{
                    "DAG Name"                  = $dag.Name
                    "Member Count"              = $dag.MemberServers.Count
                    "Witness Server"            = $dag.WitnessServer
                    "Witness Dir"               = $dag.WitnessDirectory
                    "Replication Port"          = $dag.ReplicationPort
                }
                
                $serverHTML += "<h4>DAG Informationen</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data @($dagInfo))

                # Replication Health pro Server
                try {
                    $repHealth = Get-MailboxDatabaseCopyStatus -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DatabaseName -like "*" } |
                        Select-Object -Property ServerName, DatabaseName, Status, ReplayQueueLength, CopyQueueLength, ContentIndexState |
                        Sort-Object DatabaseName
                    
                    if ($repHealth) {
                        $repHealthFormatted = foreach ($rep in $repHealth) {
                            $statusIcon = switch ($rep.Status) {
                                "Healthy" { "✅ Healthy" }
                                "Failed" { "🔴 Failed" }
                                "Suspended" { "⚠️ Suspended" }
                                "Seeding" { "🔵 Seeding" }
                                default { $rep.Status }
                            }
                            
                            [PSCustomObject]@{
                                "Server"              = $rep.ServerName
                                "Database"            = $rep.DatabaseName
                                "Status"              = $statusIcon
                                "ReplayQueue"         = $rep.ReplayQueueLength
                                "CopyQueue"           = $rep.CopyQueueLength
                                "ContentIndexStatus"  = $rep.ContentIndexState
                            }
                        }
                        
                        $serverHTML += "<h4>Datenbank-Kopien Status</h4>"
                        $serverHTML += (ConvertTo-HTMLTable -Data $repHealthFormatted)
                    }
                }
                catch {
                    Write-Log -Message "Fehler beim Abrufen der Replication Health: $_" -Level "WARNING"
                    $serverHTML += "<p class='error'>Replication Health konnte nicht abgerufen werden: $_</p>"
                }

                $dagHTML += $serverHTML
            }
            catch {
                Write-Log -Message "Fehler bei DAG $($dag.Name): $_" -Level "ERROR"
                $dagHTML += "<h3>DAG: $($dag.Name)</h3><p class='error'>Fehler: $_</p>"
            }
        }
    }
    catch {
        Write-Log -Message "Fehler bei DAG-Abfrage: $_" -Level "WARNING"
        $dagHTML = "<p class='error'>DAG-Informationen konnten nicht abgerufen werden: $_</p>"
    }

    if ($dagHTML) {
        New-HTMLSection -Title "DAG Replication Health" -Content $dagHTML
    }
}

# ---------------------------------------------------------------
# 1.10 PROCESSOR CORE CHECK (MIT DC:EXCHANGE RATIO)
# ---------------------------------------------------------------
function Get-ProcessorCoreAnalysis {
    <#
    .SYNOPSIS
        Prüft Anzahl der Prozessor-Kerne und berechnet DC:Exchange Ratio.
        Exchange empfiehlt mindestens 4 Kerne.
    #>
    Write-Log -Message "=== Sammle Processor Core Information ===" -Level "INFO"

    $coreHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Processor Cores für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            try {
                $cimSession = New-ServerCimSession -ComputerName $server
                if ($cimSession) {
                    $processorInfo = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor -ErrorAction Stop
                    
                    $totalCores = 0
                    $totalLogicalProcessors = 0
                    $processorData = @()
                    
                    foreach ($proc in $processorInfo) {
                        $totalCores += $proc.NumberOfCores
                        $totalLogicalProcessors += $proc.NumberOfLogicalProcessors
                        
                        $processorData += [PSCustomObject]@{
                            "Prozessor"        = $proc.Name
                            "Physische Kerne"  = $proc.NumberOfCores
                            "Logische Prozesse" = $proc.NumberOfLogicalProcessors
                            "Hyperthreading"   = if ($proc.NumberOfLogicalProcessors -gt $proc.NumberOfCores) { "Aktiviert" } else { "Deaktiviert" }
                        }
                    }
                    
                    # Bewertung
                    $coreStatus = if ($totalCores -ge 4) { "✅ OK (>= 4 Kerne)" } else { "🔴 WARNUNG (< 4 Kerne!)" }
                    $coreRecommendation = "Exchange empfiehlt mindestens 4 physische Kerne für optimale Performance."
                    
                    $serverHTML += "<h4>Prozessor-Übersicht</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data @([PSCustomObject]@{
                        "Gesamt Physische Kerne" = $totalCores
                        "Gesamt Logische Prozesse" = $totalLogicalProcessors
                        "Bewertung" = $coreStatus
                    }))
                    
                    $serverHTML += "<h4>Prozessor Details</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $processorData)
                    $serverHTML += "<p><strong>Hinweis:</strong> $coreRecommendation</p>"
                    
                    Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Log -Message "Processor-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Processor-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $coreHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Processor Core Analysis für ${server}: $_" -Level "ERROR"
            $coreHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Processor Core Analyse" -Content $coreHTML
}

# ---------------------------------------------------------------
# 1.11 RAM REQUIREMENTS CHECK
# ---------------------------------------------------------------
function Get-RAMRequirementsInfo {
    <#
    .SYNOPSIS
        Prüft ob die RAM-Anforderungen für Exchange erfüllt sind.
        Exchange 2019: mindestens 64GB
        Exchange SE: mindestens 128GB empfohlen
    #>
    Write-Log -Message "=== Sammle RAM Requirements Information ===" -Level "INFO"

    $ramHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "RAM-Check für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            try {
                $cimSession = New-ServerCimSession -ComputerName $server
                if ($cimSession) {
                    $osInfo = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction Stop
                    $ramGB = [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 2)
                    
                    # Bestimme die Exchange Edition (wurde bereits in Get-ExchangeEdition bestimmt)
                    $minRAM = if ($script:ExchangeEdition -eq "SE") { 128 } else { 64 }
                    $recommendedRAM = if ($script:ExchangeEdition -eq "SE") { 256 } else { 128 }
                    
                    # Bewertung
                    $ramStatus = if ($ramGB -ge $recommendedRAM) {
                        "✅ SEHR GUT (empfohlenes Minimum erreicht)"
                    }
                    elseif ($ramGB -ge $minRAM) {
                        "🟡 OK (Minimum erfüllt, aber nicht optimal)"
                    }
                    else {
                        "🔴 KRITISCH (unter Minimum!)"
                    }
                    
                    $ramInfo = [PSCustomObject]@{
                        "Exchange Edition"        = $script:ExchangeEdition
                        "Installierte RAM (GB)"   = $ramGB
                        "Minimum erforderlich"    = "$minRAM GB"
                        "Empfohlen"               = "$recommendedRAM GB"
                        "Bewertung"               = $ramStatus
                    }
                    
                    $serverHTML += "<h4>RAM Anforderungen</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data @($ramInfo))
                    
                    if ($script:ExchangeEdition -eq "SE") {
                        $serverHTML += "<p><strong>Hinweis:</strong> Exchange SE empfiehlt mindestens 128GB RAM für optimale Performance. 256GB sind recommended.</p>"
                    }
                    else {
                        $serverHTML += "<p><strong>Hinweis:</strong> Exchange 2019 benötigt mindestens 64GB RAM. 128GB werden empfohlen.</p>"
                    }
                    
                    Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Log -Message "RAM-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>RAM-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $ramHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei RAM Requirements für ${server}: $_" -Level "ERROR"
            $ramHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "RAM Requirements Check" -Content $ramHTML
}

# ---------------------------------------------------------------
# 1.12 CERTIFICATE EXPIRATION CHECK
# ---------------------------------------------------------------
function Get-CertificateExpirationInfo {
    <#
    .SYNOPSIS
        Prüft alle Exchange Zertifikate und deren Ablaufdatum.
        Warnt wenn Ablauf < 30 Tage bevorsteht.
    #>
    Write-Log -Message "=== Sammle Certificate Expiration Information ===" -Level "INFO"

    $certHTML = ""

    try {
        $certs = Get-ExchangeCertificate -ErrorAction SilentlyContinue | 
                 Sort-Object NotAfter
        
        if ($certs) {
            $today = Get-Date
            $certData = @()
            
            foreach ($cert in $certs) {
                $daysUntilExpiry = ($cert.NotAfter - $today).Days
                
                # Bewertung
                if ($daysUntilExpiry -lt 0) {
                    $certStatus = "🔴 ABGELAUFEN"
                }
                elseif ($daysUntilExpiry -le 7) {
                    $certStatus = "🔴 KRITISCH (< 7 Tage)"
                }
                elseif ($daysUntilExpiry -le 30) {
                    $certStatus = "🟡 WARNUNG (< 30 Tage)"
                }
                elseif ($daysUntilExpiry -le 90) {
                    $certStatus = "🟠 BALD ABLAUFEN (< 90 Tage)"
                }
                else {
                    $certStatus = "✅ OK"
                }
                
                $certData += [PSCustomObject]@{
                    "Fingerabdruck"      = $cert.Thumbprint.Substring(0, 10) + "..."
                    "Subject"            = $cert.Subject
                    "Issuer"             = $cert.Issuer
                    "Nicht nach"         = $cert.NotAfter.ToString("dd.MM.yyyy")
                    "Tage bis Ablauf"    = $daysUntilExpiry
                    "Status"             = $certStatus
                }
            }
            
            $certHTML = "<h3>Exchange Zertifikate</h3>"
            $certHTML += (ConvertTo-HTMLTable -Data $certData)
            $certHTML += "<p><strong>Hinweis:</strong> Zertifikate sollten rechtzeitig erneuert werden. Ein Ablauf führt zu Service-Unterbrechungen!</p>"
        }
        else {
            $certHTML = "<p class='no-data'>Keine Exchange Zertifikate gefunden.</p>"
        }
    }
    catch {
        Write-Log -Message "Certificate-Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $certHTML = "<p class='error'>Zertifikate konnten nicht abgerufen werden: $_</p>"
    }

    New-HTMLSection -Title "Certificate Expiration Status" -Content $certHTML
}

# ---------------------------------------------------------------
# 1.13 IIS APPLICATION POOL SETTINGS
# ---------------------------------------------------------------
function Get-IISAppPoolSettingsInfo {
    <#
    .SYNOPSIS
        Prüft IIS Application Pool Recycling über ApplicationHost.config XML (zuverlässiger als AppCmd-Parsing).
    #>
    Write-Log -Message "=== Sammle IIS Application Pool Settings ===" -Level "INFO"

    $iisHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "IIS AppPool Settings für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $appPools = @()
            
            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                
                # Scriptblock: Lese ApplicationHost.config XML direkt
                $getPoolsScript = {
                    param($ComputerName)
                    
                    $pools = @()
                    $configPath = "$env:windir\System32\inetsrv\config\applicationHost.config"
                    
                    try {
                        if (Test-Path $configPath) {
                            [xml]$config = Get-Content $configPath -ErrorAction Stop
                            
                            # Navigiere zu den AppPool <add> Elementen
                            $appPoolParent = $config.configuration.'system.applicationHost'.applicationPools
                            
                            if ($appPoolParent) {
                                # Zugriff auf alle <add> Kinder-Elemente
                                foreach ($poolElement in $appPoolParent.add) {
                                    if ($poolElement.name) {
                                        $poolName = $poolElement.name
                                        $autoStart = if ($poolElement.autoStart -eq "true") { $true } else { $false }
                                        
                                        # recycling.periodicRestart.time
                                        $recycleTime = $null
                                        if ($poolElement.recycling -and $poolElement.recycling.periodicRestart) {
                                            $recycleTime = $poolElement.recycling.periodicRestart.time
                                        }
                                        
                                        # processModel.idleTimeout
                                        $idleTimeout = $null
                                        if ($poolElement.processModel -and $poolElement.processModel.idleTimeout) {
                                            $idleTimeout = $poolElement.processModel.idleTimeout
                                        }
                                        
                                        $pools += [PSCustomObject]@{
                                            Name = $poolName
                                            RecycleTime = $recycleTime
                                            IdleTimeout = $idleTimeout
                                            AutoStart = $autoStart
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Fehler beim Lesen von ApplicationHost.config: $_" -ForegroundColor Red
                    }
                    
                    return $pools
                }
                
                # Führe aus - lokal oder remote
                if ($isLocal) {
                    Write-Log -Message "IIS AppPool lokal aus ApplicationHost.config auslesen..." -Level "INFO"
                    $appPools = @(& $getPoolsScript $env:COMPUTERNAME)
                }
                else {
                    Write-Log -Message "IIS AppPool remote von $server auslesen..." -Level "INFO"
                    $appPools = @(Invoke-Command -ComputerName $server -ScriptBlock $getPoolsScript -ArgumentList $server -ErrorAction Stop)
                }
                
                Write-Log -Message "AppPools gefunden: $($appPools.Count)" -Level "INFO"
            }
            catch {
                Write-Log -Message "AppPool Abfrage für $server fehlgeschlagen: $_" -Level "WARNING"
            }

            if ($appPools -and @($appPools).Count -gt 0) {
                $poolData = @()
                foreach ($pool in $appPools) {
                    # Format RecycleTime
                    $recycleDisplay = if ($pool.RecycleTime -and $pool.RecycleTime -ne "00:00:00") {
                        $pool.RecycleTime -as [string]
                    }
                    else {
                        "-"
                    }
                    
                    $isExchangePool = $pool.Name -like "MSExchange*"
                    $recycleStatus = if ($recycleDisplay -ne "-") {
                        "✅ Konfiguriert ($recycleDisplay)"
                    }
                    elseif ($isExchangePool) {
                        "🟠 WARNUNG (Keine Recycling)"
                    }
                    else {
                        "ℹ️ Standard (Keine Recycling)"
                    }
                    
                    $poolData += [PSCustomObject]@{
                        "App Pool Name"    = $pool.Name
                        "Recycling"        = $recycleDisplay
                        "Idle Timeout"     = $pool.IdleTimeout
                        "AutoStart"        = $pool.AutoStart
                        "Status"           = $recycleStatus
                    }
                }
                
                $serverHTML += "<h4>IIS Application Pools</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data $poolData)
                $serverHTML += "<p><strong>Best Practice:</strong> Exchange Pools sollten Zeit-basierte Recycling haben (z.B. 4h, 14h). Pools ohne Recycling-Zeit nutzen Memory-basierte oder Event-basierte Recycling.</p>"
            }
            else {
                $serverHTML += "<p class='no-data'>Keine App Pools gefunden oder IIS nicht konfiguriert.</p>"
            }

            $iisHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei IIS Settings für ${server}: $_" -Level "ERROR"
            $iisHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "IIS Application Pool Konfiguration" -Content $iisHTML
}

# ---------------------------------------------------------------
# 1.14 EXCHANGE SERVICE STATUS
# ---------------------------------------------------------------
function Get-ExchangeServiceStatusInfo {
    <#
    .SYNOPSIS
        Prüft Status aller wichtigen Exchange Services.
    #>
    Write-Log -Message "=== Sammle Exchange Service Status ===" -Level "INFO"

    $serviceHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Exchange Services für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            try {
                # Wichtige Exchange Services
                $criticalServices = @(
                    "MSExchangeServiceHost",
                    "MSExchangeIS",
                    "MSExchangeADTopology",
                    "MSExchangeMailboxTransport",
                    "MSExchangeTransport",
                    "MSExchangeImap4",
                    "MSExchangePop3",
                    "MSExchangeRepl",
                    "MSExchangeRPC",
                    "MSExchangeIMAP4BE"
                )
                
                $cimSession = New-ServerCimSession -ComputerName $server
                if ($cimSession) {
                    $services = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service `
                        -Filter "Name LIKE 'MSExchange%'" -ErrorAction Stop
                    
                    $serviceData = @()
                    $criticalDown = 0
                    
                    foreach ($svc in $services | Sort-Object Name) {
                        $isCritical = $svc.Name -in $criticalServices
                        
                        $svcStatus = if ($svc.State -eq "Running") { 
                            "✅ Running" 
                        }
                        elseif ($isCritical) {
                            "🔴 KRITISCH GESTOPPT"
                            $criticalDown++
                        }
                        else {
                            "⚠️ Gestoppt"
                        }
                        
                        $svcStartType = if ($svc.StartMode -eq "Auto") { "Auto (erforderlich)" } else { $svc.StartMode }
                        
                        $serviceData += [PSCustomObject]@{
                            "Service Name"    = $svc.Name
                            "Anzeige Name"    = $svc.DisplayName
                            "Status"          = $svcStatus
                            "Starttyp"        = $svcStartType
                            "Kritisch"        = if ($isCritical) { "Ja" } else { "Nein" }
                        }
                    }
                    
                    $serverHTML += "<h4>Exchange Services</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $serviceData)
                    
                    if ($criticalDown -gt 0) {
                        $serverHTML += "<p class='error'><strong>WARNUNG:</strong> $criticalDown kritische Service(s) sind gestoppt!</p>"
                    }
                    else {
                        $serverHTML += "<p><strong>Status:</strong> ✅ Alle kritischen Services sind aktiv.</p>"
                    }
                    
                    Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Log -Message "Service-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Service-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $serviceHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Exchange Service Status für ${server}: $_" -Level "ERROR"
            $serviceHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Exchange Service Status" -Content $serviceHTML
}

# ---------------------------------------------------------------
# 2. WINDOWS & EXCHANGE PATCHES
# ---------------------------------------------------------------
function Get-PatchInformation {
    <#
    .SYNOPSIS
        Sammelt Exchange-Versionsinformationen via Exchange Management Shell.
        (Windows-Hotfixes werden bereits in der Hardware-Sektion über CIM erfasst.)
    #>
    Write-Log -Message "=== Sammle Patch-Informationen (Exchange) ===" -Level "INFO"

    $allPatchHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            # Exchange Server Details via EMS (läuft lokal über Snap-In, kein WinRM nötig)
            $exServerDetail = Get-ExchangeServer -Identity $server -ErrorAction SilentlyContinue |
                Select-Object Name, Edition, AdminDisplayVersion, ServerRole, Site,
                    IsMailboxServer, IsClientAccessServer

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $serverHTML += "<h4>Exchange Server Details (Management Shell)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($exServerDetail) -NoDataMessage "Nicht verfügbar")

            $allPatchHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Patch-Info für ${server}: $_" -Level "ERROR"
            $allPatchHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Exchange Version & Build-Informationen" -Content $allPatchHTML
}

# ---------------------------------------------------------------
# 3. FSMO ROLLEN
# ---------------------------------------------------------------
function Get-FSMORoles {
    <#
    .SYNOPSIS
        Ermittelt die FSMO-Rolleninhaber der Active Directory Gesamtstruktur und Domäne.
    #>
    Write-Log -Message "=== Sammle FSMO-Rollen ===" -Level "INFO"

    $fsmoHTML = ""

    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $forest = Get-ADForest -ErrorAction Stop
        $forestFSMO = [PSCustomObject]@{
            "Schema Master"        = $forest.SchemaMaster
            "Domain Naming Master" = $forest.DomainNamingMaster
        }

        $domain = Get-ADDomain -ErrorAction Stop
        $domainFSMO = [PSCustomObject]@{
            "PDC Emulator"          = $domain.PDCEmulator
            "RID Master"            = $domain.RIDMaster
            "Infrastructure Master" = $domain.InfrastructureMaster
        }

        $fsmoHTML += "<h3>Gesamtstruktur-FSMO-Rollen</h3>"
        $fsmoHTML += (ConvertTo-HTMLTable -Data @($forestFSMO))
        $fsmoHTML += "<h3>Domänen-FSMO-Rollen</h3>"
        $fsmoHTML += (ConvertTo-HTMLTable -Data @($domainFSMO))
    }
    catch {
        Write-Log -Message "Fehler bei FSMO-Rollen: $_" -Level "ERROR"
        $fsmoHTML = "<p class='error'>Fehler: Ist das ActiveDirectory-Modul installiert? $_</p>"
    }

    New-HTMLSection -Title "FSMO-Rollen" -Content $fsmoHTML
}

# ---------------------------------------------------------------
# 4. AD-INFORMATIONEN & SCHEMA VERSION
# ---------------------------------------------------------------
function Get-ADInformation {
    <#
    .SYNOPSIS
        Sammelt AD-Informationen inkl. Schema-Version, Funktionsebenen und Sites.
    #>
    Write-Log -Message "=== Sammle AD-Informationen ===" -Level "INFO"

    $adHTML = ""

    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        # Gesamtstruktur
        $forest = Get-ADForest -ErrorAction Stop
        $forestInfo = [PSCustomObject]@{
            "Gesamtstrukturname"               = $forest.Name
            "Gesamtstrukturfunktionsebene"     = $forest.ForestMode
            "Root-Domäne"                      = $forest.RootDomain
            "Gesamtstruktur-Domänen"           = ($forest.Domains -join ", ")
            "Global Catalog Server"            = ($forest.GlobalCatalogs -join ", ")
            "UPN-Suffixe"                      = if ($forest.UPNSuffixes) { ($forest.UPNSuffixes -join ", ") } else { "Keine zusätzlichen" }
        }

        # Domäne
        $domain = Get-ADDomain -ErrorAction Stop
        $domainInfo = [PSCustomObject]@{
            "Domänenname"           = $domain.DNSRoot
            "NetBIOS-Name"          = $domain.NetBIOSName
            "Domänenfunktionsebene" = $domain.DomainMode
            "Domänencontroller"     = ($domain.ReplicaDirectoryServers -join ", ")
            "DistinguishedName"     = $domain.DistinguishedName
        }

        # AD Schema Version
        $schemaVersionNumber = (Get-ADObject (Get-ADRootDSE).schemaNamingContext -Property objectVersion).objectVersion

        # Exchange Schema Version
        $exchangeSchemaVersion = "Nicht gefunden"
        try {
            $exchangeSchemaPath = "CN=ms-Exch-Schema-Version-Pt," + (Get-ADRootDSE).schemaNamingContext
            $exchangeSchemaVersion = (Get-ADObject $exchangeSchemaPath -Property rangeUpper -ErrorAction Stop).rangeUpper
        }
        catch {
            Write-Log -Message "Exchange Schema nicht gefunden: $_" -Level "WARNING"
        }

        # Exchange Organisation Container
        $exchOrgContainer = $null
        try {
            $configNC = (Get-ADRootDSE).configurationNamingContext
            $exchOrgContainer = Get-ADObject -SearchBase "CN=Microsoft Exchange,CN=Services,$configNC" `
                -Filter { objectClass -eq "msExchOrganizationContainer" } `
                -Property objectVersion, whenCreated -ErrorAction Stop |
                Select-Object @{N='Organisation';E={$_.Name}}, objectVersion,
                    @{N='Erstellt';E={$_.whenCreated.ToString("dd.MM.yyyy")}}
        }
        catch {
            Write-Log -Message "Exchange Org-Container nicht gefunden: $_" -Level "WARNING"
        }

        # Schema-Mapping (inkl. Exchange SE)
        $schemaMap = @{
            17004 = "Exchange Server SE RTM (Subscription Edition)"
            17003 = "Exchange 2019 CU12+ / Exchange SE Preview"
            17002 = "Exchange 2019 CU11"
            17001 = "Exchange 2019 CU8-CU10"
            17000 = "Exchange 2019 RTM - CU7"
            15334 = "Exchange 2016 CU21+"
            15332 = "Exchange 2016 CU7-CU20"
            15326 = "Exchange 2013 SP1+"
        }
        $schemaDescription = if ($schemaMap.ContainsKey([int]$exchangeSchemaVersion)) {
            $schemaMap[[int]$exchangeSchemaVersion]
        }
        else {
            "Unbekannt ($exchangeSchemaVersion)"
        }

        $schemaInfo = [PSCustomObject]@{
            "AD Schema Version (objectVersion)"    = $schemaVersionNumber
            "Exchange Schema Version (rangeUpper)"  = $exchangeSchemaVersion
            "Exchange Schema Beschreibung"          = $schemaDescription
        }

        # AD Sites
        $sites = Get-ADReplicationSite -Filter * -ErrorAction SilentlyContinue |
            Select-Object Name, Description, @{N='Erstellt';E={$_.Created.ToString("dd.MM.yyyy")}}

        # AD Site Links
        $siteLinks = Get-ADReplicationSiteLink -Filter * -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='Kosten';E={$_.Cost}},
                @{N='Replikationsintervall_Min';E={$_.ReplicationFrequencyInMinutes}},
                @{N='Verknüpfte Sites';E={($_.SitesIncluded -join ", ")}}

        # Domänencontroller Übersicht
        $dcList = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue |
            Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly,
                @{N='FSMO_Rollen';E={$_.OperationMasterRoles -join ", "}}

        $adHTML += "<h3>Gesamtstruktur</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data @($forestInfo))
        $adHTML += "<h3>Domäne</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data @($domainInfo))
        $adHTML += "<h3>Schema-Version</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data @($schemaInfo))

        if ($exchOrgContainer) {
            $adHTML += "<h3>Exchange Organisation</h3>"
            $adHTML += (ConvertTo-HTMLTable -Data @($exchOrgContainer))
        }

        $adHTML += "<h3>Domänencontroller</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data $dcList -NoDataMessage "Nicht verfügbar")
        $adHTML += "<h3>AD Sites</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data $sites -NoDataMessage "Nicht verfügbar")
        $adHTML += "<h3>AD Site Links</h3>"
        $adHTML += (ConvertTo-HTMLTable -Data $siteLinks -NoDataMessage "Nicht verfügbar")
    }
    catch {
        Write-Log -Message "Fehler bei AD-Informationen: $_" -Level "ERROR"
        $adHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Active Directory Informationen & Schema-Version" -Content $adHTML
}

# ---------------------------------------------------------------
# 5. EXCHANGE SERVER ÜBERSICHT
# ---------------------------------------------------------------
function Get-ExchangeServerOverview {
    <#
    .SYNOPSIS
        Sammelt eine Übersicht aller Exchange Server in der Organisation.
    #>
    Write-Log -Message "=== Sammle Exchange Server Übersicht ===" -Level "INFO"

    $exHTML = ""

    try {
        $allExServers = Get-ExchangeServer -ErrorAction Stop |
            Select-Object Name, Edition, AdminDisplayVersion, ServerRole, Site,
                @{N='IsMailbox';E={$_.IsMailboxServer}},
                @{N='IsClientAccess';E={$_.IsClientAccessServer}},
                @{N='IsEdge';E={$_.IsEdgeServer}},
                @{N='ExchangeEdition';E={
                    if ($_.AdminDisplayVersion -match 'Version\s+(\d+\.\d+)') {
                        $v = $Matches[1]
                        if ($v -eq "15.0") { '2013' }
                        elseif ($v -eq "15.1") { '2016' }
                        elseif ($v -eq "15.2") {
                            if ($_.AdminDisplayVersion -match 'Build\s+(\d+)' -and [int]$Matches[1] -ge 2000) { 'SE' }
                            else { '2019' }
                        }
                        else { 'Unbekannt' }
                    } else { 'Unbekannt' }
                }},
                WhenCreated, WhenChanged

        $exHTML += "<h3>Alle Exchange Server in der Organisation</h3>"
        $exHTML += "<p><strong>Erkannte Exchange Edition: $($script:ExchangeEdition)</strong></p>"
        $exHTML += (ConvertTo-HTMLTable -Data $allExServers)
    }
    catch {
        Write-Log -Message "Fehler bei Exchange Server Übersicht: $_" -Level "ERROR"
        $exHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange Server Übersicht" -Content $exHTML
}

# ---------------------------------------------------------------
# 6. ORGANISATIONS-KONFIGURATION
# ---------------------------------------------------------------
function Get-OrganizationConfigInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Exchange-Organisationskonfiguration.
    #>
    Write-Log -Message "=== Sammle Organisationskonfiguration ===" -Level "INFO"

    $orgHTML = ""

    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction Stop |
            Select-Object @{N='Organisationsname';E={$_.Name}},
                @{N='DefaultPublicFolderMailbox';E={$_.DefaultPublicFolderMailbox}},
                @{N='MailTipsAllTipsEnabled';E={$_.MailTipsAllTipsEnabled}},
                @{N='MailTipsExternalRecipientsTipsEnabled';E={$_.MailTipsExternalRecipientsTipsEnabled}},
                @{N='MailTipsGroupMetricsEnabled';E={$_.MailTipsGroupMetricsEnabled}},
                @{N='MailTipsLargeAudienceThreshold';E={$_.MailTipsLargeAudienceThreshold}},
                @{N='MaxSendSize';E={$_.MaxSendSize}},
                @{N='MaxReceiveSize';E={$_.MaxReceiveSize}},
                @{N='EwsEnabled';E={$_.EwsEnabled}},
                @{N='MapiHttpEnabled';E={$_.MapiHttpEnabled}},
                @{N='OAuth2ClientProfileEnabled';E={$_.OAuth2ClientProfileEnabled}},
                @{N='PublicFoldersEnabled';E={$_.PublicFoldersEnabled}}

        $orgHTML += "<h3>Exchange Organisationskonfiguration</h3>"
        $orgHTML += (ConvertTo-HTMLTable -Data @($orgConfig))

        # Transport Config
        $transportConfig = Get-TransportConfig -ErrorAction SilentlyContinue |
            Select-Object @{N='MaxSendSize';E={$_.MaxSendSize}},
                @{N='MaxReceiveSize';E={$_.MaxReceiveSize}},
                @{N='ExternalPostmasterAddress';E={$_.ExternalPostmasterAddress}},
                @{N='InternalSMTPServers';E={$_.InternalSMTPServers -join ", "}},
                @{N='TLSSendDomainSecureList';E={$_.TLSSendDomainSecureList -join ", "}},
                @{N='TLSReceiveDomainSecureList';E={$_.TLSReceiveDomainSecureList -join ", "}},
                @{N='ShadowRedundancyEnabled';E={$_.ShadowRedundancyEnabled}},
                @{N='SafetyNetHoldTime';E={$_.SafetyNetHoldTime}},
                @{N='JournalingReportNdrTo';E={$_.JournalingReportNdrTo}}

        $orgHTML += "<h3>Transport-Konfiguration</h3>"
        $orgHTML += (ConvertTo-HTMLTable -Data @($transportConfig))
    }
    catch {
        Write-Log -Message "Fehler bei Organisationskonfiguration: $_" -Level "ERROR"
        $orgHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange Organisationskonfiguration" -Content $orgHTML
}

# ---------------------------------------------------------------
# 7. EXCHANGE URLS / VIRTUELLE VERZEICHNISSE
# ---------------------------------------------------------------
function Get-ExchangeURLs {
    <#
    .SYNOPSIS
        Sammelt alle konfigurierten URLs aller virtuellen Verzeichnisse mit Authentication.
        Zusätzlich erfasst EEMS (Exchange Emergency Mitigation Service) Status und Informationen.
    #>
    Write-Log -Message "=== Sammle Exchange URLs ===" -Level "INFO"

    $urlHTML = ""

    try {
        # OWA mit Authentication
        $owa = Get-OwaVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={
                    $auth = @()
                    if ($_.BasicAuthentication) { $auth += "Basic" }
                    if ($_.WindowsAuthentication) { $auth += "Windows" }
                    if ($_.FormsAuthentication) { $auth += "Forms" }
                    $auth -join ", "
                }}

        # ECP mit Authentication
        $ecp = Get-EcpVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={
                    $auth = @()
                    if ($_.BasicAuthentication) { $auth += "Basic" }
                    if ($_.WindowsAuthentication) { $auth += "Windows" }
                    if ($_.FormsAuthentication) { $auth += "Forms" }
                    $auth -join ", "
                }}

        # ActiveSync mit Authentication
        $eas = Get-ActiveSyncVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={
                    $auth = @()
                    if ($_.BasicAuthentication) { $auth += "Basic" }
                    if ($_.WindowsAuthentication) { $auth += "Windows" }
                    $auth -join ", "
                }}

        # EWS mit Authentication
        $ews = Get-WebServicesVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={
                    $auth = @()
                    if ($_.BasicAuthentication) { $auth += "Basic" }
                    if ($_.WindowsAuthentication) { $auth += "Windows" }
                    $auth -join ", "
                }},
                MRSProxyEnabled

        # MAPI mit Authentication
        $mapi = Get-MapiVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={$_.IISAuthenticationMethods -join ", "}}

        # OAB mit Authentication
        $oabVDir = Get-OabVirtualDirectory -ErrorAction SilentlyContinue |
            Select-Object Server, Name,
                @{N='InternalURL';E={$_.InternalUrl}},
                @{N='ExternalURL';E={$_.ExternalUrl}},
                @{N='Authentifizierung';E={
                    $auth = @()
                    if ($_.BasicAuthentication) { $auth += "Basic" }
                    if ($_.WindowsAuthentication) { $auth += "Windows" }
                    $auth -join ", "
                }}

        # Autodiscover
        $autodiscover = Get-ClientAccessServer -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='AutoDiscoverServiceInternalUri';E={$_.AutoDiscoverServiceInternalUri}}

        # Outlook Anywhere
        $outlookAnywhere = Get-OutlookAnywhere -ErrorAction SilentlyContinue |
            Select-Object Server,
                @{N='InternalHostname';E={$_.InternalHostname}},
                @{N='ExternalHostname';E={$_.ExternalHostname}},
                @{N='InternalAuth';E={$_.InternalClientAuthenticationMethod}},
                @{N='ExternalAuth';E={$_.ExternalClientAuthenticationMethod}},
                SSLOffloading

        $urlHTML += "<h3>OWA Virtual Directory</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $owa)
        $urlHTML += "<h3>ECP Virtual Directory</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $ecp)
        $urlHTML += "<h3>ActiveSync Virtual Directory</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $eas)
        $urlHTML += "<h3>Exchange Web Services (EWS)</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $ews)
        $urlHTML += "<h3>MAPI Virtual Directory</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $mapi)
        $urlHTML += "<h3>OAB Virtual Directory</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $oabVDir)
        $urlHTML += "<h3>Autodiscover</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $autodiscover)
        $urlHTML += "<h3>Outlook Anywhere</h3>"
        $urlHTML += (ConvertTo-HTMLTable -Data $outlookAnywhere)

        # ===== EXCHANGE EMERGENCY MITIGATION SERVICE (EEMS) =====
        $urlHTML += "<h2>Exchange Emergency Mitigation Service (EEMS)</h2>"
        
        foreach ($server in ($ExchangeServers | Select-Object -Unique)) {
            $eemsSectionHTML = ""
            
            try {
                $eemsSectionHTML += "<h3 class='server-break'>Server: $server - EEMS Status</h3>"
                
                # EEMS Windows Service Status
                $eemService = Get-Service -ComputerName $server -Name "MSExchangeServiceHost" -ErrorAction SilentlyContinue
                if ($eemService) {
                    $eemsSectionHTML += "<p><strong>Windows Service (MSExchangeServiceHost):</strong> $($eemService.Status)</p>"
                } else {
                    $eemsSectionHTML += "<p><strong>Windows Service:</strong> <span class='warning'>Nicht gefunden</span></p>"
                }

                # Pattern Service Erreichbarkeit
                $patternStatus = "Checking..."
                try {
                    $testConnection = Test-NetConnection -ComputerName $server -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
                    $patternStatus = if ($testConnection) { "200 - Reachable" } else { "Unreachable" }
                } catch {
                    $patternStatus = "Error - $($_.Exception.Message)"
                }
                $eemsSectionHTML += "<p><strong>Pattern Service:</strong> $patternStatus</p>"

                # Mitigations via Invoke-Command
                $mitigations = Invoke-Command -ComputerName $server -ScriptBlock {
                    try {
                        $emsRegPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EEMS"
                        if (Test-Path -LiteralPath $emsRegPath) {
                            $regProps = Get-ItemProperty -LiteralPath $emsRegPath -ErrorAction SilentlyContinue
                            if ($regProps) {
                                return ($regProps | Select-Object -Property * -ExcludeProperty "PS*" | ConvertTo-Json -Depth 1)
                            }
                        }
                        return $null
                    } catch {
                        return $null
                    }
                } -ErrorAction SilentlyContinue

                if ($mitigations) {
                    try {
                        $mitObj = $mitigations | ConvertFrom-Json
                        $eemsSectionHTML += "<p><strong>Applied Mitigations:</strong></p><ul>"
                        foreach ($prop in ($mitObj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)) {
                            $eemsSectionHTML += "<li>${prop}: $($mitObj.$prop)</li>"
                        }
                        $eemsSectionHTML += "</ul>"
                    } catch {
                        $eemsSectionHTML += "<p><strong>Applied Mitigations:</strong> Nicht verfügbar</p>"
                    }
                } else {
                    $eemsSectionHTML += "<p><strong>Applied Mitigations:</strong> Keine Mitigationen aktiv</p>"
                }

                # Telemetry Status
                $telemetryStatus = Invoke-Command -ComputerName $server -ScriptBlock {
                    try {
                        $telePath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EEMS\Telemetry"
                        if (Test-Path -LiteralPath $telePath) {
                            $regProps = Get-ItemProperty -LiteralPath $telePath -ErrorAction SilentlyContinue
                            if ($regProps -and $regProps.Enabled) {
                                return $regProps.Enabled
                            }
                        }
                        return "Unknown"
                    } catch {
                        return "Unknown"
                    }
                } -ErrorAction SilentlyContinue

                $eemsSectionHTML += "<p><strong>Telemetry Enabled:</strong> $(if ($telemetryStatus -eq 1) { "True" } elseif ($telemetryStatus -eq 0) { "False" } else { "Unknown" })</p>"

                # IIS Module Anomalies
                $iisAnomalies = Invoke-Command -ComputerName $server -ScriptBlock {
                    try {
                        $iisPath = "HKLM:\SYSTEM\CurrentControlSet\Services\W3SVC\Parameters\ManagedPipelineMode"
                        if (Test-Path $iisPath) {
                            return $false  # Vereinfacht - würde Modul-Prüfung benötigen
                        }
                        return $false
                    } catch {
                        return "Unknown"
                    }
                } -ErrorAction SilentlyContinue

                $eemsSectionHTML += "<p><strong>IIS Module Anomalies Detected:</strong> $(if ($iisAnomalies) { "True" } else { "False" })</p>"

                # Mitigation Service Info
                $eemsSectionHTML += "<p><em>Hinweis: Um detailliertere EEMS-Informationen zu erhalten, führen Sie folgendes Skript auf dem Exchange Server aus: Get-Mitigations.ps1</em></p>"

            } catch {
                Write-Log -Message "Fehler bei EEMS Abfrage auf $server : $_" -Level "WARNING"
                $eemsSectionHTML += "<p class='error'>Fehler bei EEMS-Abfrage: $_</p>"
            }

            $urlHTML += $eemsSectionHTML
        }
    }
    catch {
        Write-Log -Message "Fehler bei Exchange URLs: $_" -Level "ERROR"
        $urlHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange URLs / Virtuelle Verzeichnisse / Namespace & EEMS" -Content $urlHTML
}

# ---------------------------------------------------------------
# 8. DATENBANKEN & DAG
# ---------------------------------------------------------------
function Get-DatabaseAndDAGInfo {
    <#
    .SYNOPSIS
        Sammelt Informationen zu Datenbanken und DAG-Konfiguration.
    #>
    Write-Log -Message "=== Sammle Datenbank- und DAG-Informationen ===" -Level "INFO"

    $dbHTML = ""

    try {
        # Mailbox-Datenbanken
        $databases = Get-MailboxDatabase -Status -ErrorAction Stop | Select-Object Name, Server,
            @{N='EdbFilePath';E={$_.EdbFilePath}},
            @{N='LogFolderPath';E={$_.LogFolderPath}},
            @{N='DatabaseSize_GB';E={if($_.DatabaseSize){[math]::Round($_.DatabaseSize.ToBytes()/1GB,2)}else{"N/A"}}},
            @{N='AvailableSpace_GB';E={if($_.AvailableNewMailboxSpace){[math]::Round($_.AvailableNewMailboxSpace.ToBytes()/1GB,2)}else{"N/A"}}},
            @{N='Mailboxanzahl';E={(Get-Mailbox -Database $_.Name -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count}},
            MasterServerOrAvailabilityGroup,
            @{N='CircularLogging';E={$_.CircularLoggingEnabled}},
            @{N='BackupInProgress';E={$_.BackupInProgress}},
            @{N='LastFullBackup';E={if($_.LastFullBackup){$_.LastFullBackup.ToString("dd.MM.yyyy HH:mm")}else{"NIE!"}}},
            @{N='LastIncrBackup';E={if($_.LastIncrementalBackup){$_.LastIncrementalBackup.ToString("dd.MM.yyyy HH:mm")}else{"Nie"}}},
            @{N='Mounted';E={$_.Mounted}},
            @{N='ProhibitSendQuota';E={$_.ProhibitSendQuota}},
            @{N='ProhibitSendReceiveQuota';E={$_.ProhibitSendReceiveQuota}},
            @{N='IssueWarningQuota';E={$_.IssueWarningQuota}},
            @{N='DeletedItemRetention';E={$_.DeletedItemRetention}},
            @{N='MailboxRetention';E={$_.MailboxRetention}},
            Recovery, ReplicationType

        $dbHTML += "<h3>Mailbox-Datenbanken</h3>"
        $dbHTML += (ConvertTo-HTMLTable -Data $databases)

        # Datenbankkopien
        $dbCopies = Get-MailboxDatabaseCopyStatus * -ErrorAction SilentlyContinue |
            Select-Object Name, Status,
                @{N='CopyQueueLength';E={$_.CopyQueueLength}},
                @{N='ReplayQueueLength';E={$_.ReplayQueueLength}},
                @{N='ContentIndexState';E={$_.ContentIndexState}},
                @{N='ContentIndexErrorMsg';E={$_.ContentIndexErrorMessage}},
                LastInspectedLogTime, ActivationPreference

        if ($dbCopies) {
            $dbHTML += "<h3>Datenbankkopie-Status (DAG)</h3>"
            $dbHTML += (ConvertTo-HTMLTable -Data $dbCopies)
        }

        # DAG Konfiguration
        $dags = Get-DatabaseAvailabilityGroup -Status -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='Mitglieder';E={$_.Servers -join ", "}},
                @{N='WitnessServer';E={$_.WitnessServer}},
                @{N='WitnessDirectory';E={$_.WitnessDirectory}},
                @{N='AltWitnessServer';E={$_.AlternateWitnessServer}},
                @{N='AltWitnessDirectory';E={$_.AlternateWitnessDirectory}},
                @{N='DAC_Mode';E={$_.DatacenterActivationMode}},
                @{N='OperationalServers';E={$_.OperationalServers -join ", "}},
                @{N='PrimaryActiveManager';E={$_.PrimaryActiveManager}},
                @{N='DAG_IP';E={$_.DatabaseAvailabilityGroupIpv4Addresses -join ", "}},
                @{N='ReplicationPort';E={$_.ReplicationPort}},
                @{N='NetworkCompression';E={$_.NetworkCompression}},
                @{N='NetworkEncryption';E={$_.NetworkEncryption}}

        if ($dags) {
            $dbHTML += "<h3>DAG Konfiguration</h3>"
            $dbHTML += (ConvertTo-HTMLTable -Data $dags)

            # DAG Netzwerke
            foreach ($dag in (Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue)) {
                try {
                    $dagNetworks = Get-DatabaseAvailabilityGroupNetwork -Identity $dag.Name -ErrorAction SilentlyContinue |
                        Select-Object Name,
                            @{N='Subnets';E={$_.Subnets -join ", "}},
                            @{N='MapiAccessEnabled';E={$_.MapiAccessEnabled}},
                            @{N='ReplicationEnabled';E={$_.ReplicationEnabled}},
                            @{N='IgnoreNetwork';E={$_.IgnoreNetwork}}

                    if ($dagNetworks) {
                        $dbHTML += "<h3>DAG Netzwerke: $($dag.Name)</h3>"
                        $dbHTML += (ConvertTo-HTMLTable -Data $dagNetworks)
                    }
                }
                catch {
                    Write-Log -Message "DAG Netzwerke Fehler: $_" -Level "WARNING"
                }
            }
        }
    }
    catch {
        Write-Log -Message "Fehler bei Datenbank/DAG-Informationen: $_" -Level "ERROR"
        $dbHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Datenbanken & DAG-Konfiguration" -Content $dbHTML
}

# ---------------------------------------------------------------
# 9. PUBLIC FOLDER
# ---------------------------------------------------------------
function Get-PublicFolderInfo {
    <#
    .SYNOPSIS
        Sammelt Public Folder Informationen.
    #>
    Write-Log -Message "=== Sammle Public Folder Informationen ===" -Level "INFO"

    $pfHTML = ""

    try {
        $pfMailboxes = Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object Name, Alias, Database,
                @{N='IsRootPFMailbox';E={$_.IsRootPublicFolderMailbox}},
                WhenCreated

        $pfHTML += "<h3>Public Folder Mailboxen</h3>"
        $pfHTML += (ConvertTo-HTMLTable -Data $pfMailboxes -NoDataMessage "Keine Public Folder Mailboxen vorhanden.")

        $publicFolders = Get-PublicFolder -Recurse -ResultSize 100 -ErrorAction SilentlyContinue |
            Select-Object Name, Identity, FolderClass,
                @{N='MailEnabled';E={$_.MailEnabled}},
                @{N='HasSubFolders';E={$_.HasSubFolders}},
                ParentPath

        $pfHTML += "<h3>Public Folder Struktur (max. 100)</h3>"
        $pfHTML += (ConvertTo-HTMLTable -Data $publicFolders -NoDataMessage "Keine Public Folder vorhanden.")

        $pfStats = Get-PublicFolderStatistics -ResultSize 100 -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='TotalItemSize';E={$_.TotalItemSize}},
                ItemCount, FolderPath,
                @{N='LastModified';E={if($_.LastModificationTime){$_.LastModificationTime.ToString("dd.MM.yyyy HH:mm")}else{"N/A"}}}

        $pfHTML += "<h3>Public Folder Statistiken (max. 100)</h3>"
        $pfHTML += (ConvertTo-HTMLTable -Data $pfStats -NoDataMessage "Keine Statistiken verfügbar.")
    }
    catch {
        Write-Log -Message "Fehler bei Public Folder: $_" -Level "ERROR"
        $pfHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Public Folder" -Content $pfHTML
}

# ---------------------------------------------------------------
# 10. SENDE-CONNECTOREN
# ---------------------------------------------------------------
function Get-SendConnectorInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Send Connectoren.
    #>
    Write-Log -Message "=== Sammle Sendeconnectoren ===" -Level "INFO"

    $scHTML = ""

    try {
        $sendConnectors = Get-SendConnector -ErrorAction Stop |
            Select-Object Name,
                @{N='AddressSpaces';E={$_.AddressSpaces -join ", "}},
                @{N='SmartHosts';E={$_.SmartHosts -join ", "}},
                @{N='SourceTransportServers';E={$_.SourceTransportServers -join ", "}},
                @{N='MaxMessageSize';E={$_.MaxMessageSize}},
                @{N='Enabled';E={$_.Enabled}},
                @{N='DNSRoutingEnabled';E={$_.DNSRoutingEnabled}},
                @{N='Port';E={$_.Port}},
                @{N='RequireTLS';E={$_.RequireTLS}},
                @{N='TlsAuthLevel';E={$_.TlsAuthLevel}},
                @{N='TlsDomain';E={$_.TlsDomain}},
                @{N='SmartHostAuthMechanism';E={$_.SmartHostAuthMechanism}},
                @{N='ProtocolLoggingLevel';E={$_.ProtocolLoggingLevel}},
                @{N='FrontendProxyEnabled';E={$_.FrontendProxyEnabled}},
                @{N='Fqdn';E={$_.Fqdn}},
                @{N='Comment';E={$_.Comment}}

        $scHTML += (ConvertTo-HTMLTable -Data $sendConnectors -NoDataMessage "Keine Sendeconnectoren konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei Sendeconnectoren: $_" -Level "ERROR"
        $scHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Sendeconnectoren" -Content $scHTML
}

# ---------------------------------------------------------------
# 11. EMPFANGSCONNECTOREN
# ---------------------------------------------------------------
function Get-ReceiveConnectorInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Empfangsconnectoren.
    #>
    Write-Log -Message "=== Sammle Empfangsconnectoren ===" -Level "INFO"

    $rcHTML = ""

    try {
        $receiveConnectors = Get-ReceiveConnector -ErrorAction Stop |
            Select-Object Identity, Server,
                @{N='Bindings';E={$_.Bindings -join ", "}},
                @{N='RemoteIPRanges';E={($_.RemoteIPRanges | Select-Object -First 10) -join ", "}},
                @{N='Enabled';E={$_.Enabled}},
                @{N='TransportRole';E={$_.TransportRole}},
                @{N='AuthMechanism';E={$_.AuthMechanism}},
                @{N='PermissionGroups';E={$_.PermissionGroups}},
                @{N='MaxMessageSize';E={$_.MaxMessageSize}},
                @{N='MaxRecipientsPerMessage';E={$_.MaxRecipientsPerMessage}},
                @{N='RequireTLS';E={$_.RequireTLS}},
                @{N='ProtocolLoggingLevel';E={$_.ProtocolLoggingLevel}},
                @{N='Fqdn';E={$_.Fqdn}},
                @{N='Banner';E={$_.Banner}},
                @{N='Comment';E={$_.Comment}}

        $rcHTML += (ConvertTo-HTMLTable -Data $receiveConnectors -NoDataMessage "Keine Empfangsconnectoren konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei Empfangsconnectoren: $_" -Level "ERROR"
        $rcHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Empfangsconnectoren" -Content $rcHTML
}

# ---------------------------------------------------------------
# 12. REMOTE DOMAINS
# ---------------------------------------------------------------
function Get-RemoteDomainInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Remote Domains.
    #>
    Write-Log -Message "=== Sammle Remote Domains ===" -Level "INFO"

    $rdHTML = ""

    try {
        $remoteDomains = Get-RemoteDomain -ErrorAction Stop |
            Select-Object Name, DomainName,
                @{N='IsInternal';E={$_.IsInternal}},
                @{N='AllowedOOFType';E={$_.AllowedOOFType}},
                @{N='AutoForwardEnabled';E={$_.AutoForwardEnabled}},
                @{N='AutoReplyEnabled';E={$_.AutoReplyEnabled}},
                @{N='DeliveryReportEnabled';E={$_.DeliveryReportEnabled}},
                @{N='NDREnabled';E={$_.NDREnabled}},
                @{N='TNEFEnabled';E={$_.TNEFEnabled}},
                @{N='CharacterSet';E={$_.CharacterSet}},
                @{N='ContentType';E={$_.ContentType}}

        $rdHTML += (ConvertTo-HTMLTable -Data $remoteDomains -NoDataMessage "Keine Remote Domains konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei Remote Domains: $_" -Level "ERROR"
        $rdHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Remote Domains" -Content $rdHTML
}

# ---------------------------------------------------------------
# 13. ACCEPTED DOMAINS
# ---------------------------------------------------------------
function Get-AcceptedDomainInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Accepted Domains.
    #>
    Write-Log -Message "=== Sammle Accepted Domains ===" -Level "INFO"

    $adHTML = ""

    try {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction Stop |
            Select-Object Name, DomainName, DomainType, Default,
                @{N='MatchSubDomains';E={$_.MatchSubDomains}},
                @{N='AddressBookEnabled';E={$_.AddressBookEnabled}}

        $adHTML += (ConvertTo-HTMLTable -Data $acceptedDomains -NoDataMessage "Keine Accepted Domains konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei Accepted Domains: $_" -Level "ERROR"
        $adHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Accepted Domains" -Content $adHTML
}

# ---------------------------------------------------------------
# 14. MX RECORDS & DNS-PRÜFUNG
# ---------------------------------------------------------------
function Get-MXRecordInfo {
    <#
    .SYNOPSIS
        Ermittelt MX-Records, SPF, DMARC und Autodiscover-DNS für alle Accepted Domains.
    #>
    Write-Log -Message "=== Sammle MX-Record Informationen ===" -Level "INFO"

    $mxHTML = ""

    try {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction Stop
        $mxResults = @()
        $dnsCheckResults = @()

        foreach ($domain in $acceptedDomains) {
            $domainName = $domain.DomainName.ToString()
            Write-Log -Message "DNS-Abfragen für: $domainName" -Level "INFO"

            # MX-Records
            try {
                $mxParams = @{ Name = $domainName; Type = "MX"; ErrorAction = "Stop" }
                if ($script:DNSServer -ne "") { $mxParams.Add("Server", $script:DNSServer) }
                $mxRecords = Resolve-DnsName @mxParams

                foreach ($mx in $mxRecords) {
                    if ($mx.Type -eq "MX") {
                        $mxResults += [PSCustomObject]@{
                            Domäne    = $domainName
                            MX_Host   = $mx.NameExchange
                            Priorität = $mx.Preference
                            TTL       = $mx.TTL
                        }
                    }
                }
            }
            catch {
                $mxResults += [PSCustomObject]@{
                    Domäne    = $domainName
                    MX_Host   = "Kein MX-Record gefunden"
                    Priorität = "N/A"
                    TTL       = "N/A"
                }
                Write-Log -Message "MX-Abfrage fehlgeschlagen für ${domainName}: $_" -Level "WARNING"
            }

            # SPF, DMARC, Autodiscover
            try {
                $spf = try { (Resolve-DnsName -Name $domainName -Type TXT -ErrorAction Stop |
                    Where-Object { $_.Strings -like "*v=spf1*" }).Strings -join " " } catch { "Nicht gefunden" }
                $dmarc = try { (Resolve-DnsName -Name "_dmarc.$domainName" -Type TXT -ErrorAction Stop).Strings -join " " } catch { "Nicht gefunden" }
                $autodiscoverDNS = try {
                    $adResult = Resolve-DnsName -Name "autodiscover.$domainName" -Type CNAME -ErrorAction SilentlyContinue
                    if ($adResult.NameHost) { $adResult.NameHost }
                    else {
                        $adA = Resolve-DnsName -Name "autodiscover.$domainName" -Type A -ErrorAction SilentlyContinue
                        if ($adA) { $adA.IPAddress -join ", " } else { "Nicht gefunden" }
                    }
                } catch { "Nicht gefunden" }

                $dnsCheckResults += [PSCustomObject]@{
                    Domäne           = $domainName
                    SPF_Record       = $spf
                    DMARC_Record     = $dmarc
                    Autodiscover_DNS = $autodiscoverDNS
                }
            }
            catch {
                Write-Log -Message "DNS-Check fehlgeschlagen für ${domainName}: $_" -Level "WARNING"
            }
        }

        $mxHTML += "<h3>MX-Records</h3>"
        $mxHTML += (ConvertTo-HTMLTable -Data $mxResults -NoDataMessage "Keine MX-Records ermittelt.")
        $mxHTML += "<h3>DNS-Prüfung (SPF, DMARC, Autodiscover)</h3>"
        $mxHTML += (ConvertTo-HTMLTable -Data $dnsCheckResults -NoDataMessage "Keine DNS-Ergebnisse.")
    }
    catch {
        Write-Log -Message "Fehler bei MX-Records: $_" -Level "ERROR"
        $mxHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "MX-Records & DNS-Prüfung" -Content $mxHTML
}

# ---------------------------------------------------------------
# 15. E-MAIL-ADRESSRICHTLINIEN
# ---------------------------------------------------------------
function Get-EmailAddressPolicyInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle E-Mail-Adressrichtlinien.
    #>
    Write-Log -Message "=== Sammle E-Mail-Adressrichtlinien ===" -Level "INFO"

    $eapHTML = ""

    try {
        $policies = Get-EmailAddressPolicy -ErrorAction Stop |
            Select-Object Name, Priority,
                @{N='EnabledEmailAddressTemplates';E={$_.EnabledEmailAddressTemplates -join "; "}},
                @{N='RecipientFilter';E={$_.RecipientFilter}},
                @{N='RecipientFilterApplied';E={$_.RecipientFilterApplied}},
                @{N='EnabledPrimarySMTPAddressTemplate';E={$_.EnabledPrimarySMTPAddressTemplate}},
                WhenCreated, WhenChanged

        $eapHTML += (ConvertTo-HTMLTable -Data $policies -NoDataMessage "Keine E-Mail-Adressrichtlinien konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei E-Mail-Adressrichtlinien: $_" -Level "ERROR"
        $eapHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "E-Mail-Adressrichtlinien" -Content $eapHTML
}

# ---------------------------------------------------------------
# 16. ADRESSLISTEN
# ---------------------------------------------------------------
function Get-AddressListInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Adresslisten und GAL.
    #>
    Write-Log -Message "=== Sammle Adresslisten ===" -Level "INFO"

    $alHTML = ""

    try {
        $addressLists = Get-AddressList -ErrorAction Stop |
            Select-Object Name, Path,
                @{N='RecipientFilter';E={$_.RecipientFilter}},
                @{N='IncludedRecipients';E={$_.IncludedRecipients}},
                WhenCreated

        $gal = Get-GlobalAddressList -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='RecipientFilter';E={$_.RecipientFilter}},
                @{N='IncludedRecipients';E={$_.IncludedRecipients}},
                IsDefaultGlobalAddressList

        $abp = Get-AddressBookPolicy -ErrorAction SilentlyContinue |
            Select-Object Name, AddressLists, RoomList, GlobalAddressList, OfflineAddressBook

        $alHTML += "<h3>Adresslisten</h3>"
        $alHTML += (ConvertTo-HTMLTable -Data $addressLists -NoDataMessage "Keine Adresslisten konfiguriert.")
        $alHTML += "<h3>Globale Adressliste (GAL)</h3>"
        $alHTML += (ConvertTo-HTMLTable -Data $gal -NoDataMessage "Keine GAL gefunden.")

        if ($abp) {
            $alHTML += "<h3>Adressbuchrichtlinien</h3>"
            $alHTML += (ConvertTo-HTMLTable -Data $abp)
        }
    }
    catch {
        Write-Log -Message "Fehler bei Adresslisten: $_" -Level "ERROR"
        $alHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Adresslisten & Globale Adressliste" -Content $alHTML
}

# ---------------------------------------------------------------
# 17. OFFLINE-ADRESSBUCH
# ---------------------------------------------------------------
function Get-OABInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Offline-Adressbücher.
    #>
    Write-Log -Message "=== Sammle Offline-Adressbuch ===" -Level "INFO"

    $oabHTML = ""

    try {
        $oabs = Get-OfflineAddressBook -ErrorAction Stop |
            Select-Object Name,
                @{N='Server';E={$_.Server}},
                @{N='AddressLists';E={$_.AddressLists -join ", "}},
                @{N='GeneratingMailbox';E={$_.GeneratingMailbox}},
                @{N='IsDefault';E={$_.IsDefault}},
                @{N='Versions';E={$_.Versions -join ", "}},
                @{N='VirtualDirectories';E={$_.VirtualDirectories -join ", "}},
                @{N='GlobalWebDistEnabled';E={$_.GlobalWebDistributionEnabled}},
                WhenCreated

        $oabHTML += (ConvertTo-HTMLTable -Data $oabs -NoDataMessage "Keine OABs konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei OAB: $_" -Level "ERROR"
        $oabHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Offline-Adressbuch (OAB)" -Content $oabHTML
}

# ---------------------------------------------------------------
# 18. MOBILE DEVICE POLICIES
# ---------------------------------------------------------------
function Get-MobileDeviceInfo {
    <#
    .SYNOPSIS
        Dokumentiert Mobile Device Policies und Statistiken.
    #>
    Write-Log -Message "=== Sammle Mobile Device Informationen ===" -Level "INFO"

    $mdHTML = ""

    try {
        $mdPolicies = Get-MobileDeviceMailboxPolicy -ErrorAction Stop |
            Select-Object Name,
                @{N='IsDefault';E={$_.IsDefault}},
                @{N='PasswordEnabled';E={$_.PasswordEnabled}},
                @{N='MinPasswordLength';E={$_.MinPasswordLength}},
                @{N='AlphanumericPwRequired';E={$_.AlphanumericPasswordRequired}},
                @{N='AllowSimplePassword';E={$_.AllowSimplePassword}},
                @{N='MaxInactivityTimeLock';E={$_.MaxInactivityTimeLock}},
                @{N='MaxPwFailedAttempts';E={$_.MaxPasswordFailedAttempts}},
                @{N='PasswordExpiration';E={$_.PasswordExpiration}},
                @{N='AllowNonProvisionable';E={$_.AllowNonProvisionableDevices}},
                @{N='DeviceEncryptionEnabled';E={$_.DeviceEncryptionEnabled}},
                @{N='AllowCamera';E={$_.AllowCamera}},
                @{N='AllowWiFi';E={$_.AllowWiFi}},
                @{N='AllowBrowser';E={$_.AllowBrowser}},
                @{N='RequireDeviceEncryption';E={$_.RequireDeviceEncryption}}

        $mdHTML += "<h3>Mobile Device Mailbox Policies</h3>"
        $mdHTML += (ConvertTo-HTMLTable -Data $mdPolicies -NoDataMessage "Keine Policies konfiguriert.")

        $asOrg = Get-ActiveSyncOrganizationSettings -ErrorAction SilentlyContinue |
            Select-Object DefaultAccessLevel, UserMailInsert

        if ($asOrg) {
            $mdHTML += "<h3>ActiveSync Organisationseinstellungen</h3>"
            $mdHTML += (ConvertTo-HTMLTable -Data @($asOrg))
        }

        $mdStats = Get-MobileDevice -ResultSize 100 -ErrorAction SilentlyContinue |
            Group-Object DeviceType |
            Select-Object @{N='Gerätetyp';E={$_.Name}}, @{N='Anzahl';E={$_.Count}} |
            Sort-Object Anzahl -Descending

        if ($mdStats) {
            $mdHTML += "<h3>Mobile Geräte nach Typ (max. 100)</h3>"
            $mdHTML += (ConvertTo-HTMLTable -Data $mdStats)
        }
    }
    catch {
        Write-Log -Message "Fehler bei Mobile Device: $_" -Level "ERROR"
        $mdHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Mobile Device Konfiguration" -Content $mdHTML
}

# ---------------------------------------------------------------
# 19. ZERTIFIKATE
# ---------------------------------------------------------------
function Get-CertificateInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Exchange-Zertifikate inkl. Bindungen und Ablauf-Ampel.
    #>
    Write-Log -Message "=== Sammle Zertifikat-Informationen ===" -Level "INFO"

    $certHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Zertifikate für Server: $server" -Level "INFO"

            $certs = Get-ExchangeCertificate -Server $server -ErrorAction Stop |
                Select-Object @{N='Server';E={$server}},
                    Thumbprint,
                    @{N='Subject';E={$_.Subject}},
                    @{N='FriendlyName';E={$_.FriendlyName}},
                    @{N='CertificateDomains';E={$_.CertificateDomains -join ", "}},
                    @{N='Services';E={$_.Services}},
                    @{N='NotBefore';E={$_.NotBefore.ToString("dd.MM.yyyy")}},
                    @{N='NotAfter';E={$_.NotAfter.ToString("dd.MM.yyyy")}},
                    @{N='TageVerbleibend';E={[math]::Round(($_.NotAfter - (Get-Date)).TotalDays, 0)}},
                    @{N='Status';E={
                        $daysLeft = ($_.NotAfter - (Get-Date)).TotalDays
                        if ($daysLeft -lt 0) { "ABGELAUFEN!" }
                        elseif ($daysLeft -lt $script:WarningCertDaysExpiry) { "Läuft bald ab!" }
                        else { "OK" }
                    }},
                    @{N='SelfSigned';E={$_.IsSelfSigned}},
                    @{N='RootCAType';E={$_.RootCAType}},
                    @{N='PublicKeySize';E={$_.PublicKeySize}},
                    Issuer

            $certHTML += "<h3 class='server-break'>Server: $server</h3>"
            $certHTML += (ConvertTo-HTMLTable -Data $certs -NoDataMessage "Keine Zertifikate gefunden.")
        }
        catch {
            Write-Log -Message "Fehler bei Zertifikaten für ${server}: $_" -Level "ERROR"
            $certHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Zertifikate" -Content $certHTML
}

# ---------------------------------------------------------------
# 20. TRANSPORT-REGELN
# ---------------------------------------------------------------
function Get-TransportRuleInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Transport-Regeln.
    #>
    Write-Log -Message "=== Sammle Transport-Regeln ===" -Level "INFO"

    $trHTML = ""

    try {
        $rules = Get-TransportRule -ErrorAction Stop |
            Select-Object Name, State, Priority, Mode,
                @{N='Beschreibung';E={$_.Description}},
                @{N='Bedingungen';E={
                    $cond = @()
                    if ($_.From) { $cond += "Von: $($_.From -join ', ')" }
                    if ($_.SentTo) { $cond += "An: $($_.SentTo -join ', ')" }
                    if ($_.SubjectContainsWords) { $cond += "Betreff: $($_.SubjectContainsWords -join ', ')" }
                    if ($_.FromMemberOf) { $cond += "Von Mitglied: $($_.FromMemberOf -join ', ')" }
                    if ($cond.Count -eq 0) { "Siehe Detail-Regel" } else { $cond -join "; " }
                }},
                @{N='Aktionen';E={
                    $act = @()
                    if ($_.AddToRecipients) { $act += "Weiterleiten: $($_.AddToRecipients -join ', ')" }
                    if ($_.PrependSubject) { $act += "Prefix: $($_.PrependSubject)" }
                    if ($_.RejectMessageReasonText) { $act += "Ablehnen: $($_.RejectMessageReasonText)" }
                    if ($_.ApplyHtmlDisclaimerText) { $act += "Disclaimer" }
                    if ($act.Count -eq 0) { "Siehe Detail-Regel" } else { $act -join "; " }
                }},
                WhenChanged

        $trHTML += (ConvertTo-HTMLTable -Data $rules -NoDataMessage "Keine Transport-Regeln konfiguriert.")
    }
    catch {
        Write-Log -Message "Fehler bei Transport-Regeln: $_" -Level "ERROR"
        $trHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Transport-Regeln (Mailflow Rules)" -Content $trHTML
}

# ---------------------------------------------------------------
# 21. TRANSPORTKOMPONENTEN - PHYSISCHE SPEICHERORTE
# ---------------------------------------------------------------
function Get-TransportComponentStorageInfo {
    <#
    .SYNOPSIS
        Dokumentiert die physischen Speicherorte der Transportkomponenten.
        Umfasst: Queue-Datenbank, Queue-Logs, Message-Tracking-Logs, SMTP-Protokolllogs, Safety-Net.
    #>
    Write-Log -Message "=== Sammle Transportkomponenten-Speicherorte ===" -Level "INFO"

    $tcHTML = ""

    try {
        # Transport-Konfiguration abrufen
        $transportConfig = Get-TransportConfig -ErrorAction Stop
        $transportServers = Get-TransportServer -ErrorAction Stop

        # === Queue-Datenbank und Queue-Logs ===
        $queueStorage = @()
        foreach ($server in $transportServers) {
            try {
                $queueStorage += [PSCustomObject]@{
                    Server = $server.Identity
                    'Queue-Datenbank Pfad' = $server.QueueLogPath
                    'Queue-Logs Pfad' = $server.QueueLogPath
                    'Queue-Datenbank Größe' = $server.QueueDatabasePath
                }
            }
            catch {
                Write-Log -Message "Fehler beim Abrufen der Queue-Pfade für $($server.Identity): $_" -Level "WARNING"
            }
        }

        if ($queueStorage.Count -gt 0) {
            $tcHTML += "<h3>Queue-Datenbank und Queue-Logs</h3>"
            $tcHTML += (ConvertTo-HTMLTable -Data $queueStorage -NoDataMessage "Keine Queue-Informationen verfügbar.")
        }

        # === Message-Tracking-Logs ===
        $messageTrackingStorage = @()
        foreach ($server in $transportServers) {
            try {
                $messageTrackingStorage += [PSCustomObject]@{
                    Server = $server.Identity
                    'Message-Tracking-Log Pfad' = $server.MessageTrackingLogPath
                    'Maximale Größe (GB)' = $server.MessageTrackingLogMaxAge
                    'Aufbewahrungsdauer (Tage)' = $server.MessageTrackingLogMaxDirectorySize
                }
            }
            catch {
                Write-Log -Message "Fehler beim Abrufen der Message-Tracking-Pfade für $($server.Identity): $_" -Level "WARNING"
            }
        }

        if ($messageTrackingStorage.Count -gt 0) {
            $tcHTML += "<h3>Message-Tracking-Logs</h3>"
            $tcHTML += (ConvertTo-HTMLTable -Data $messageTrackingStorage -NoDataMessage "Keine Message-Tracking-Informationen verfügbar.")
        }

        # === SMTP-Protokolllogs für Empfangsconnectoren ===
        $receiveConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue
        $smtpReceiveLogs = @()
        if ($receiveConnectors) {
            foreach ($connector in $receiveConnectors) {
                try {
                    $smtpReceiveLogs += [PSCustomObject]@{
                        Connector = $connector.Identity
                        Typ = "Empfänger"
                        'SMTP-Log Pfad' = if ($connector.ProtocolLoggingLevel -eq "Verbose") { "Aktiviert - Pfad systemabhängig" } else { "Nicht aktiviert / $($connector.ProtocolLoggingLevel)" }
                        'Log-Level' = $connector.ProtocolLoggingLevel
                    }
                }
                catch {
                    Write-Log -Message "Fehler beim Abrufen des SMTP-Logs für $($connector.Identity): $_" -Level "WARNING"
                }
            }
        }

        # === SMTP-Protokolllogs für Sendeconnectoren ===
        $sendConnectors = Get-SendConnector -ErrorAction SilentlyContinue
        if ($sendConnectors) {
            foreach ($connector in $sendConnectors) {
                try {
                    $smtpReceiveLogs += [PSCustomObject]@{
                        Connector = $connector.Identity
                        Typ = "Sender"
                        'SMTP-Log Pfad' = if ($connector.ProtocolLoggingLevel -eq "Verbose") { "Aktiviert - Pfad systemabhängig" } else { "Nicht aktiviert / $($connector.ProtocolLoggingLevel)" }
                        'Log-Level' = $connector.ProtocolLoggingLevel
                    }
                }
                catch {
                    Write-Log -Message "Fehler beim Abrufen des SMTP-Logs für $($connector.Identity): $_" -Level "WARNING"
                }
            }
        }

        if ($smtpReceiveLogs.Count -gt 0) {
            $tcHTML += "<h3>SMTP-Protokolllogs (Sende- und Empfangsconnectoren)</h3>"
            $tcHTML += (ConvertTo-HTMLTable -Data $smtpReceiveLogs -NoDataMessage "Keine SMTP-Log-Informationen verfügbar.")
        }

        # === Safety-Net Konfiguration ===
        $safetyNetInfo = @()
        try {
            $safetyNetInfo += [PSCustomObject]@{
                Parameter = "Safety-Net Aktiviert"
                Wert = $transportConfig.SafetyNetEnabled
            }
            $safetyNetInfo += [PSCustomObject]@{
                Parameter = "Safety-Net Hold Time (Stunden)"
                Wert = if ($transportConfig.SafetyNetHoldTime) { $transportConfig.SafetyNetHoldTime.TotalHours } else { "Standard" }
            }
            $safetyNetInfo += [PSCustomObject]@{
                Parameter = "Safety-Net Quota (GB)"
                Wert = if ($transportConfig.SafetyNetQuotaInGigabytes) { $transportConfig.SafetyNetQuotaInGigabytes } else { "Standard" }
            }
            $safetyNetInfo += [PSCustomObject]@{
                Parameter = "Maximum Recipients For Safety Net"
                Wert = $transportConfig.MaximumRecipientsForSafetyNet
            }

            $tcHTML += "<h3>Safety-Net Konfiguration</h3>"
            $tcHTML += (ConvertTo-HTMLTable -Data $safetyNetInfo -NoDataMessage "Keine Safety-Net-Informationen verfügbar.")
        }
        catch {
            Write-Log -Message "Fehler beim Abrufen der Safety-Net-Konfiguration: $_" -Level "WARNING"
        }

        # === Zusammenfassung der Standard-Speicherorte ===
        $tcHTML += "<h3>Standard-Speicherorte für Transportkomponenten</h3>"
        $standardPaths = @()
        $standardPaths += [PSCustomObject]@{
            Komponente = "Queue-Datenbank"
            'Standardpfad' = "%ExchangeInstallPath%Bin\Queue"
            Beschreibung = "Enthält die Transportdatenbank mit ausstehenden Nachrichten"
        }
        $standardPaths += [PSCustomObject]@{
            Komponente = "Queue-Logs"
            'Standardpfad' = "%ExchangeInstallPath%Bin\Queue"
            Beschreibung = "Transaktionslogs für die Queue-Datenbank"
        }
        $standardPaths += [PSCustomObject]@{
            Komponente = "Message-Tracking-Logs"
            'Standardpfad' = "%ExchangeInstallPath%Logging\MessageTracking"
            Beschreibung = "Protokolliert alle Transportereignisse für Audit und Compliance"
        }
        $standardPaths += [PSCustomObject]@{
            Komponente = "SMTP-Empfangs-Protokolllogs"
            'Standardpfad' = "%ExchangeInstallPath%Logging\Protocols\SmtpReceive"
            Beschreibung = "SMTP-Protokolllogs für eingehende Verbindungen"
        }
        $standardPaths += [PSCustomObject]@{
            Komponente = "SMTP-Send-Protokolllogs"
            'Standardpfad' = "%ExchangeInstallPath%Logging\Protocols\SmtpSend"
            Beschreibung = "SMTP-Protokolllogs für ausgehende Verbindungen"
        }
        $standardPaths += [PSCustomObject]@{
            Komponente = "Safety-Net Daten"
            'Standardpfad' = "%ExchangeInstallPath%TransportRoles\Data\Queue"
            Beschreibung = "Shadow-Redundanz Daten für Nachrichtenwiederherstellung"
        }

        $tcHTML += (ConvertTo-HTMLTable -Data $standardPaths -NoDataMessage "Keine Standard-Pfade verfügbar.")
    }
    catch {
        Write-Log -Message "Fehler bei Transportkomponenten-Speicherorten: $_" -Level "ERROR"
        $tcHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Transportkomponenten - Physische Speicherorte" -Content $tcHTML
}

# ---------------------------------------------------------------
# 22. MAILBOX-STATISTIKEN
# ---------------------------------------------------------------
function Get-MailboxStatisticsOverview {
    <#
    .SYNOPSIS
        Erstellt eine Empfänger-Übersicht und Mailboxen-pro-DB Statistik.
    #>
    Write-Log -Message "=== Sammle Mailbox-Statistiken ===" -Level "INFO"

    $mbxHTML = ""

    try {
        $recipientStats = @()
        $recipientStats += [PSCustomObject]@{Typ="UserMailbox";         Anzahl=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="SharedMailbox";       Anzahl=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="RoomMailbox";         Anzahl=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="EquipmentMailbox";    Anzahl=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails EquipmentMailbox -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="DiscoveryMailbox";    Anzahl=(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails DiscoveryMailbox -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="MailContacts";        Anzahl=(Get-MailContact -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="MailUser";            Anzahl=(Get-MailUser -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="Verteilergruppen";    Anzahl=(Get-DistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count}
        $recipientStats += [PSCustomObject]@{Typ="Dynamische VG";       Anzahl=(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count}

        $mbxHTML += "<h3>Empfänger-Übersicht</h3>"
        $mbxHTML += (ConvertTo-HTMLTable -Data $recipientStats)

        $archiveCount = (Get-Mailbox -ResultSize Unlimited -Archive -ErrorAction SilentlyContinue | Measure-Object).Count
        $mbxHTML += "<p><strong>Mailboxen mit Archiv:</strong> $archiveCount</p>"

        # Mailboxen pro DB
        $databases = Get-MailboxDatabase -ErrorAction SilentlyContinue
        $dbStats = @()
        foreach ($db in $databases) {
            try {
                $mbxCount = (Get-Mailbox -Database $db.Name -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count
                $dbStats += [PSCustomObject]@{
                    Datenbank = $db.Name
                    Mailboxen = $mbxCount
                    Server    = $db.Server
                }
            }
            catch {
                Write-Log -Message "DB-Statistik Fehler für $($db.Name): $_" -Level "WARNING"
            }
        }

        $mbxHTML += "<h3>Mailboxen pro Datenbank</h3>"
        $mbxHTML += (ConvertTo-HTMLTable -Data $dbStats)
    }
    catch {
        Write-Log -Message "Fehler bei Mailbox-Statistiken: $_" -Level "ERROR"
        $mbxHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Mailbox-Statistiken & Empfänger-Übersicht" -Content $mbxHTML
}

# ---------------------------------------------------------------
# 23. THROTTLING POLICIES
# ---------------------------------------------------------------
function Get-ThrottlingPolicyInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Throttling Policies.
    #>
    Write-Log -Message "=== Sammle Throttling Policies ===" -Level "INFO"

    $tpHTML = ""

    try {
        $policies = Get-ThrottlingPolicy -ErrorAction Stop |
            Select-Object Name, IsDefault,
                @{N='EWSMaxConcurrency';E={$_.EwsMaxConcurrency}},
                @{N='EWSMaxSubscriptions';E={$_.EwsMaxSubscriptions}},
                @{N='OWAMaxConcurrency';E={$_.OwaMaxConcurrency}},
                @{N='RCAMaxConcurrency';E={$_.RcaMaxConcurrency}},
                @{N='MessageRateLimit';E={$_.MessageRateLimit}},
                @{N='RecipientRateLimit';E={$_.RecipientRateLimit}}

        $tpHTML += (ConvertTo-HTMLTable -Data $policies -NoDataMessage "Keine benutzerdefinierten Throttling Policies.")
    }
    catch {
        Write-Log -Message "Fehler bei Throttling Policies: $_" -Level "ERROR"
        $tpHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Throttling Policies" -Content $tpHTML
}

# ---------------------------------------------------------------
# 24. RETENTION / COMPLIANCE
# ---------------------------------------------------------------
function Get-RetentionPolicyInfo {
    <#
    .SYNOPSIS
        Dokumentiert Retention Policies, Tags und Journal Rules.
    #>
    Write-Log -Message "=== Sammle Retention Policies ===" -Level "INFO"

    $retHTML = ""

    try {
        $retPolicies = Get-RetentionPolicy -ErrorAction SilentlyContinue |
            Select-Object Name,
                @{N='RetentionPolicyTagLinks';E={$_.RetentionPolicyTagLinks -join ", "}},
                @{N='IsDefault';E={$_.IsDefault}},
                WhenCreated

        $retHTML += "<h3>Retention Policies</h3>"
        $retHTML += (ConvertTo-HTMLTable -Data $retPolicies -NoDataMessage "Keine Retention Policies.")

        $retTags = Get-RetentionPolicyTag -ErrorAction SilentlyContinue |
            Select-Object Name, Type,
                @{N='AgeLimitForRetention';E={$_.AgeLimitForRetention}},
                @{N='RetentionAction';E={$_.RetentionAction}},
                @{N='RetentionEnabled';E={$_.RetentionEnabled}},
                @{N='MessageClass';E={$_.MessageClass}}

        $retHTML += "<h3>Retention Tags</h3>"
        $retHTML += (ConvertTo-HTMLTable -Data $retTags -NoDataMessage "Keine Retention Tags.")

        $journalRules = Get-JournalRule -ErrorAction SilentlyContinue |
            Select-Object Name, Recipient, JournalEmailAddress, Scope, Enabled

        $retHTML += "<h3>Journal Rules</h3>"
        $retHTML += (ConvertTo-HTMLTable -Data $journalRules -NoDataMessage "Keine Journal Rules.")
    }
    catch {
        Write-Log -Message "Fehler bei Retention Policies: $_" -Level "ERROR"
        $retHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Retention Policies & Journal Rules" -Content $retHTML
}

# ---------------------------------------------------------------
# 25. RBAC / ADMIN-ROLLEN
# ---------------------------------------------------------------
function Get-RBACInfo {
    <#
    .SYNOPSIS
        Dokumentiert RBAC-Rollengruppen und deren Mitglieder.
    #>
    Write-Log -Message "=== Sammle RBAC-Informationen ===" -Level "INFO"

    $rbacHTML = ""

    try {
        $roleGroups = Get-RoleGroup -ErrorAction Stop | ForEach-Object {
            $members = try {
                (Get-RoleGroupMember -Identity $_.Name -ErrorAction SilentlyContinue).Name -join ", "
            }
            catch { "Fehler beim Abruf" }

            [PSCustomObject]@{
                Name         = $_.Name
                Beschreibung = $_.Description
                Mitglieder   = if ($members) { $members } else { "Keine Mitglieder" }
                Rollen       = ($_.Roles -join ", ")
            }
        }

        $rbacHTML += (ConvertTo-HTMLTable -Data $roleGroups -NoDataMessage "Keine Rollengruppen gefunden.")
    }
    catch {
        Write-Log -Message "Fehler bei RBAC: $_" -Level "ERROR"
        $rbacHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "RBAC - Administratorenrollen & Mitglieder" -Content $rbacHTML
}

# ---------------------------------------------------------------
# 26. EVENT LOGS (CIM/DCOM)
# ---------------------------------------------------------------
function Get-EventLogInfo {
    <#
    .SYNOPSIS
        Sammelt Event Logs über CIM-Sessions (kein WinRM nötig).
    #>
    Write-Log -Message "=== Sammle Event Log Informationen (CIM) ===" -Level "INFO"

    $elHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            # CIM-Session aufbauen
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $elHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Keine CIM-Verbindung möglich.</p>"
                continue
            }

            $cutoffDate = (Get-Date).AddDays(-7).ToString("yyyyMMddHHmmss.000000+000")

            # Application Log Errors
            $appErrors = Get-CimInstance -CimSession $cimSession -ClassName Win32_NTLogEvent `
                -Filter "LogFile='Application' AND EventType<=2 AND TimeGenerated>='$cutoffDate'" `
                -ErrorAction SilentlyContinue |
                Select-Object -First 25 @{N='Zeitpunkt';E={$_.TimeGenerated}},
                    @{N='Log';E={$_.LogFile}},
                    @{N='Level';E={if($_.EventType -eq 1){"Error"}else{"Warning"}}},
                    @{N='EventID';E={$_.EventCode}},
                    @{N='Quelle';E={$_.SourceName}},
                    @{N='Nachricht';E={
                        if ($_.Message -and $_.Message.Length -gt 200) { $_.Message.Substring(0,200) + "..." }
                        elseif ($_.Message) { $_.Message }
                        else { "N/A" }
                    }}

            # System Log Errors
            $sysErrors = Get-CimInstance -CimSession $cimSession -ClassName Win32_NTLogEvent `
                -Filter "LogFile='System' AND EventType<=2 AND TimeGenerated>='$cutoffDate'" `
                -ErrorAction SilentlyContinue |
                Select-Object -First 25 @{N='Zeitpunkt';E={$_.TimeGenerated}},
                    @{N='Log';E={$_.LogFile}},
                    @{N='Level';E={if($_.EventType -eq 1){"Error"}else{"Warning"}}},
                    @{N='EventID';E={$_.EventCode}},
                    @{N='Quelle';E={$_.SourceName}},
                    @{N='Nachricht';E={
                        if ($_.Message -and $_.Message.Length -gt 200) { $_.Message.Substring(0,200) + "..." }
                        elseif ($_.Message) { $_.Message }
                        else { "N/A" }
                    }}

            $allEvents = @()
            if ($appErrors) { $allEvents += $appErrors }
            if ($sysErrors) { $allEvents += $sysErrors }

            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue

            $elHTML += "<h3 class='server-break'>Server: $server (letzte 7 Tage, max. 50 Fehler/Warnungen)</h3>"
            $elHTML += (ConvertTo-HTMLTable -Data $allEvents -NoDataMessage "Keine kritischen Events - sehr gut!")
        }
        catch {
            Write-Log -Message "Fehler bei Event Logs für ${server}: $_" -Level "WARNING"
            $elHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Event Logs (letzte 7 Tage - Fehler/Kritisch)" -Content $elHTML
}

# ---------------------------------------------------------------
# 27. SICHERHEIT & AUTHENTIFIZIERUNG
# ---------------------------------------------------------------
function Get-SecurityAndAuthInfo {
    <#
    .SYNOPSIS
        Dokumentiert Sicherheitseinstellungen: TLS, OAuth, SMTP-Auth, Audit-Logging.
    #>
    Write-Log -Message "=== Sammle Sicherheit & Authentifizierung ===" -Level "INFO"

    $secHTML = ""

    # --- TLS-Einstellungen auf Transport-Ebene ---
    try {
        $transportConfig = Get-TransportConfig -ErrorAction SilentlyContinue
        if ($transportConfig) {
            $tlsInfo = [PSCustomObject]@{
                "TLS Sende Domain Secure List"      = ($transportConfig.TLSSendDomainSecureList -join ', ')
                "TLS Empfang Domain Secure List"    = ($transportConfig.TLSReceiveDomainSecureList -join ', ')
                "External Postmaster Address"       = $transportConfig.ExternalPostmasterAddress
                "Internal SMTP Servers"             = ($transportConfig.InternalSMTPServers -join ', ')
            }
            $secHTML += "<h3>Transport TLS-Konfiguration</h3>"
            $secHTML += (ConvertTo-HTMLTable -Data @($tlsInfo))
        }
    }
    catch {
        Write-Log -Message "TransportConfig TLS-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- SMTP Auth Einstellungen pro Receive Connector ---
    try {
        $rcConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue
        if ($rcConnectors) {
            $authData = foreach ($rc in $rcConnectors) {
                [PSCustomObject]@{
                    Name                    = $rc.Name
                    Server                  = $rc.Server
                    "Auth Mechanisms"       = ($rc.AuthMechanism -join ', ')
                    "Permission Groups"     = ($rc.PermissionGroups -join ', ')
                    RequireTLS              = $rc.RequireTLS
                    "SMTP Auth Deaktiviert" = if ($rc.AuthMechanism -notmatch "Login|BasicAuth") { "Ja (sicher)" } else { "Nein (Auth aktiv)" }
                }
            }
            $secHTML += "<h3>SMTP-Authentifizierung je Receive Connector</h3>"
            $secHTML += (ConvertTo-HTMLTable -Data $authData)
        }
    }
    catch {
        Write-Log -Message "ReceiveConnector Auth-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Audit Logging ---
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($orgConfig) {
            $auditInfo = [PSCustomObject]@{
                "Mailbox Audit Log aktiviert"      = $orgConfig.AuditDisabled -eq $false
                "Admin Audit Log aktiviert"        = $orgConfig.AdminAuditLogEnabled
                "Admin Audit Log Cmdlets"          = $orgConfig.AdminAuditLogCmdlets -join ', '
                "Admin Audit Log Age Limit"        = $orgConfig.AdminAuditLogAgeLimit
            }
            $secHTML += "<h3>Audit Logging Konfiguration</h3>"
            $secHTML += (ConvertTo-HTMLTable -Data @($auditInfo))
        }
    }
    catch {
        Write-Log -Message "Audit Logging Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- OAuth / Modern Auth Status ---
    try {
        $oauthEnabled = $orgConfig.OAuth2ClientProfileEnabled
        $secHTML += "<h3>Moderne Authentifizierung (OAuth)</h3>"
        $secHTML += "<p><strong>OAuth2 Client Profile Enabled:</strong> $oauthEnabled</p>"
        if (-not $oauthEnabled) {
            $secHTML += "<p class='error'>⚠️ OAuth ist NICHT aktiviert. Empfohlen für Hybrid und Sicherheit!</p>"
        }
    }
    catch {
        Write-Log -Message "OAuth-Status Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    New-HTMLSection -Title "Sicherheit & Authentifizierung" -Content $secHTML
}

# ---------------------------------------------------------------
# 28. ANTI-SPAM / ANTI-MALWARE KONFIGURATION
# ---------------------------------------------------------------
function Get-AntiSpamMalwareInfo {
    <#
    .SYNOPSIS
        Dokumentiert Anti-Spam und Anti-Malware Einstellungen.
    #>
    Write-Log -Message "=== Sammle Anti-Spam/Anti-Malware Konfiguration ===" -Level "INFO"

    $asmHTML = ""

    # --- Malware Filter Policy ---
    try {
        $malwarePolicies = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue
        if ($malwarePolicies) {
            $mfData = foreach ($mp in $malwarePolicies) {
                [PSCustomObject]@{
                    Name                           = $mp.Name
                    "Action bei Erkennung"         = $mp.Action
                    "Admin Benachrichtigung"       = $mp.EnableInternalSenderAdminNotifications
                    "Admin E-Mail"                 = $mp.InternalSenderAdminAddress
                    "Custom Alert Text"            = if ($mp.CustomAlertText) { "Ja" } else { "Nein" }
                    "File Filter aktiviert"        = $mp.EnableFileFilter
                    "Gesperrte Dateitypen"         = ($mp.FileTypes -join ', ')
                    "ZAP aktiviert"                = $mp.ZapEnabled
                }
            }
            $asmHTML += "<h3>Malware Filter Policies</h3>"
            $asmHTML += (ConvertTo-HTMLTable -Data $mfData)
        }
    }
    catch {
        Write-Log -Message "MalwareFilterPolicy-Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $asmHTML += "<p class='no-data'>Malware Filter Policies nicht verfügbar.</p>"
    }

    # --- Content Filter (Anti-Spam) ---
    try {
        $contentFilter = Get-ContentFilterConfig -ErrorAction SilentlyContinue
        if ($contentFilter) {
            $cfInfo = [PSCustomObject]@{
                "Enabled"                    = $contentFilter.Enabled
                "SCL Delete Threshold"       = $contentFilter.SCLDeleteThreshold
                "SCL Reject Threshold"       = $contentFilter.SCLRejectThreshold
                "SCL Quarantine Threshold"   = $contentFilter.SCLQuarantineThreshold
                "SCL Junk Threshold"         = $contentFilter.SCLJunkThreshold
                "Quarantine Mailbox"         = $contentFilter.QuarantineMailbox
            }
            $asmHTML += "<h3>Content Filter (Anti-Spam) Konfiguration</h3>"
            $asmHTML += (ConvertTo-HTMLTable -Data @($cfInfo))
        }
    }
    catch {
        Write-Log -Message "ContentFilterConfig nicht verfügbar (ggf. nur auf Edge): $_" -Level "WARNING"
        $asmHTML += "<p class='no-data'>Content Filter Config nicht verfügbar (nur auf Edge Transport / Hub mit aktiviertem Agent).</p>"
    }

    # --- Sender Filter ---
    try {
        $senderFilter = Get-SenderFilterConfig -ErrorAction SilentlyContinue
        if ($senderFilter) {
            $sfInfo = [PSCustomObject]@{
                "Enabled"             = $senderFilter.Enabled
                "Blocked Senders"     = ($senderFilter.BlockedSenders -join ', ')
                "Blocked Domains"     = ($senderFilter.BlockedDomains -join ', ')
                "BlankSender Reject"  = $senderFilter.BlankSenderBlockingEnabled
            }
            $asmHTML += "<h3>Sender Filter Konfiguration</h3>"
            $asmHTML += (ConvertTo-HTMLTable -Data @($sfInfo))
        }
    }
    catch {
        Write-Log -Message "SenderFilterConfig nicht verfügbar: $_" -Level "WARNING"
    }

    # --- Connection Filter (IP Block/Allow) ---
    try {
        $connFilter = Get-IPBlockListConfig -ErrorAction SilentlyContinue
        if ($connFilter) {
            $asmHTML += "<h3>IP Block List Konfiguration</h3>"
            $ipBlocked = Get-IPBlockListEntry -ErrorAction SilentlyContinue
            if ($ipBlocked) {
                $asmHTML += (ConvertTo-HTMLTable -Data ($ipBlocked | Select-Object IPRange, ExpirationTime, Comment))
            }
            else {
                $asmHTML += "<p class='no-data'>Keine IP Block List Einträge.</p>"
            }
        }
    }
    catch {
        Write-Log -Message "IPBlockList nicht verfügbar: $_" -Level "WARNING"
    }

    if (-not $asmHTML) {
        $asmHTML = "<p class='no-data'>Anti-Spam/Anti-Malware Agents sind möglicherweise nicht auf Mailbox-Servern aktiviert. Prüfen Sie Edge Transport Server.</p>"
    }

    New-HTMLSection -Title "Anti-Spam / Anti-Malware Konfiguration" -Content $asmHTML
}

# ---------------------------------------------------------------
# 29. COMPLIANCE & DLP
# ---------------------------------------------------------------
function Get-ComplianceInfo {
    <#
    .SYNOPSIS
        Dokumentiert Compliance-Einstellungen: DLP, Litigation Hold, In-Place Hold, eDiscovery.
    #>
    Write-Log -Message "=== Sammle Compliance & DLP Informationen ===" -Level "INFO"

    $compHTML = ""

    # --- DLP Policies ---
    try {
        $dlpPolicies = Get-DlpPolicy -ErrorAction SilentlyContinue
        if ($dlpPolicies) {
            $dlpData = foreach ($dlp in $dlpPolicies) {
                [PSCustomObject]@{
                    Name        = $dlp.Name
                    State       = $dlp.State
                    Mode        = $dlp.Mode
                    Beschreibung = $dlp.Description
                }
            }
            $compHTML += "<h3>DLP-Richtlinien (Data Loss Prevention)</h3>"
            $compHTML += (ConvertTo-HTMLTable -Data $dlpData)
        }
        else {
            $compHTML += "<h3>DLP-Richtlinien</h3>"
            $compHTML += "<p class='no-data'>Keine DLP-Richtlinien konfiguriert.</p>"
        }
    }
    catch {
        Write-Log -Message "DLP Policy Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $compHTML += "<h3>DLP-Richtlinien</h3><p class='no-data'>DLP Cmdlet nicht verfügbar.</p>"
    }

    # --- Litigation Hold Übersicht ---
    try {
        $litHoldMailboxes = Get-Mailbox -ResultSize Unlimited -Filter { LitigationHoldEnabled -eq $true } -ErrorAction SilentlyContinue
        if ($litHoldMailboxes) {
            $compHTML += "<h3>Litigation Hold (aktive Postfächer: $($litHoldMailboxes.Count))</h3>"
            $lhData = $litHoldMailboxes | Select-Object -First 50 DisplayName,
                @{N='LitHold Datum';E={$_.LitigationHoldDate}},
                @{N='LitHold Dauer';E={$_.LitigationHoldDuration}},
                @{N='LitHold Besitzer';E={$_.LitigationHoldOwner}}
            $compHTML += (ConvertTo-HTMLTable -Data $lhData)
            if ($litHoldMailboxes.Count -gt 50) {
                $compHTML += "<p><em>Es werden nur die ersten 50 angezeigt. Gesamt: $($litHoldMailboxes.Count)</em></p>"
            }
        }
        else {
            $compHTML += "<h3>Litigation Hold</h3>"
            $compHTML += "<p class='no-data'>Keine Postfächer unter Litigation Hold.</p>"
        }
    }
    catch {
        Write-Log -Message "Litigation Hold Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- In-Place Hold / Compliance Search ---
    try {
        $inPlaceHolds = Get-MailboxSearch -ErrorAction SilentlyContinue
        if ($inPlaceHolds) {
            $iphData = foreach ($iph in $inPlaceHolds) {
                [PSCustomObject]@{
                    Name          = $iph.Name
                    Status        = $iph.Status
                    "In-Place Hold" = $iph.InPlaceHoldEnabled
                    "Start Date"  = $iph.StartDate
                    "End Date"    = $iph.EndDate
                    "Sources"     = $iph.SourceMailboxes.Count
                }
            }
            $compHTML += "<h3>In-Place Hold / Mailbox Searches</h3>"
            $compHTML += (ConvertTo-HTMLTable -Data $iphData)
        }
    }
    catch {
        Write-Log -Message "MailboxSearch Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    New-HTMLSection -Title "Compliance & Data Loss Prevention" -Content $compHTML
}

# ---------------------------------------------------------------
# 30. MAILBOX-QUOTAS
# ---------------------------------------------------------------
function Get-MailboxQuotaInfo {
    <#
    .SYNOPSIS
        Dokumentiert Mailbox-Quotas auf Datenbank- und individueller Ebene.
    #>
    Write-Log -Message "=== Sammle Mailbox-Quota Informationen ===" -Level "INFO"

    $quotaHTML = ""

    # --- Datenbank-Quotas ---
    try {
        $databases = Get-MailboxDatabase -ErrorAction Stop | Select-Object Name,
            @{N='ProhibitSend_GB';E={
                if ($_.ProhibitSendQuota -and $_.ProhibitSendQuota.ToString() -ne "Unlimited") {
                    [math]::Round($_.ProhibitSendQuota.Value.ToBytes() / 1GB, 2)
                } else { "Unlimited" }
            }},
            @{N='ProhibitSendReceive_GB';E={
                if ($_.ProhibitSendReceiveQuota -and $_.ProhibitSendReceiveQuota.ToString() -ne "Unlimited") {
                    [math]::Round($_.ProhibitSendReceiveQuota.Value.ToBytes() / 1GB, 2)
                } else { "Unlimited" }
            }},
            @{N='IssueWarning_GB';E={
                if ($_.IssueWarningQuota -and $_.IssueWarningQuota.ToString() -ne "Unlimited") {
                    [math]::Round($_.IssueWarningQuota.Value.ToBytes() / 1GB, 2)
                } else { "Unlimited" }
            }},
            @{N='RecoverableItems_GB';E={
                if ($_.RecoverableItemsQuota -and $_.RecoverableItemsQuota.ToString() -ne "Unlimited") {
                    [math]::Round($_.RecoverableItemsQuota.Value.ToBytes() / 1GB, 2)
                } else { "Unlimited" }
            }},
            @{N='DeletedItemRetention_Tage';E={$_.DeletedItemRetention.Days}}

        $quotaHTML += "<h3>Datenbank-Quotas (Standard-Limits)</h3>"
        $quotaHTML += (ConvertTo-HTMLTable -Data $databases)
    }
    catch {
        Write-Log -Message "Datenbank-Quota Abfrage fehlgeschlagen: $_" -Level "ERROR"
        $quotaHTML += "<p class='error'>Fehler: $_</p>"
    }

    # --- Individuelle Quota-Overrides (nur Mailboxen mit eigenen Limits) ---
    try {
        $customQuotaMbx = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Where-Object { $_.UseDatabaseQuotaDefaults -eq $false } |
            Select-Object -First 100 DisplayName,
                @{N='ProhibitSend';E={$_.ProhibitSendQuota}},
                @{N='ProhibitSendReceive';E={$_.ProhibitSendReceiveQuota}},
                @{N='IssueWarning';E={$_.IssueWarningQuota}},
                Database

        if ($customQuotaMbx) {
            $quotaHTML += "<h3>Postfächer mit individuellen Quota-Overrides (max. 100)</h3>"
            $quotaHTML += "<p><em>Diese Postfächer verwenden NICHT die Datenbank-Standardwerte.</em></p>"
            $quotaHTML += (ConvertTo-HTMLTable -Data $customQuotaMbx)
        }
        else {
            $quotaHTML += "<h3>Individuelle Quota-Overrides</h3>"
            $quotaHTML += "<p class='no-data'>Alle Postfächer verwenden Datenbank-Standardwerte.</p>"
        }
    }
    catch {
        Write-Log -Message "Individuelle Quota Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    New-HTMLSection -Title "Mailbox-Quotas & Speicher-Limits" -Content $quotaHTML
}

# ---------------------------------------------------------------
# 31. SMTP-RELAY KONFIGURATION
# ---------------------------------------------------------------
function Get-SMTPRelayInfo {
    <#
    .SYNOPSIS
        Dokumentiert SMTP-Relay Einstellungen (Geräte/Applikationen die über Exchange relayen).
    #>
    Write-Log -Message "=== Sammle SMTP-Relay Informationen ===" -Level "INFO"

    $relayHTML = ""

    # --- Receive Connectors mit Anonymous Relay ---
    try {
        $relayConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue |
            Where-Object { $_.PermissionGroups -match "Anonymous" -or $_.Name -like "*Relay*" -or $_.Name -like "*Application*" }

        if ($relayConnectors) {
            $rcData = foreach ($rc in $relayConnectors) {
                [PSCustomObject]@{
                    Name              = $rc.Name
                    Server            = $rc.Server
                    Bindings          = ($rc.Bindings -join ', ')
                    "Remote IPs"      = ($rc.RemoteIPRanges -join ', ')
                    "Permission Groups" = ($rc.PermissionGroups -join ', ')
                    "Auth Mechanism"  = ($rc.AuthMechanism -join ', ')
                    RequireTLS        = $rc.RequireTLS
                    MaxMessageSize    = $rc.MaxMessageSize
                    Enabled           = $rc.Enabled
                }
            }
            $relayHTML += "<h3>Receive Connectors mit Relay-Berechtigung</h3>"
            $relayHTML += "<p><em>Diese Connectors erlauben anonymes oder Applikations-Relay. Prüfen Sie die Remote IP Ranges!</em></p>"
            $relayHTML += (ConvertTo-HTMLTable -Data $rcData)
        }
        else {
            $relayHTML += "<h3>SMTP-Relay Connectors</h3>"
            $relayHTML += "<p class='no-data'>Keine expliziten Relay-Connectors gefunden.</p>"
        }
    }
    catch {
        Write-Log -Message "Relay Connector Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Sicherheitshinweis ---
    $relayHTML += "<h3>Sicherheitshinweis SMTP-Relay</h3>"
    $relayHTML += "<p><em>Empfehlung: Dokumentieren Sie welche Geräte/Applikationen über welchen Connector relayen. "
    $relayHTML += "Beschränken Sie Remote IP Ranges auf das Minimum. Verwenden Sie TLS wo möglich.</em></p>"
    $relayHTML += "<table><thead><tr><th>Typische Relay-Quellen</th><th>Empfehlung</th></tr></thead><tbody>"
    $relayHTML += "<tr class='even'><td>Multifunktionsdrucker/Scanner</td><td>Eigener Connector, IP-Beschränkung</td></tr>"
    $relayHTML += "<tr class='odd'><td>ERP/CRM Systeme (SAP, Dynamics)</td><td>Authentifiziertes Relay oder eigener Connector</td></tr>"
    $relayHTML += "<tr class='even'><td>Monitoring (SCOM, Zabbix, PRTG)</td><td>Eigener Connector, IP-Beschränkung</td></tr>"
    $relayHTML += "<tr class='odd'><td>Ticketsysteme (OTRS, ServiceNow)</td><td>Authentifizierung + TLS erzwingen</td></tr>"
    $relayHTML += "<tr class='even'><td>Backup-Systeme (Veeam, Commvault)</td><td>IP-Beschränkung, Alert-Mailbox prüfen</td></tr>"
    $relayHTML += "</tbody></table>"

    New-HTMLSection -Title "SMTP-Relay Konfiguration" -Content $relayHTML
}

# ---------------------------------------------------------------
# 32. FIREWALL-REGELN
# ---------------------------------------------------------------
function Get-FirewallInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Windows-Firewall-Regeln der Exchange-Server sowie
        eine Referenztabelle der für Exchange benötigten Ports.
    #>
    Write-Log -Message "=== Sammle Firewall-Informationen ===" -Level "INFO"

    $fwHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Firewall-Regeln für Server: $server" -Level "INFO"

            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $fwHTML += "<h3 class='server-break'>Server: $server</h3>"
                $fwHTML += "<p class='error'>Keine Verbindung möglich. Firewall-Regeln nicht abrufbar.</p>"
                continue
            }

            # --- Firewall-Profile (Status) ---
            try {
                $fwProfiles = Get-NetFirewallProfile -CimSession $cimSession -ErrorAction Stop
                $profileData = foreach ($fp in $fwProfiles) {
                    [PSCustomObject]@{
                        Profil               = $fp.Name
                        Aktiviert            = $fp.Enabled
                        "Eingehend Standard" = $fp.DefaultInboundAction
                        "Ausgehend Standard" = $fp.DefaultOutboundAction
                        "Logging Datei"      = $fp.LogFileName
                    }
                }
                $fwHTML += "<h3 class='server-break'>Server: $server - Firewall-Profile</h3>"
                $fwHTML += (ConvertTo-HTMLTable -Data $profileData -NoDataMessage "Nicht verfügbar")
            }
            catch {
                Write-Log -Message "Firewall-Profile fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $fwHTML += "<h3 class='server-break'>Server: $server - Firewall-Profile</h3>"
                $fwHTML += "<p class='no-data'>Profile nicht abrufbar (Get-NetFirewallProfile benötigt WinRM/CIM v2).</p>"
            }

            # --- Aktive eingehende Allow-Regeln (Exchange-relevant) ---
            try {
                $fwRules = Get-NetFirewallRule -CimSession $cimSession -ErrorAction Stop |
                    Where-Object { $_.Enabled -eq $true -and $_.Direction -eq "Inbound" -and $_.Action -eq "Allow" -and
                        ($_.DisplayName -match "Exchange|SMTP|HTTP|IMAP|POP|MAPI|RPC|Autodiscover|MSExchange") }

                if ($fwRules) {
                    $ruleData = foreach ($rule in ($fwRules | Select-Object -First 60)) {
                        $portFilter = $rule | Get-NetFirewallPortFilter -CimSession $cimSession -ErrorAction SilentlyContinue
                        [PSCustomObject]@{
                            Regel        = $rule.DisplayName
                            Profil       = $rule.Profile
                            Protokoll    = $portFilter.Protocol
                            "Lokaler Port" = ($portFilter.LocalPort -join ', ')
                            Gruppe       = $rule.DisplayGroup
                        }
                    }
                    $fwHTML += "<h4>Exchange-relevante eingehende Allow-Regeln (max. 60)</h4>"
                    $fwHTML += (ConvertTo-HTMLTable -Data $ruleData -NoDataMessage "Keine relevanten Regeln gefunden")
                }
                else {
                    $fwHTML += "<h4>Exchange-relevante eingehende Allow-Regeln</h4>"
                    $fwHTML += "<p class='no-data'>Keine Exchange-bezogenen Firewall-Regeln gefunden.</p>"
                }
            }
            catch {
                Write-Log -Message "Firewall-Regeln fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $fwHTML += "<p class='no-data'>Detaillierte Firewall-Regeln nicht abrufbar (benötigt WinRM für CIMv2-Namespace).</p>"
            }

            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "Allgemeiner Fehler bei Firewall-Info für ${server}: $_" -Level "WARNING"
            $fwHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    # --- Referenztabelle: Benötigte Exchange-Ports ---
    $fwHTML += "<h3>Referenz: Benötigte Exchange-Ports (Firewall-Freigaben)</h3>"
    $fwHTML += "<table><thead><tr><th>Port</th><th>Protokoll</th><th>Dienst / Verwendung</th><th>Richtung</th></tr></thead><tbody>"
    $portRef = @(
        @("25",    "TCP/SMTP",   "Mail-Transport (intern & extern eingehend)", "Eingehend/Ausgehend"),
        @("587",   "TCP/SMTP",   "Client-SMTP-Submission (authentifiziert)",   "Eingehend"),
        @("443",   "TCP/HTTPS",  "OWA, ECP, EWS, ActiveSync, OAB, MAPI, Autodiscover, RPS", "Eingehend"),
        @("80",    "TCP/HTTP",   "Redirect zu HTTPS, Zertifikatsabruf (CRL)",  "Eingehend"),
        @("135",   "TCP/RPC",    "RPC Endpoint Mapper (AD, Management)",        "Intern"),
        @("143/993","TCP/IMAP",  "IMAP4 / IMAP4 über SSL (falls aktiviert)",   "Eingehend"),
        @("110/995","TCP/POP",   "POP3 / POP3 über SSL (falls aktiviert)",     "Eingehend"),
        @("3268/3269","TCP/LDAP-GC","Globaler Katalog (AD-Anbindung)",         "Intern"),
        @("389/636","TCP/LDAP",  "LDAP / LDAPS (Domänencontroller)",           "Intern"),
        @("64327","TCP",         "DAG Replikation (Standard, anpassbar)",      "Intern (DAG-Knoten)"),
        @("2525/475","TCP",      "Edge/Hub Transport (umgebungsabhängig)",     "Intern")
    )
    $i = 0
    foreach ($p in $portRef) {
        $cls = if ($i % 2 -eq 0) { "even" } else { "odd" }
        $fwHTML += "<tr class='$cls'><td>$($p[0])</td><td>$($p[1])</td><td>$($p[2])</td><td>$($p[3])</td></tr>"
        $i++
    }
    $fwHTML += "</tbody></table>"
    $fwHTML += "<p><em>Hinweis: DAG-Replikationsport kann mit Set-DatabaseAvailabilityGroup -ReplicationPort angepasst werden. "
    $fwHTML += "Dokumentieren Sie zusätzlich Load-Balancer-VIPs und externe Firewall-NAT-Regeln.</em></p>"

    New-HTMLSection -Title "Firewall-Regeln & Port-Anforderungen" -Content $fwHTML
}

# ---------------------------------------------------------------
# 33. LIZENZIERUNG
# ---------------------------------------------------------------
function Get-LicensingInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Exchange-Edition, Windows-Lizenzierung und liefert
        Anhaltspunkte zur CAL-Planung (Server- und Client-Lizenzen).
    #>
    Write-Log -Message "=== Sammle Lizenzierungs-Informationen ===" -Level "INFO"

    $licHTML = ""

    # --- Exchange Server Edition (Standard/Enterprise) ---
    try {
        $exServers = Get-ExchangeServer -ErrorAction Stop
        $exLicData = foreach ($ex in $exServers) {
            $dbCount = (Get-MailboxDatabase -Server $ex.Name -ErrorAction SilentlyContinue | Measure-Object).Count
            [PSCustomObject]@{
                Server              = $ex.Name
                Edition             = $ex.Edition
                Version             = $ex.AdminDisplayVersion
                "Anzahl Datenbanken" = $dbCount
                "DB-Limit (Edition)" = if ($ex.Edition -match "Standard") { "5 (Standard)" } elseif ($ex.Edition -match "Enterprise") { "100 (Enterprise)" } else { "Unbekannt" }
                "Limit-Warnung"     = if ($ex.Edition -match "Standard" -and $dbCount -ge 5) { "⚠️ Standard-Limit erreicht!" } else { "✅ OK" }
            }
        }
        $licHTML += "<h3>Exchange Server Editionen & Datenbank-Limits</h3>"
        $licHTML += (ConvertTo-HTMLTable -Data $exLicData)
    }
    catch {
        Write-Log -Message "Exchange Edition Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $licHTML += "<p class='error'>Exchange-Editionen nicht abrufbar: $_</p>"
    }

    # --- Postfach-Anzahl (Basis für CAL-Bedarf) ---
    try {
        $allMbx       = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
        $userMbx      = ($allMbx | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" }).Count
        $sharedMbx    = ($allMbx | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }).Count
        $roomMbx      = ($allMbx | Where-Object { $_.RecipientTypeDetails -eq "RoomMailbox" }).Count
        $equipMbx     = ($allMbx | Where-Object { $_.RecipientTypeDetails -eq "EquipmentMailbox" }).Count

        $calInfo = [PSCustomObject]@{
            "Benutzer-Postfächer (CAL-relevant)" = $userMbx
            "Shared Mailboxes (keine CAL)"       = $sharedMbx
            "Raum-Postfächer (keine CAL)"        = $roomMbx
            "Geräte-Postfächer (keine CAL)"      = $equipMbx
            "Geschätzter CAL-Bedarf (min.)"      = $userMbx
        }
        $licHTML += "<h3>CAL-Planung (Client Access Licenses)</h3>"
        $licHTML += (ConvertTo-HTMLTable -Data @($calInfo))
        $licHTML += "<p><em>Hinweis: CALs werden nicht technisch erzwungen (Vertrauensbasis). "
        $licHTML += "Standard-CAL deckt Basis-Postfach ab; Enterprise-CAL (additiv) wird für DLP, In-Place Hold/Archiv, erweiterte Compliance benötigt.</em></p>"
    }
    catch {
        Write-Log -Message "CAL-Postfachzählung fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Lizenzmodell-Hinweis Exchange 2019 vs. SE ---
    $licHTML += "<h3>Lizenzmodell-Hinweis (Edition: $($script:ExchangeEdition))</h3>"
    if ($script:ExchangeEdition -eq "SE") {
        $licHTML += "<p><strong>Exchange Server Subscription Edition (SE):</strong> Erfordert ein aktives Abonnement "
        $licHTML += "(Server-Lizenz mit Software Assurance bzw. Subscription). Es werden weiterhin Exchange Server-Lizenzen "
        $licHTML += "(Standard/Enterprise) und CALs benötigt. Prüfen Sie die Gültigkeit der Software Assurance.</p>"
    }
    else {
        $licHTML += "<p><strong>Exchange Server 2019:</strong> Klassisches Server+CAL-Modell. "
        $licHTML += "Beachten Sie das End-of-Support (Oktober 2025) - Migration auf Exchange SE empfohlen.</p>"
    }

    # --- Windows Server Lizenzierung (Aktivierungsstatus) ---
    foreach ($server in $ExchangeServers) {
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if ($cimSession) {
                $winLic = Get-CimInstance -CimSession $cimSession -ClassName SoftwareLicensingProduct `
                    -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus=1" -ErrorAction SilentlyContinue |
                    Select-Object -First 1
                $osRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue

                $winLicInfo = [PSCustomObject]@{
                    Server              = $server
                    Betriebssystem      = $osRaw.Caption
                    "Lizenz-Status"     = if ($winLic) { "Aktiviert (lizenziert)" } else { "⚠️ Nicht aktiviert/unbekannt" }
                    "Teil-Produktkey"   = if ($winLic) { $winLic.PartialProductKey } else { "-" }
                }
                $licHTML += "<h4>Windows-Lizenz: $server</h4>"
                $licHTML += (ConvertTo-HTMLTable -Data @($winLicInfo))
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Log -Message "Windows-Lizenz Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
        }
    }

    New-HTMLSection -Title "Lizenzierung (Exchange & Windows)" -Content $licHTML
}

# ---------------------------------------------------------------
# 2.1 DISK SPACE CHECK (DB / LOGS / TEMP)
# ---------------------------------------------------------------
function Get-DiskSpaceInfo {
    <#
    .SYNOPSIS
        Prüft verfügbaren Speicherplatz auf kritischen Exchange-Laufwerken.
        Warnt wenn < 30% verfügbar.
    #>
    Write-Log -Message "=== Sammle Disk Space Information ===" -Level "INFO"

    $diskHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Disk Space für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $diskList = @()

            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                
                if ($isLocal) {
                    $drives = Get-Volume -ErrorAction SilentlyContinue
                } else {
                    $drives = Invoke-Command -ComputerName $server -ScriptBlock {
                        Get-Volume -ErrorAction SilentlyContinue
                    } -ErrorAction SilentlyContinue
                }

                if ($drives) {
                    foreach ($drive in $drives) {
                        if ($drive.Size -gt 0) {
                            $usedPercent = [math]::Round(($drive.Size - $drive.SizeRemaining) / $drive.Size * 100, 1)
                            $freePercent = 100 - $usedPercent
                            $sizeGB = [math]::Round($drive.Size / 1GB, 2)
                            $freeGB = [math]::Round($drive.SizeRemaining / 1GB, 2)
                            
                            $status = if ($freePercent -lt 10) { 
                                "🔴 KRITISCH (< 10%)" 
                            }
                            elseif ($freePercent -lt 30) { 
                                "🟠 WARNUNG (< 30%)" 
                            }
                            elseif ($freePercent -lt 50) {
                                "🟡 HINWEIS (< 50%)"
                            }
                            else { 
                                "✅ OK" 
                            }
                            
                            $diskList += [PSCustomObject]@{
                                "Laufwerk"        = $drive.DriveLetter
                                "Beschreibung"    = $drive.FileSystemLabel
                                "Gesamt (GB)"     = $sizeGB
                                "Frei (GB)"       = $freeGB
                                "Frei (%)"        = "$freePercent %"
                                "Status"          = $status
                            }
                        }
                    }
                }
            }
            catch {
                Write-Log -Message "Disk-Abfrage für $server fehlgeschlagen: $_" -Level "WARNING"
            }

            if ($diskList -and @($diskList).Count -gt 0) {
                $serverHTML += "<h4>Festplatten-Auslastung</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data $diskList)
                $serverHTML += "<p><strong>Best Practice:</strong> Exchange empfiehlt mindestens 30% freien Speicher auf DB- und Log-Laufwerken.</p>"
            }
            else {
                $serverHTML += "<p class='no-data'>Disk Space Information nicht verfügbar.</p>"
            }

            $diskHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Disk Space Info für ${server}: $_" -Level "ERROR"
            $diskHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Disk Space (DB/Logs/Temp)" -Content $diskHTML
}

# ---------------------------------------------------------------
# 2.2 LDAP CONNECTIVITY CHECK
# ---------------------------------------------------------------
function Get-LDAPConnectivityInfo {
    <#
    .SYNOPSIS
        Prüft ob Domain Controller (LDAP) erreichbar ist.
        Basis für Exchange-Funktionalität.
    #>
    Write-Log -Message "=== Sammle LDAP Connectivity Information ===" -Level "INFO"

    $ldapHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "LDAP Connectivity für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                
                if ($isLocal) {
                    # Lokal: Nutze [ADSI]
                    $rootDSE = [ADSI]"LDAP://RootDSE"
                    $dcName = $rootDSE.Properties["dnsHostName"][0]
                    $dcAvailable = "✅ OK"
                } else {
                    # Remote: Test über Remoting
                    $dcTest = Invoke-Command -ComputerName $server -ScriptBlock {
                        try {
                            $rootDSE = [ADSI]"LDAP://RootDSE"
                            return @{ 
                                DC = $rootDSE.Properties["dnsHostName"][0]
                                Available = "✅ OK"
                            }
                        }
                        catch {
                            return @{ 
                                DC = "Nicht erreichbar"
                                Available = "🔴 FEHLER: $_"
                            }
                        }
                    } -ErrorAction SilentlyContinue
                    
                    $dcName = $dcTest.DC
                    $dcAvailable = $dcTest.Available
                }

                $ldapInfo = [PSCustomObject]@{
                    "Domain Controller"  = $dcName
                    "LDAP Verbindung"    = $dcAvailable
                    "Status"             = if ($dcAvailable -like "*✅*") { "✅ OK - AD erreichbar" } else { "🔴 FEHLER - AD nicht erreichbar" }
                }

                $serverHTML += (ConvertTo-HTMLTable -Data @($ldapInfo))
                $serverHTML += "<p><strong>Hinweis:</strong> Exchange benötigt aktive LDAP-Verbindung zu Active Directory (LDAP Port 389/636). Ohne AD funktioniert Exchange nicht.</p>"
            }
            catch {
                Write-Log -Message "LDAP-Test für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>LDAP-Verbindung konnte nicht getestet werden: $_</p>"
            }

            $ldapHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei LDAP Info für ${server}: $_" -Level "ERROR"
            $ldapHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "LDAP Konnektivität (Domain Controller)" -Content $ldapHTML
}

# ---------------------------------------------------------------
# 2.3 RPC PORT 135 CHECK
# ---------------------------------------------------------------
function Get-RPCPortStatusInfo {
    <#
    .SYNOPSIS
        Prüft ob RPC Port 135 aktiv ist (TCP und UDP).
        Notwendig für Exchange Remote Procedure Call.
    #>
    Write-Log -Message "=== Sammle RPC Port Status ===" -Level "INFO"

    $rpcHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "RPC Port für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                
                if ($isLocal) {
                    # Lokal: Prüfe RPC Endpoint Mapper
                    $rpcStatus = Get-Service -Name RpcSs -ErrorAction SilentlyContinue
                    $rpcEndpoint = Test-NetConnection -ComputerName localhost -Port 135 -ErrorAction SilentlyContinue
                } else {
                    # Remote: Test RPC Port
                    $rpcTest = Invoke-Command -ComputerName $server -ScriptBlock {
                        $rpcStatus = Get-Service -Name RpcSs -ErrorAction SilentlyContinue
                        $rpcEndpoint = Test-NetConnection -ComputerName localhost -Port 135 -ErrorAction SilentlyContinue
                        return @{
                            Service = $rpcStatus
                            Test = $rpcEndpoint
                        }
                    } -ErrorAction SilentlyContinue
                    
                    $rpcStatus = $rpcTest.Service
                    $rpcEndpoint = $rpcTest.Test
                }

                $rpcServiceStatus = if ($rpcStatus.Status -eq "Running") { "✅ Läuft" } else { "🔴 Nicht aktiv" }
                $rpcPortStatus = if ($rpcEndpoint.TcpTestSucceeded) { "✅ Erreichbar" } else { "⚠️ Test fehlgeschlagen" }

                $rpcInfo = [PSCustomObject]@{
                    "RPC Service (RpcSs)"  = $rpcServiceStatus
                    "Port 135 (TCP/UDP)"   = $rpcPortStatus
                    "Status"               = if ($rpcServiceStatus -like "*✅*" -and $rpcPortStatus -like "*✅*") { "✅ OK" } else { "🔴 FEHLER" }
                }

                $serverHTML += (ConvertTo-HTMLTable -Data @($rpcInfo))
                $serverHTML += "<p><strong>Hinweis:</strong> RPC Port 135 (TCP/UDP) wird für Exchange-Verwaltung und Cluster-Kommunikation benötigt. Service muss laufen.</p>"
            }
            catch {
                Write-Log -Message "RPC-Test für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>RPC-Status konnte nicht ermittelt werden: $_</p>"
            }

            $rpcHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei RPC Info für ${server}: $_" -Level "ERROR"
            $rpcHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "RPC Port 135 Status" -Content $rpcHTML
}

# ---------------------------------------------------------------
# 2.4 MAPI CONNECTIVITY CHECK
# ---------------------------------------------------------------
function Get-MAPIConnectivityInfo {
    <#
    .SYNOPSIS
        Prüft MAPI-Endpunkte und Konnektivität (Outlook).
        Exchange 2019 unterstützt nur MAPI über HTTP (modern).
    #>
    Write-Log -Message "=== Sammle MAPI Connectivity Information ===" -Level "INFO"

    $mapiHTML = ""

    try {
        # MAPI Information (Org-weit)
        $mapiOrg = [PSCustomObject]@{
            "MAPI über HTTP"      = "✅ Erforderlich (Standard in 2019)"
            "Legacy MAPI"         = "❌ Nicht unterstützt (nur via HTTP)"
            "Outlook-Versionen"   = "Outlook 2016+ oder Outlook für Mac 2016+"
        }

        $mapiHTML += "<h3>MAPI-Architektur (Organisation)</h3>"
        $mapiHTML += (ConvertTo-HTMLTable -Data @($mapiOrg))

        # Pro Server: MAPI Virtual Directory
        foreach ($server in $ExchangeServers) {
            try {
                $mapiVdir = Get-MapiVirtualDirectory -Server $server -ErrorAction SilentlyContinue | Select-Object Server, InternalUrl, ExternalUrl, @{N="Status";E={"✅ OK"}}
                
                if ($mapiVdir) {
                    $mapiHTML += "<h4>MAPI Virtual Directory: $server</h4>"
                    $mapiHTML += (ConvertTo-HTMLTable -Data $mapiVdir)
                }
            }
            catch {
                Write-Log -Message "MAPI VDir Abfrage für $server fehlgeschlagen: $_" -Level "WARNING"
            }
        }

        $mapiHTML += "<p><strong>Hinweis:</strong> MAPI/HTTP ist die modern Outlook-Verbindungsmethode. RPC-over-HTTP ist veraltet und wird nicht mehr empfohlen.</p>"
    }
    catch {
        Write-Log -Message "MAPI Connectivity Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $mapiHTML += "<p class='error'>MAPI-Information nicht verfügbar: $_</p>"
    }

    New-HTMLSection -Title "MAPI Connectivity (Outlook)" -Content $mapiHTML
}

# ---------------------------------------------------------------
# 2.5 DATABASE COPIES CHECK
# ---------------------------------------------------------------
function Get-DatabaseCopiesInfo {
    <#
    .SYNOPSIS
        Listet alle Datenbankgesamt und deren Kopien auf.
        Prüft Redundanz und DAG-Status.
    #>
    Write-Log -Message "=== Sammle Database Copies Information ===" -Level "INFO"

    $dbCopyHTML = ""

    try {
        $dbList = Get-MailboxDatabase -ErrorAction SilentlyContinue
        
        if ($dbList) {
            $dbCopyData = @()
            foreach ($db in $dbList) {
                $copies = Get-MailboxDatabaseCopyStatus -Identity $db.Identity -ErrorAction SilentlyContinue
                foreach ($copy in $copies) {
                    $dbCopyData += [PSCustomObject]@{
                        "Database"        = $db.Name
                        "Server"          = $copy.ServerName
                        "Status"          = if ($copy.Status -eq "Healthy") { "✅ Healthy" } else { "🔴 $($copy.Status)" }
                        "Content Index"   = if ($copy.ContentIndexStatus -eq "Healthy") { "✅ OK" } else { "🟠 $($copy.ContentIndexStatus)" }
                        "Replay Lag (ms)" = $copy.ReplayLagTime.TotalMilliseconds
                        "Truncation Lag"  = $copy.TruncationLagTime
                    }
                }
            }

            if ($dbCopyData) {
                $dbCopyHTML += "<h3>Datenbankgesamt und Kopien</h3>"
                $dbCopyHTML += (ConvertTo-HTMLTable -Data $dbCopyData)
            }
        }
    }
    catch {
        Write-Log -Message "Database Copies Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $dbCopyHTML += "<p class='error'>Database Copies nicht verfügbar: $_</p>"
    }

    $dbCopyHTML += "<p><strong>Hinweis:</strong> In einer DAG sollten kritische Datenbanken mehrere Kopien haben.</p>"
    New-HTMLSection -Title "Database Copies & Redundanz" -Content $dbCopyHTML
}

# ---------------------------------------------------------------
# 2.6 ANTIVIRUS EXCLUSIONS CHECK
# ---------------------------------------------------------------
function Get-AntivirusExclusionsInfo {
    <#
    .SYNOPSIS
        Prüft ob empfohlene Exchange-Pfade von AV ausgenommen sind.
        Häufiger Grund für Performance-Issues.
    #>
    Write-Log -Message "=== Sammle Antivirus Exclusions Information ===" -Level "INFO"

    $avHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Antivirus Exclusions für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            
            # Empfohlene Ausschlussmuster
            $recommendedExclusions = @(
                "C:\Program Files\Microsoft\Exchange Server\*\Bin\*",
                "C:\Program Files\Microsoft\Exchange Server\*\TransportRoles\*",
                "C:\Program Files\Microsoft\Exchange Server\*\EdgeTransportRoles\*",
                "C:\ExchangeData\*",
                "C:\Program Files\Microsoft\Exchange Server\*\Mailbox\*\MDBTEMP*",
                "C:\inetpub\logs\logfiles\*"
            )

            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                
                if ($isLocal) {
                    $winDefender = Get-MpPreference -ErrorAction SilentlyContinue
                    $exclusions = $winDefender.ExclusionPath
                } else {
                    $exclTest = Invoke-Command -ComputerName $server -ScriptBlock {
                        $winDefender = Get-MpPreference -ErrorAction SilentlyContinue
                        return $winDefender.ExclusionPath
                    } -ErrorAction SilentlyContinue
                    
                    $exclusions = $exclTest
                }

                if ($exclusions -and $exclusions.Count -gt 0) {
                    $avHTML += "<h4>Konfigurierte Ausschlüsse</h4>"
                    $avHTML += "<ul>"
                    foreach ($excl in $exclusions) {
                        $avHTML += "<li>$excl</li>"
                    }
                    $avHTML += "</ul>"
                    
                    # Prüfe ob empfohlene dabei sind
                    $missingExclusions = @()
                    foreach ($rec in $recommendedExclusions) {
                        if (-not ($exclusions -like $rec)) {
                            $missingExclusions += $rec
                        }
                    }
                    
                    if ($missingExclusions.Count -gt 0) {
                        $avHTML += "<h4>⚠️ Fehlende empfohlene Ausschlüsse:</h4>"
                        $avHTML += "<ul style='color: orange;'>"
                        foreach ($miss in $missingExclusions) {
                            $avHTML += "<li>$miss</li>"
                        }
                        $avHTML += "</ul>"
                    }
                }
                else {
                    $avHTML += "<p class='warning'>Keine Antivirus-Ausschlüsse konfiguriert oder Windows Defender nicht verfügbar.</p>"
                }
            }
            catch {
                Write-Log -Message "AV-Exclusion Check für $server fehlgeschlagen: $_" -Level "WARNING"
                $avHTML += "<p class='error'>Antivirus-Exclusions konnten nicht gelesen werden: $_</p>"
            }

            $avHTML += "<p><strong>Best Practice:</strong> Alle Exchange-Bin-, Logging- und Datenbank-Pfade sollten von AV-Scanning ausgenommen sein.</p>"

            $serverHTML += $avHTML
        }
        catch {
            Write-Log -Message "Fehler bei AV Info für ${server}: $_" -Level "ERROR"
            $serverHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Antivirus Ausschlüsse" -Content $serverHTML
}

# ---------------------------------------------------------------
# 34. DKIM-KONFIGURATION
# ---------------------------------------------------------------
function Get-DKIMInfo {
    <#
    .SYNOPSIS
        Dokumentiert die DKIM-Konfiguration (DNS-Selektoren) sowie installierte
        Transport-Agents, die DKIM-Signierung übernehmen könnten.
    #>
    Write-Log -Message "=== Sammle DKIM-Informationen ===" -Level "INFO"

    $dkimHTML = ""

    # --- Akzeptierte Domänen ermitteln, dann DKIM-DNS prüfen ---
    try {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DomainName
    }
    catch {
        $acceptedDomains = @()
        Write-Log -Message "AcceptedDomain-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # Übliche DKIM-Selektoren (Exchange/M365 und gängige Tools)
    $commonSelectors = @("selector1", "selector2", "default", "google", "k1", "s1", "s2", "dkim", "mail")

    $dkimResults = [System.Collections.ArrayList]::new()
    foreach ($domain in $acceptedDomains) {
        $domainStr = $domain.ToString()
        $found = $false
        foreach ($selector in $commonSelectors) {
            try {
                $dkimRecord = Resolve-DnsName -Name "$selector._domainkey.$domainStr" -Type TXT -ErrorAction SilentlyContinue
                if ($dkimRecord) {
                    $txt = ($dkimRecord | Where-Object { $_.Type -eq "TXT" }).Strings -join ""
                    if ($txt -match "v=DKIM1|p=") {
                        [void]$dkimResults.Add([PSCustomObject]@{
                            Domain    = $domainStr
                            Selektor  = $selector
                            Status    = "✅ DKIM gefunden"
                            "Record (gekürzt)" = if ($txt.Length -gt 80) { $txt.Substring(0,80) + "..." } else { $txt }
                        })
                        $found = $true
                    }
                }
            }
            catch { }
        }
        if (-not $found) {
            [void]$dkimResults.Add([PSCustomObject]@{
                Domain    = $domainStr
                Selektor  = "-"
                Status    = "⚠️ Kein DKIM-Record gefunden"
                "Record (gekürzt)" = "Keine gängigen Selektoren auflösbar"
            })
        }
    }

    $dkimHTML += "<h3>DKIM DNS-Records je akzeptierter Domäne</h3>"
    $dkimHTML += "<p><em>Geprüfte Selektoren: $($commonSelectors -join ', ')</em></p>"
    $dkimHTML += (ConvertTo-HTMLTable -Data $dkimResults -NoDataMessage "Keine akzeptierten Domänen gefunden")

    # --- Transport-Agents (DKIM-Signer etc.) ---
    try {
        $agents = Get-TransportAgent -ErrorAction SilentlyContinue
        if ($agents) {
            $agentData = foreach ($a in $agents) {
                [PSCustomObject]@{
                    Name      = $a.Identity
                    Aktiviert = $a.Enabled
                    Priorität = $a.Priority
                    "DKIM-relevant" = if ($a.Identity -match "DKIM|DomainKeys|Sign") { "Ja" } else { "Nein" }
                }
            }
            $dkimHTML += "<h3>Installierte Transport-Agents</h3>"
            $dkimHTML += "<p><em>Exchange On-Premises bietet keine native DKIM-Signierung. DKIM erfolgt i.d.R. über einen "
            $dkimHTML += "Transport-Agent (z.B. 'Exchange DkimSigner') oder ein vorgelagertes Gateway/Smarthost.</em></p>"
            $dkimHTML += (ConvertTo-HTMLTable -Data $agentData)
        }
    }
    catch {
        Write-Log -Message "TransportAgent-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Hinweis ---
    $dkimHTML += "<h3>Empfehlung DKIM/DMARC/SPF</h3>"
    $dkimHTML += "<p><em>Für vollständige E-Mail-Authentifizierung sollten SPF, DKIM und DMARC kombiniert werden. "
    $dkimHTML += "Prüfen Sie ob die DKIM-Signierung am Mail-Gateway, Smarthost oder via Transport-Agent erfolgt und "
    $dkimHTML += "ob die DNS-Records mit den verwendeten Selektoren übereinstimmen.</em></p>"

    New-HTMLSection -Title "DKIM-Konfiguration" -Content $dkimHTML
}

# ---------------------------------------------------------------
# 35. TLS/SSL KONFIGURATION MIT BEST PRACTICE BEWERTUNG
# ---------------------------------------------------------------
function Get-TLSSSLConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert die gesamte TLS/SSL-Konfiguration mit Best Practice Bewertung.
        Prüft: Windows Registry (TLS 1.0-1.3), Cipher Suites, Send/Receive Connectors,
        Zertifikate, PowerShell Remoting, und gibt eine Sicherheitsbewertung ab.
    #>
    Write-Log -Message "=== Sammle TLS/SSL Konfiguration ===" -Level "INFO"

    $tlsHTML = ""

    # --- Best Practice Matrix (Microsoft Empfehlungen für Exchange) ---
    $bestPractices = @{
        "TLS 1.0" = @{ Recommended = $false; Why = "Veraltet & unsicher (POODLE). MUSS deaktiviert sein." }
        "TLS 1.1" = @{ Recommended = $false; Why = "Veraltet (RFC 8996 veraltet). SOLLTE deaktiviert sein." }
        "TLS 1.2" = @{ Recommended = $true;  Why = "ERFORDERLICH. Mindeststandard für alle Verbindungen." }
        "TLS 1.3" = @{ Recommended = $true;  Why = "EMPFOHLEN. Neuester Standard (RFC 8446), besser Sicherheit." }
    }

    foreach ($server in $ExchangeServers) {
        try {
            $serverHTML = "<h3 class='server-break'>Server: $server - TLS/SSL Konfiguration</h3>"
            
            # ===== Methode 1: Remote Registry (kein WinRM nötig) =====
            $tlsVersions = @{}
            $regPath = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
            
            Write-Log -Message "Lese TLS-Versionen aus Registry ($server)..." -Level "INFO"
            
            foreach ($tlsVer in @("TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3")) {
                try {
                    $regSubPath = "$regPath\$tlsVer\Server"
                    $enabled = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regSubPath -ValueName "Enabled"
                    $tlsVersions[$tlsVer] = @{
                        Enabled = if ($enabled -eq 1) { "JA (AKTIVIERT)" } else { "NEIN (deaktiviert)" }
                        EnabledValue = $enabled
                        Recommended = $bestPractices[$tlsVer].Recommended
                        Why = $bestPractices[$tlsVer].Why
                    }
                }
                catch {
                    $tlsVersions[$tlsVer] = @{
                        Enabled = "Unbekannt (nicht in Registry)"
                        EnabledValue = $null
                        Recommended = $bestPractices[$tlsVer].Recommended
                        Why = $bestPractices[$tlsVer].Why
                    }
                }
            }
            
            # TLS-Versions-Tabelle
            $tlsVersionsData = @()
            $overallSecurityScore = 100
            
            foreach ($tlsVer in @("TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3")) {
                $status = $tlsVersions[$tlsVer].Enabled
                $isEnabled = ($tlsVersions[$tlsVer].EnabledValue -eq 1)
                $recommended = $tlsVersions[$tlsVer].Recommended
                
                # Bewertung
                if ($tlsVer -eq "TLS 1.0" -and $isEnabled) {
                    $assessment = "🔴 KRITISCH - MUSS deaktiviert sein!"
                    $overallSecurityScore -= 40
                }
                elseif ($tlsVer -eq "TLS 1.1" -and $isEnabled) {
                    $assessment = "🟡 WARNUNG - SOLLTE deaktiviert sein (RFC 8996)"
                    $overallSecurityScore -= 15
                }
                elseif ($tlsVer -eq "TLS 1.2" -and -not $isEnabled) {
                    $assessment = "🔴 KRITISCH - Muss aktiviert sein (Exchange Anforderung)!"
                    $overallSecurityScore -= 50
                }
                elseif ($tlsVer -eq "TLS 1.2" -and $isEnabled) {
                    $assessment = "✅ OK - Erforderlich und aktiv"
                }
                elseif ($tlsVer -eq "TLS 1.3" -and $isEnabled) {
                    $assessment = "✅ SEHR GUT - Modern & sicher"
                    $overallSecurityScore += 10
                }
                else {
                    $assessment = "ⓘ Nicht aktiviert (optional)"
                }
                
                $tlsVersionsData += [PSCustomObject]@{
                    Version = $tlsVer
                    Status = $status
                    Empfohlen = if ($recommended) { "JA" } else { "NEIN" }
                    Begründung = $tlsVersions[$tlsVer].Why
                    Bewertung = $assessment
                }
            }
            
            $serverHTML += "<h4>TLS-Versionen (Windows Registry)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $tlsVersionsData)
            
            # Gesamtbewertung
            if ($overallSecurityScore -lt 0) { $overallSecurityScore = 0 }
            if ($overallSecurityScore -gt 100) { $overallSecurityScore = 100 }
            
            $securityRating = switch -Exact ($overallSecurityScore) {
                { $_ -ge 90 } { "🟢 AUSGEZEICHNET"; break }
                { $_ -ge 70 } { "🟢 GUT"; break }
                { $_ -ge 50 } { "🟡 AKZEPTABEL (Verbesserungen empfohlen)"; break }
                { $_ -ge 30 } { "🔴 SCHWACH (Kritische Mängel)"; break }
                default { "🔴 SEHR SCHLECHT (Sofort beheben!)"; break }
            }
            
            $serverHTML += "<div class='summary-box'><h4>TLS/SSL Sicherheitsbewertung: $securityRating ($overallSecurityScore/100)</h4></div>"
            
            # ===== Cipher Suites (via PowerShell Remoting) =====
            try {
                Write-Log -Message "Lese Cipher Suites von $server..." -Level "INFO"
                $cipherData = Invoke-Command -ComputerName $server -ScriptBlock {
                    try {
                        $winRMPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp\DefaultSecureProtocols"
                        $dsPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers"
                        
                        $tlsDefault = Get-ItemProperty -Path $winRMPath -ErrorAction SilentlyContinue
                        $ciphers = Get-ChildItem -Path $dsPath -ErrorAction SilentlyContinue
                        
                        return @{
                            TLSDefault = $tlsDefault.DefaultSecureProtocols
                            CipherCount = $ciphers.Count
                            CipherNames = ($ciphers | Select-Object -ExpandProperty PSChildName | Select-Object -First 20) -join ", "
                        }
                    }
                    catch { return $null }
                } -ErrorAction SilentlyContinue
                
                if ($cipherData) {
                    $serverHTML += "<h4>Cipher Suites Konfiguration</h4>"
                    $cipherHTML = "<p><strong>Standard-Protokolle:</strong> "
                    if ($cipherData.TLSDefault) {
                        # Dekodierung der Hexadezimalwerte
                        if ($cipherData.TLSDefault -eq 0x00000008) {
                            $cipherHTML += "TLS 1.0 (VERALTET!)"
                        }
                        elseif ($cipherData.TLSDefault -eq 0x00000020) {
                            $cipherHTML += "TLS 1.1 (VERALTET!)"
                        }
                        elseif ($cipherData.TLSDefault -eq 0x00000080) {
                            $cipherHTML += "TLS 1.2 (Standard)"
                        }
                        else {
                            $cipherHTML += "Hex: 0x$([Convert]::ToString($cipherData.TLSDefault, 16))"
                        }
                    }
                    $cipherHTML += "</p>"
                    $cipherHTML += "<p><strong>Konfigurierte Cipher Suites:</strong> $($cipherData.CipherCount) Suites konfiguriert</p>"
                    $cipherHTML += "<p><strong>Aktive Ciphers (Top 20):</strong></p><pre style='background-color:#f5f5f5; padding:10px; overflow-x:auto;'>$($cipherData.CipherNames)</pre>"
                    $serverHTML += $cipherHTML
                }
            }
            catch {
                Write-Log -Message "Cipher Suite Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
                $serverHTML += "<p class='no-data'>Cipher Suite-Details nicht abrufbar (benötigt Invoke-Command).</p>"
            }
            
            # ===== Receive Connectors TLS Konfiguration =====
            try {
                Write-Log -Message "Prüfe Receive Connector TLS auf $server..." -Level "INFO"
                $rcConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue | Where-Object { $_.Server -eq $server }
                
                if ($rcConnectors) {
                    $rcTLSData = foreach ($rc in $rcConnectors) {
                        [PSCustomObject]@{
                            Name = $rc.Name
                            "TLS erforderlich" = $rc.RequireTLS
                            "TLS Auth Level" = $rc.TlsAuthLevel
                            "TLS Zertifikat" = $rc.TlsCertificateName
                            "Auth Mechanismen" = ($rc.AuthMechanism -join ", ")
                            "Domains mit sicherer TLS" = if ((Get-TransportConfig).TLSReceiveDomainSecureList.Count -gt 0) { "Ja" } else { "Nein" }
                            Bewertung = if ($rc.RequireTLS) { "✅ OK" } else { "🟡 Warnung - TLS sollte erzwungen sein" }
                        }
                    }
                    $serverHTML += "<h4>Receive Connectors - TLS Konfiguration</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $rcTLSData)
                }
            }
            catch {
                Write-Log -Message "Receive Connector TLS-Prüfung fehlgeschlagen: $_" -Level "WARNING"
            }
            
            # ===== Send Connectors TLS Konfiguration =====
            try {
                Write-Log -Message "Prüfe Send Connector TLS auf $server..." -Level "INFO"
                $scConnectors = Get-SendConnector -ErrorAction SilentlyContinue | Where-Object { $_.SourceTransportServers -contains $server }
                
                if ($scConnectors) {
                    $scTLSData = foreach ($sc in $scConnectors) {
                        [PSCustomObject]@{
                            Name = $sc.Name
                            "TLS erforderlich" = $sc.RequireTLS
                            "TLS Auth Level" = $sc.TlsAuthLevel
                            "TLS Domain" = $sc.TlsDomain
                            "TLS Zertifikat" = $sc.TlsCertificateName
                            "Smart Host Auth" = $sc.SmartHostAuthMechanism
                            Bewertung = if ($sc.RequireTLS) { "✅ OK" } else { "🟡 Warnung - TLS sollte erzwungen sein" }
                        }
                    }
                    $serverHTML += "<h4>Send Connectors - TLS Konfiguration</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $scTLSData)
                }
            }
            catch {
                Write-Log -Message "Send Connector TLS-Prüfung fehlgeschlagen: $_" -Level "WARNING"
            }
            
            # ===== Zertifikate mit TLS-Validierung =====
            try {
                Write-Log -Message "Prüfe Zertifikate auf $server..." -Level "INFO"
                $certs = Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue
                
                if ($certs) {
                    $certTLSData = foreach ($cert in $certs) {
                        $daysLeft = [math]::Round(($cert.NotAfter - (Get-Date)).TotalDays, 0)
                        if ($daysLeft -lt 0) {
                            $certStatus = "ABGELAUFEN"
                        }
                        elseif ($daysLeft -lt 30) {
                            $certStatus = "DRINGEND - laeuft in $daysLeft Tagen ab"
                        }
                        elseif ($daysLeft -lt 90) {
                            $certStatus = "WARNUNG - laeuft in $daysLeft Tagen ab"
                        }
                        else {
                            $certStatus = "OK - $daysLeft Tage verbleibend"
                        }
                        
                        # Zertifikat-Typ Bewertung
                        $certTypeAssess = if ($cert.IsSelfSigned) { "⚠️ SELF-SIGNED" } else { "✅ CA-signiert" }
                        
                        [PSCustomObject]@{
                            Subject = $cert.Subject
                            Dienste = ($cert.Services -join ", ")
                            "Nicht nach" = $cert.NotAfter.ToString("dd.MM.yyyy")
                            Status = $certStatus
                            "Zertifikatstyp" = $certTypeAssess
                            "Public Key Länge" = "$($cert.PublicKeySize) Bit"
                        }
                    }
                    $serverHTML += "<h4>Zertifikate (TLS/SSL)</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $certTLSData)
                }
            }
            catch {
                Write-Log -Message "Zertifikat-Prüfung fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }
            
            # ===== PowerShell Remoting TLS =====
            try {
                $psRemotingPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service"
                $psAuthValue = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $psRemotingPath -ValueName "Auth"
                
                $serverHTML += "<h4>PowerShell Remoting TLS-Einstellungen</h4>"
                $serverHTML += "<p><strong>WinRM Service Auth (Registry):</strong> $(if ($psAuthValue) { $psAuthValue } else { 'Standard' })</p>"
                $serverHTML += "<p><em>Hinweis: WinRM sollte nur Kerberos oder Certificate-basierte Auth verwenden, nicht Basic Auth über HTTP.</em></p>"
            }
            catch {
                Write-Log -Message "PowerShell Remoting Prüfung fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }
            
            # ===== Best Practice Empfehlungen =====
            $serverHTML += "<h4>Best Practice Empfehlungen (Microsoft Exchange)</h4>"
            $serverHTML += "<ul>"
            $serverHTML += "<li><strong>TLS 1.0 &amp; 1.1:</strong> Deaktivieren (POODLE, RC4, andere Schwachstellen). Registry-Wert 'Enabled' auf 0 setzen.</li>"
            $serverHTML += "<li><strong>TLS 1.2:</strong> Muss aktiviert sein (Mindestanforderung Exchange). Verwende mindestens SHA-256 Signierung.</li>"
            $serverHTML += "<li><strong>TLS 1.3:</strong> Falls Windows Server 2022+ oder höher - aktivieren für zusätzliche Sicherheit.</li>"
            $serverHTML += "<li><strong>Cipher Suites:</strong> Verwende nur Strong Ciphers (ECDHE, AES-GCM). Deaktiviere anon, NULL, MD5, RC4 Suites.</li>"
            $serverHTML += "<li><strong>Zertifikate:</strong> Mindestens 2048-Bit RSA oder 256-Bit ECDSA. SHA-256 Signierung. Validität &gt; 1 Jahr.</li>"
            $serverHTML += "<li><strong>Connector TLS:</strong> RequireTLS=True für interne &amp; externe Connectors. TlsAuthLevel=DomainValidation oder CertificateValidation.</li>"
            $serverHTML += "<li><strong>DNS-Sicherheit:</strong> Nutze DNSSEC für MX-Records. Authentifizierung per SPF/DKIM/DMARC.</li>"
            $serverHTML += "<li><strong>Monitoring:</strong> Überwache Zertifikat-Ablauf &amp; TLS-Handshake-Fehler in Event Logs (SCHANNEL Events).</li>"
            $serverHTML += "</ul>"
            
            # Verweis auf MS-Dokumentation
            $serverHTML += "<h4>Weiterführende Ressourcen</h4>"
            $serverHTML += "<p><em><a href='https://docs.microsoft.com/en-us/Exchange/security-and-compliance/best-practices-for-security' target='_blank'>Microsoft Exchange Security Best Practices</a></em></p>"
            $serverHTML += "<p><em><a href='https://docs.microsoft.com/en-us/windows-server/security/tls/tls-registry-settings' target='_blank'>TLS Registry Settings - Microsoft Docs</a></em></p>"
            
            $tlsHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Allgemeiner Fehler bei TLS/SSL-Konfiguration für ${server}: $_" -Level "ERROR"
            $tlsHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "TLS/SSL Konfiguration & Best Practices" -Content $tlsHTML
}

# ---------------------------------------------------------------
# 36. HYBRID-KONFIGURATION MIT EXCHANGE ONLINE
# ---------------------------------------------------------------
function Get-HybridConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Hybrid-Konfiguration zwischen Exchange On-Premises und Exchange Online.
    #>
    Write-Log -Message "=== Sammle Hybrid-Konfiguration ===" -Level "INFO"

    $hybridHTML = ""

    # --- Hybrid Configuration Object ---
    try {
        $hybridConfig = Get-HybridConfiguration -ErrorAction Stop
        if ($hybridConfig) {
            $hybridOverview = [PSCustomObject]@{
                "Hybrid-Domains"            = ($hybridConfig.Domains -join ', ')
                "Features"                  = ($hybridConfig.Features -join ', ')
                "On-Premises Server"        = ($hybridConfig.OnPremisesSmartHost)
                "Receiving Transport SRV"   = ($hybridConfig.ReceivingTransportServers -join ', ')
                "Sending Transport SRV"     = ($hybridConfig.SendingTransportServers -join ', ')
                "Edge Transport SRV"        = if ($hybridConfig.EdgeTransportServers) { ($hybridConfig.EdgeTransportServers -join ', ') } else { "Nicht konfiguriert" }
                "TLS Zertifikat Name"       = $hybridConfig.TlsCertificateName
                "Service Instance"          = $hybridConfig.ServiceInstance
            }
            $hybridHTML += "<h3>Hybrid-Konfiguration (Übersicht)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data @($hybridOverview))
        }
        else {
            $hybridHTML += "<p class='no-data'>Keine Hybrid-Konfiguration gefunden. Die Umgebung ist nicht hybrid konfiguriert.</p>"
            New-HTMLSection -Title "Hybrid-Konfiguration (Exchange Online)" -Content $hybridHTML
            return
        }
    }
    catch {
        Write-Log -Message "Hybrid-Konfiguration nicht vorhanden oder Fehler: $_" -Level "WARNING"
        $hybridHTML += "<p class='no-data'>Keine Hybrid-Konfiguration vorhanden oder Cmdlet nicht verfügbar. Umgebung ist möglicherweise nicht hybrid.</p>"
        New-HTMLSection -Title "Hybrid-Konfiguration (Exchange Online)" -Content $hybridHTML
        return
    }

    # --- Intra-Organization Connector (OAuth) ---
    try {
        $intraOrgConnectors = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue
        if ($intraOrgConnectors) {
            $iocData = foreach ($ioc in $intraOrgConnectors) {
                [PSCustomObject]@{
                    Name               = $ioc.Name
                    "Target Address"   = ($ioc.TargetAddressDomains -join ', ')
                    "Discovery Endpoint" = $ioc.DiscoveryEndpoint
                    Enabled            = $ioc.Enabled
                }
            }
            $hybridHTML += "<h3>Intra-Organization Connectors (OAuth/Free-Busy)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $iocData)
        }
    }
    catch {
        Write-Log -Message "IntraOrganizationConnector-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Organization Relationship (Federation) ---
    try {
        $orgRelationships = Get-OrganizationRelationship -ErrorAction SilentlyContinue
        if ($orgRelationships) {
            $orData = foreach ($or in $orgRelationships) {
                [PSCustomObject]@{
                    Name                  = $or.Name
                    "Domain Names"        = ($or.DomainNames -join ', ')
                    Enabled               = $or.Enabled
                    FreeBusyAccess        = $or.FreeBusyAccessEnabled
                    "FreeBusy Level"      = $or.FreeBusyAccessLevel
                    MailTipsAccess        = $or.MailTipsAccessEnabled
                    "Target Application"  = $or.TargetApplicationUri
                    "Target Autodiscover" = $or.TargetAutodiscoverEpr
                }
            }
            $hybridHTML += "<h3>Organization Relationships (Frei/Gebucht-Freigabe)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $orData)
        }
    }
    catch {
        Write-Log -Message "OrganizationRelationship-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- OAuth Configuration ---
    try {
        $authServer = Get-AuthServer -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*ACS*" -or $_.Name -like "*evoSTS*" }
        if ($authServer) {
            $authData = foreach ($as in $authServer) {
                [PSCustomObject]@{
                    Name                = $as.Name
                    "Auth Metadata URL" = $as.AuthMetadataUrl
                    Enabled             = $as.Enabled
                    Type                = $as.Type
                }
            }
            $hybridHTML += "<h3>Auth Server (OAuth-Konfiguration)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $authData)
        }
    }
    catch {
        Write-Log -Message "AuthServer-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Partner Applications ---
    try {
        $partnerApps = Get-PartnerApplication -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq $true }
        if ($partnerApps) {
            $paData = foreach ($pa in $partnerApps) {
                [PSCustomObject]@{
                    Name             = $pa.Name
                    "Application ID" = $pa.ApplicationIdentifier
                    Enabled          = $pa.Enabled
                    "Auth Type"      = $pa.AuthMetadataUrl
                }
            }
            $hybridHTML += "<h3>Partner Applications</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $paData)
        }
    }
    catch {
        Write-Log -Message "PartnerApplication-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Hybrid Send/Receive Connectors ---
    try {
        $hybridSendConn = Get-SendConnector -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Office 365*" -or $_.Name -like "*Microsoft 365*" -or $_.Name -like "*Hybrid*" -or $_.Name -like "*Online*" }
        if ($hybridSendConn) {
            $hscData = foreach ($sc in $hybridSendConn) {
                [PSCustomObject]@{
                    Name              = $sc.Name
                    "Address Spaces"  = ($sc.AddressSpaces -join ', ')
                    "Smart Hosts"     = ($sc.SmartHosts -join ', ')
                    TlsAuthLevel      = $sc.TlsAuthLevel
                    TlsDomain         = $sc.TlsDomain
                    RequireTLS        = $sc.RequireTLS
                    Enabled           = $sc.Enabled
                    CloudServicesMailEnabled = $sc.CloudServicesMailEnabled
                }
            }
            $hybridHTML += "<h3>Hybrid Send Connectors (Richtung Exchange Online)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $hscData)
        }
    }
    catch {
        Write-Log -Message "Hybrid SendConnector-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    try {
        $hybridRecvConn = Get-ReceiveConnector -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Office 365*" -or $_.Name -like "*Microsoft 365*" -or $_.Name -like "*Hybrid*" -or $_.Name -like "*Online*" }
        if ($hybridRecvConn) {
            $hrcData = foreach ($rc in $hybridRecvConn) {
                [PSCustomObject]@{
                    Name              = $rc.Name
                    Server            = $rc.Server
                    Bindings          = ($rc.Bindings -join ', ')
                    "Remote IPs"      = ($rc.RemoteIPRanges -join ', ')
                    TlsCertificateName = $rc.TlsCertificateName
                    RequireTLS        = $rc.RequireTLS
                    AuthMechanism     = ($rc.AuthMechanism -join ', ')
                }
            }
            $hybridHTML += "<h3>Hybrid Receive Connectors (von Exchange Online)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $hrcData)
        }
    }
    catch {
        Write-Log -Message "Hybrid ReceiveConnector-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Migration Endpoints ---
    try {
        $migEndpoints = Get-MigrationEndpoint -ErrorAction SilentlyContinue
        if ($migEndpoints) {
            $meData = foreach ($me in $migEndpoints) {
                [PSCustomObject]@{
                    Name             = $me.Identity
                    Typ              = $me.EndpointType
                    "Remote Server"  = $me.RemoteServer
                    "Max Concurrent" = $me.MaxConcurrentMigrations
                    "Max Incremental" = $me.MaxConcurrentIncrementalSyncs
                }
            }
            $hybridHTML += "<h3>Migration Endpoints</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $meData)
        }
    }
    catch {
        Write-Log -Message "MigrationEndpoint-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Mailbox Move Requests (aktive Migrationen) ---
    try {
        $moveRequests = Get-MoveRequest -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne "Completed" }
        if ($moveRequests) {
            $mrData = foreach ($mr in $moveRequests) {
                [PSCustomObject]@{
                    DisplayName  = $mr.DisplayName
                    Status       = $mr.Status
                    Direction    = $mr.Direction
                    "Percent"    = $mr.PercentComplete
                    "Target DB"  = $mr.TargetDatabase
                    BatchName    = $mr.BatchName
                }
            }
            $hybridHTML += "<h3>Aktive Postfach-Migrationen (Move Requests)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data $mrData)
        }
        else {
            $hybridHTML += "<h3>Aktive Postfach-Migrationen</h3>"
            $hybridHTML += "<p class='no-data'>Keine aktiven Move Requests vorhanden.</p>"
        }
    }
    catch {
        Write-Log -Message "MoveRequest-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Remote Mailboxes (Zusammenfassung) ---
    try {
        $remoteMailboxes = Get-RemoteMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
        if ($remoteMailboxes) {
            $rmSummary = [PSCustomObject]@{
                "Anzahl Remote Mailboxes"     = $remoteMailboxes.Count
                "Remote User Mailbox"         = ($remoteMailboxes | Where-Object { $_.RemoteRecipientType -match "ProvisionMailbox" }).Count
                "Remote Room Mailbox"         = ($remoteMailboxes | Where-Object { $_.RemoteRecipientType -match "RoomMailbox" }).Count
                "Remote Shared Mailbox"       = ($remoteMailboxes | Where-Object { $_.RemoteRecipientType -match "SharedMailbox" }).Count
                "Remote Equipment Mailbox"    = ($remoteMailboxes | Where-Object { $_.RemoteRecipientType -match "EquipmentMailbox" }).Count
            }
            $hybridHTML += "<h3>Remote Mailboxes (in Exchange Online)</h3>"
            $hybridHTML += (ConvertTo-HTMLTable -Data @($rmSummary))
        }
    }
    catch {
        Write-Log -Message "RemoteMailbox-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Azure AD Connect / Entra Connect Status (via AD-Attribute) ---
    try {
        $hybridHTML += "<h3>Microsoft Entra Connect (AAD Connect) - Hinweis</h3>"
        $hybridHTML += "<p><em>Für vollständige Entra Connect-Dokumentation prüfen Sie bitte den Entra Connect Server direkt. "
        $hybridHTML += "Relevante Informationen: Sync-Intervall, gefilterte OUs, Attribut-Mappings, Sync-Regeln, Staging-Mode.</em></p>"

        # Prüfe ob AADConnect-Attribute im AD vorhanden sind
        $aadConnectCheck = Get-ADUser -Filter "msDS-cloudExtensionAttribute1 -like '*'" -Properties msDS-cloudExtensionAttribute1 -ResultSetSize 1 -ErrorAction SilentlyContinue
        if ($aadConnectCheck) {
            $hybridHTML += "<p>&#9989; Cloud-Attribute im AD vorhanden - Entra Connect scheint aktiv zu sein.</p>"
        }
    }
    catch {
        Write-Log -Message "AAD Connect Check fehlgeschlagen: $_" -Level "WARNING"
    }

    New-HTMLSection -Title "Hybrid-Konfiguration (Exchange Online)" -Content $hybridHTML
}

# ---------------------------------------------------------------
# 36. MESSAGE QUEUE ANALYSE
# ---------------------------------------------------------------
function Get-MessageQueueInfo {
    <#
    .SYNOPSIS
        Analysiert die aktuellen Transport-Warteschlangen (Queues) auf allen Exchange-Servern.
        Zeigt Anzahl, Status, Größe und älteste Nachrichten.
    #>
    Write-Log -Message "=== Sammle Message Queue Informationen ===" -Level "INFO"

    $queueHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Message Queues für Server: $server" -Level "INFO"

            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"

                if ($isLocal) {
                    $queues = Get-Queue -ErrorAction SilentlyContinue
                } else {
                    $queues = Get-Queue -Server $server -ErrorAction SilentlyContinue
                }

                if ($queues) {
                    $queueData = foreach ($q in $queues) {
                        $statusIcon = switch ($q.Status) {
                            "Ready"     { "✅" }
                            "Retry"     { "🟡" }
                            "Suspended" { "🔴" }
                            "Active"    { "🔄" }
                            default     { "ℹ️" }
                        }
                        [PSCustomObject]@{
                            "Server"        = $q.Identity.ToString().Split('\')[0]
                            "Queue Name"    = $q.Identity.ToString().Split('\')[1]
                            "Status"        = "$statusIcon $($q.Status)"
                            "Nachrichten"   = $q.MessageCount
                            "Nächster Versuch" = if ($q.NextHopConnector) { $q.NextHopConnector.ToString().Substring(0, 8) + "..." } else { "-" }
                            "Delivery Type"  = $q.DeliveryType
                            "Risk Level"    = $q.RiskLevel
                        }
                    }
                    $serverHTML += "<h4>Transport-Warteschlangen</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $queueData)

                    # Zusammenfassung
                    $totalMessages = ($queues | Measure-Object -Property MessageCount -Sum).Sum
                    $retryQueues = ($queues | Where-Object { $_.Status -eq "Retry" }).Count
                    $suspendedQueues = ($queues | Where-Object { $_.Status -eq "Suspended" }).Count
                    $serverHTML += "<p><strong>Gesamt Nachrichten in Warteschlangen:</strong> $totalMessages | "
                    $serverHTML += "<strong>Retry-Queues:</strong> $retryQueues | "
                    $serverHTML += "<strong>Suspended-Queues:</strong> $suspendedQueues</p>"

                    if ($retryQueues -gt 0 -or $suspendedQueues -gt 0) {
                        $serverHTML += "<p class='error'>⚠️ Es gibt Warteschlangen mit Problemen (Retry/Suspended)! Bitte prüfen.</p>"
                    }
                }
                else {
                    $serverHTML += "<p class='no-data'>Keine Transport-Warteschlangen gefunden oder Transportdienst nicht verfügbar.</p>"
                }
            }
            catch {
                Write-Log -Message "Queue-Abfrage für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Queue-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $queueHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Fehler bei Message Queue für ${server}: $_" -Level "ERROR"
            $queueHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Message Queue Analyse" -Content $queueHTML
}

# ---------------------------------------------------------------
# 37. CALENDAR & RESOURCE MAILBOX KONFIGURATION
# ---------------------------------------------------------------
function Get-CalendarResourceConfigInfo {
    <#
    .SYNOPSIS
        Dokumentiert Raum- und Ressourcenpostfächer sowie deren Buchungsoptionen.
    #>
    Write-Log -Message "=== Sammle Calendar & Resource Mailbox Konfiguration ===" -Level "INFO"

    $calHTML = ""

    try {
        # Raumpostfächer
        $roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
        if ($roomMailboxes) {
            $roomData = foreach ($rm in $roomMailboxes) {
                $calProc = Get-CalendarProcessing -Identity $rm.Identity -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    "Anzeigename"           = $rm.DisplayName
                    "Alias"                 = $rm.Alias
                    "Datenbank"             = $rm.Database
                    "Raumkapazität"         = $rm.ResourceCapacity
                    "Automatische Buchung"  = if ($calProc -and $calProc.AutomateProcessing -eq "AutoAccept") { "✅ Ja" } else { "❌ Nein" }
                    "Buchungsgenehmigung"   = if ($calProc) { $calProc.BookingWindowInDays } else { "-" }
                    "Max. Dauer (Tage)"     = if ($calProc) { $calProc.MaximumDurationInMinutes / 1440 } else { "-" }
                    "Nur ab/bis"            = if ($calProc) { "$($calProc.AllowConflicts)" } else { "-" }
                    "Zulässige Buchungen"   = if ($calProc) { $calProc.BookInPolicy } else { "-" }
                }
            }
            $calHTML += "<h3>Raumpostfächer (Room Mailboxes)</h3>"
            $calHTML += "<p><strong>Anzahl:</strong> $($roomMailboxes.Count)</p>"
            $calHTML += (ConvertTo-HTMLTable -Data $roomData)
        }
        else {
            $calHTML += "<h3>Raumpostfächer</h3>"
            $calHTML += "<p class='no-data'>Keine Raumpostfächer vorhanden.</p>"
        }

        # Ressourcenpostfächer (Equipment)
        $equipMailboxes = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
        if ($equipMailboxes) {
            $equipData = foreach ($em in $equipMailboxes) {
                $calProc = Get-CalendarProcessing -Identity $em.Identity -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    "DisplayName"           = $em.DisplayName
                    "Alias"                 = $em.Alias
                    "Datenbank"             = $em.Database
                    "Automatische Buchung"  = if ($calProc -and $calProc.AutomateProcessing -eq "AutoAccept") { "✅ Ja" } else { "❌ Nein" }
                    "Buchungsgenehmigung"   = if ($calProc) { $calProc.BookingWindowInDays } else { "-" }
                    "Max. Dauer (Tage)"     = if ($calProc) { [math]::Round($calProc.MaximumDurationInMinutes / 1440, 1) } else { "-" }
                }
            }
            $calHTML += "<h3>Ressourcenpostfächer (Equipment Mailboxes)</h3>"
            $calHTML += "<p><strong>Anzahl:</strong> $($equipMailboxes.Count)</p>"
            $calHTML += (ConvertTo-HTMLTable -Data $equipData)
        }
        else {
            $calHTML += "<h3>Ressourcenpostfächer</h3>"
            $calHTML += "<p class='no-data'>Keine Ressourcenpostfächer vorhanden.</p>"
        }

        # Calendar Processing Defaults
        try {
            $calHTML += "<h3>Calendar Processing Standard-Konfiguration</h3>"
            $calHTML += "<p><em>Die CalendarProcessing-Einstellungen steuern, wie Raum-/Ressourcenpostfächer Buchungen verarbeiten:</em></p>"
            $calHTML += "<table><thead><tr><th>Einstellung</th><th>Beschreibung</th></tr></thead><tbody>"
            $calHTML += "<tr class='even'><td>AutomateProcessing</td><td>AutoAccept = Automatische Buchung, AutoUpdate = Nur Update, None = Manuell</td></tr>"
            $calHTML += "<tr class='odd'><td>BookingWindowInDays</td><td>Zeitraum in Tagen für den Buchungen möglich sind (Standard: 180)</td></tr>"
            $calHTML += "<tr class='even'><td>MaximumDurationInMinutes</td><td>Maximale Dauer einer Buchung (Standard: 1440 = 24h)</td></tr>"
            $calHTML += "<tr class='odd'><td>AllowConflicts</td><td>Dürfen sich Buchungen überschneiden? (Standard: False)</td></tr>"
            $calHTML += "<tr class='even'><td>BookInPolicy</td><td>Nur autorisierte Benutzer dürfen buchen (Standard: True)</td></tr>"
            $calHTML += "<tr class='odd'><td>AllBookInPolicy</td><td>Alle Benutzer dürfen buchen (Standard: False)</td></tr>"
            $calHTML += "<tr class='even'><td>EnforceSchedulingBehavior</td><td>Buchungsrichtlinien erzwingen (Standard: True)</td></tr>"
            $calHTML += "</tbody></table>"
        }
        catch {
            Write-Log -Message "CalendarProcessing Defaults fehlgeschlagen: $_" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Fehler bei Calendar/Resource Mailbox Abfrage: $_" -Level "ERROR"
        $calHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Calendar & Resource Mailbox Konfiguration" -Content $calHTML
}

# ---------------------------------------------------------------
# 38. EXCHANGE ARCHIVE KONFIGURATION
# ---------------------------------------------------------------
function Get-ArchiveConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Archivpostfach-Konfiguration: aktivierte Archive,
        Archiv-Quotas, Auto-Expanding Archive und Archiv-Status.
    #>
    Write-Log -Message "=== Sammle Exchange Archive Konfiguration ===" -Level "INFO"

    $archiveHTML = ""

    try {
        # Alle Postfächer mit Archiv
        $archivedMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Where-Object { $_.ArchiveDatabase -ne $null -or $_.ArchiveStatus -eq "Active" }

        if ($archivedMailboxes) {
            $archiveData = foreach ($am in $archivedMailboxes) {
                $archiveStats = Get-MailboxStatistics -Identity $am.Identity -Archive -ErrorAction SilentlyContinue
                $archiveSize = if ($archiveStats) { [math]::Round($archiveStats.TotalItemSize.Value.ToBytes() / 1GB, 2) } else { "N/A" }
                $archiveItems = if ($archiveStats) { $archiveStats.ItemCount } else { "N/A" }

                [PSCustomObject]@{
                    "DisplayName"           = $am.DisplayName
                    "Archiv-Datenbank"      = $am.ArchiveDatabase
                    "Archiv-Status"         = $am.ArchiveStatus
                    "Archiv-Quota (GB)"      = if ($am.ArchiveQuota) { [math]::Round($am.ArchiveQuota.Value.ToBytes() / 1GB, 2) } else { "Nicht gesetzt" }
                    "Archiv-Warnung (GB)"   = if ($am.ArchiveWarningQuota) { [math]::Round($am.ArchiveWarningQuota.Value.ToBytes() / 1GB, 2) } else { "Nicht gesetzt" }
                    "Auto-Expanding"          = if ($am.AutoExpandingArchiveEnabled) { "✅ Ja" } else { "❌ Nein" }
                    "Archiv-Größe (GB)"        = $archiveSize
                    "Archiv-Items"          = $archiveItems
                }
            }

            $archiveHTML += "<h3>Postfächer mit aktiviertem Archiv</h3>"
            $archiveHTML += "<p><strong>Anzahl:</strong> $($archivedMailboxes.Count) von $((Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue).Count) Postfächern</p>"
            $archiveHTML += (ConvertTo-HTMLTable -Data $archiveData)
        }
        else {
            $archiveHTML += "<h3>Archivpostfächer</h3>"
            $archiveHTML += "<p class='no-data'>Keine Postfächer mit aktiviertem Archiv gefunden.</p>"
        }

        # Archiv-Datenbanken
        try {
            $archiveDatabases = Get-MailboxDatabase -ErrorAction SilentlyContinue |
                Where-Object { $_.IsExcludedFromProvisioning -eq $false -and $_.Name -like "*Archive*" }

            if (-not $archiveDatabases) {
                # Fallback: Alle DBs, die als Archiv markiert sind
                $archiveDatabases = Get-MailboxDatabase -ErrorAction SilentlyContinue |
                    Where-Object { $_.IsArchive }
            }

            if ($archiveDatabases) {
                $adbData = foreach ($adb in $archiveDatabases) {
                    [PSCustomObject]@{
                        "Name"              = $adb.Name
                        "Server"            = $adb.Server
                        "EdbFilePath"       = $adb.EdbFilePath
                        "LogFolderPath"     = $adb.LogFolderPath
                        "Verfügbarer Platz" = if ($adb.AvailableNewMailboxSpace) { [math]::Round($adb.AvailableNewMailboxSpace.Value.ToBytes() / 1GB, 2) } else { "N/A" }
                    }
                }
                $archiveHTML += "<h3>Archiv-Datenbanken</h3>"
                $archiveHTML += (ConvertTo-HTMLTable -Data $adbData)
            }
        }
        catch {
            Write-Log -Message "Archiv-Datenbank Abfrage fehlgeschlagen: $_" -Level "WARNING"
        }

        # Auto-Expanding Archive Übersicht
        try {
            $autoExpandMailboxes = $archivedMailboxes | Where-Object { $_.AutoExpandingArchiveEnabled -eq $true }
            if ($autoExpandMailboxes) {
                $archiveHTML += "<h3>Auto-Expanding Archive</h3>"
                $archiveHTML += "<p><strong>Auto-Expanding Archive aktiviert für:</strong> $($autoExpandMailboxes.Count) Postfächer</p>"
                $archiveHTML += "<p><em>Auto-Expanding Archive erlauben Postfächern > 100 GB Archiv-Größe durch automatische "`
                    + "Hinzufügung zusätzlicher Archiv-Datenbanken.</em></p>"
            }
        }
        catch {
            Write-Log -Message "Auto-Expanding Check fehlgeschlagen: $_" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Fehler bei Archive-Konfiguration: $_" -Level "ERROR"
        $archiveHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange Archive Konfiguration" -Content $archiveHTML
}

# ---------------------------------------------------------------
# 39. EXCHANGE MESSAGE SIZE LIMITS
# ---------------------------------------------------------------
function Get-MessageSizeLimitsInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Nachrichtengrößen-Limits auf Organisations-, Connector-,
        Remote-Domain- und Benutzerebene.
    #>
    Write-Log -Message "=== Sammle Exchange Message Size Limits ===" -Level "INFO"

    $limitsHTML = ""

    try {
        # Organisations-Level (TransportConfig)
        try {
            $tc = Get-TransportConfig -ErrorAction SilentlyContinue
            if ($tc) {
                $orgLimits = [PSCustomObject]@{
                    "Max Send Size (Org)"        = $tc.MaxSendSize
                    "Max Receive Size (Org)"    = $tc.MaxReceiveSize
                    "Max Recipient Limit (Org)" = $tc.MaxRecipientEnvelopeLimit
                }
                $limitsHTML += "<h3>Organisations-Level Limits (TransportConfig)</h3>"
                $limitsHTML += (ConvertTo-HTMLTable -Data @($orgLimits))
            }
        }
        catch {
            Write-Log -Message "TransportConfig Limits fehlgeschlagen: $_" -Level "WARNING"
        }

        # Connector-Level (Send Connectors)
        try {
            $sendConnectors = Get-SendConnector -ErrorAction SilentlyContinue
            if ($sendConnectors) {
                $scLimits = foreach ($sc in $sendConnectors) {
                    [PSCustomObject]@{
                        "Connector Name"    = $sc.Name
                        "Typ"               = "Send Connector"
                        "Max Message Size"  = $sc.MaxMessageSize
                    }
                }
                $limitsHTML += "<h3>Send Connector Limits</h3>"
                $limitsHTML += (ConvertTo-HTMLTable -Data $scLimits)
            }
        }
        catch {
            Write-Log -Message "SendConnector Limits fehlgeschlagen: $_" -Level "WARNING"
        }

        # Connector-Level (Receive Connectors)
        try {
            $recvConnectors = Get-ReceiveConnector -ErrorAction SilentlyContinue
            if ($recvConnectors) {
                $rcLimits = foreach ($rc in $recvConnectors) {
                    [PSCustomObject]@{
                        "Connector Name"    = $rc.Name
                        "Server"            = $rc.Server
                        "Typ"               = "Receive Connector"
                        "Max Message Size"  = $rc.MaxMessageSize
                        "Max Recipients"    = $rc.MaxRecipientsPerMessage
                    }
                }
                $limitsHTML += "<h3>Receive Connector Limits</h3>"
                $limitsHTML += (ConvertTo-HTMLTable -Data $rcLimits)
            }
        }
        catch {
            Write-Log -Message "ReceiveConnector Limits fehlgeschlagen: $_" -Level "WARNING"
        }

        # Remote Domain Limits
        try {
            $remoteDomains = Get-RemoteDomain -ErrorAction SilentlyContinue
            if ($remoteDomains) {
                $rdLimits = foreach ($rd in $remoteDomains) {
                    [PSCustomObject]@{
                        "Domain Name"           = $rd.Name
                        "Max Message Size"     = if ($rd.MaxMessageSize) { $rd.MaxMessageSize } else { "Organisations-Standard" }
                        "Auto Reply Enabled"    = $rd.AutoReplyEnabled
                        "Auto Forward Enabled"  = $rd.AutoForwardEnabled
                    }
                }
                $limitsHTML += "<h3>Remote Domain Limits</h3>"
                $limitsHTML += (ConvertTo-HTMLTable -Data $rdLimits)
            }
        }
        catch {
            Write-Log -Message "RemoteDomain Limits fehlgeschlagen: $_" -Level "WARNING"
        }

        # Mailbox-Level Limits (Top 50 größten Limits)
        try {
            $mbxLimits = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
                Where-Object { $_.MaxSendSize -ne $null -or $_.MaxReceiveSize -ne $null } |
                Select-Object -First 50 -Property DisplayName,
                    @{N='Max Send Size';E={$_.MaxSendSize}},
                    @{N='Max Receive Size';E={$_.MaxReceiveSize}},
                    @{N='ProhibitSendQuota';E={$_.ProhibitSendQuota}},
                    @{N='MaxRecipientsPerMessage';E={$_.MaxRecipientsPerMessage}}

            if ($mbxLimits) {
                $limitsHTML += "<h3>Postfach-Individuelle Limits (Top 50)</h3>"
                $limitsHTML += "<p><em>Nur Postfächer mit abweichenden Limits (nicht Organisations-Standard)</em></p>"
                $limitsHTML += (ConvertTo-HTMLTable -Data $mbxLimits)
            }
        }
        catch {
            Write-Log -Message "Mailbox Limits fehlgeschlagen: $_" -Level "WARNING"
        }

        # Zusammenfassung
        $limitsHTML += "<h3>Best Practice Empfehlungen</h3>"
        $limitsHTML += "<table><thead><tr><th>Limit-Typ</th><th>Empfohlener Wert</th><th>Begründung</th></tr></thead><tbody>"
        $limitsHTML += "<tr class='even'><td>Max Send Size (Org)</td><td>25-35 MB</td><td>Kompatibilität mit externen Systemen</td></tr>"
        $limitsHTML += "<tr class='odd'><td>Max Receive Size (Org)</td><td>35-50 MB</td><td>Größere E-Mails werden oft benötigt</td></tr>"
        $limitsHTML += "<tr class='even'><td>Max Recipient Limit</td><td>500-1000</td><td>Spam- und Massenmail-Schutz</td></tr>"
        $limitsHTML += "<tr class='odd'><td>Send Connector Limit</td><td>10-35 MB</td><td>Abhängig vom Zielsystem</td></tr>"
        $limitsHTML += "</tbody></table>"
    }
    catch {
        Write-Log -Message "Fehler bei Message Size Limits: $_" -Level "ERROR"
        $limitsHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange Message Size Limits" -Content $limitsHTML
}

# ---------------------------------------------------------------
# 40. EXCHANGE PARTNER APPLICATIONS
# ---------------------------------------------------------------
function Get-PartnerApplicationsInfo {
    <#
    .SYNOPSIS
        Dokumentiert Partner Applications (z.B. SharePoint, Skype/Lync, CRM).
    #>
    Write-Log -Message "=== Sammle Exchange Partner Applications ===" -Level "INFO"

    $paHTML = ""

    try {
        $partnerApps = Get-PartnerApplication -ErrorAction SilentlyContinue

        if ($partnerApps) {
            $paData = foreach ($pa in $partnerApps) {
                [PSCustomObject]@{
                    "Name"              = $pa.Name
                    "Application ID"    = $pa.ApplicationIdentifier
                    "Aktiviert"         = $pa.Enabled
                    "Realm"             = if ($pa.Realm) { $pa.Realm } else { "-" }
                    "Auth Metadata URL" = if ($pa.AuthMetadataUrl) { $pa.AuthMetadataUrl } else { "-" }
                    "Akteptierte"       = if ($pa.AcceptSecurityIdentifierInformation) { "Ja" } else { "Nein" }
                    "Linked Account"    = if ($pa.LinkedAccount) { $pa.LinkedAccount } else { "-" }
                }
            }
            $paHTML += "<h3>Partner Applications</h3>"
            $paHTML += "<p><em>Partner Applications ermöglichen OAuth-Authentifizierung zwischen Exchange und anderen Diensten.</em></p>"
            $paHTML += (ConvertTo-HTMLTable -Data $paData)
        }
        else {
            $paHTML += "<h3>Partner Applications</h3>"
            $paHTML += "<p class='no-data'>Keine Partner Applications konfiguriert.</p>"
        }
    }
    catch {
        Write-Log -Message "Fehler bei Partner Applications: $_" -Level "ERROR"
        $paHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Exchange Partner Applications" -Content $paHTML
}

# ---------------------------------------------------------------
# 41. EXCHANGE FEDERATED SHARING
# ---------------------------------------------------------------
function Get-FederatedSharingInfo {
    <#
    .SYNOPSIS
        Dokumentiert Federation Trust, Organization Relationships und
        Federated Sharing-Konfiguration.
    #>
    Write-Log -Message "=== Sammle Exchange Federated Sharing ===" -Level "INFO"

    $fedHTML = ""

    # --- Federation Trust ---
    try {
        $fedTrust = Get-FederationTrust -ErrorAction SilentlyContinue
        if ($fedTrust) {
            $ftData = foreach ($ft in $fedTrust) {
                [PSCustomObject]@{
                    "Name"                  = $ft.Name
                    "Org Prerequisit"       = $ft.OrgPrerequisites
                    "Account Namespace"     = $ft.AccountNamespace
                    "Token Issuer URI"      = $ft.TokenIssuerUri
                    "Status"                 = $ft.Status
                    "Application URI"       = $ft.ApplicationUri
                    "Identifier"            = $ft.Identity
                }
            }
            $fedHTML += "<h3>Federation Trust</h3>"
            $fedHTML += (ConvertTo-HTMLTable -Data $ftData)
        }
        else {
            $fedHTML += "<h3>Federation Trust</h3>"
            $fedHTML += "<p class='no-data'>Kein Federation Trust konfiguriert.</p>"
        }
    }
    catch {
        Write-Log -Message "FederationTrust-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Organization Relationships ---
    try {
        $orgRelationships = Get-OrganizationRelationship -ErrorAction SilentlyContinue
        if ($orgRelationships) {
            $orData = foreach ($or in $orgRelationships) {
                [PSCustomObject]@{
                    "Name"                  = $or.Name
                    "Domain Names"          = ($or.DomainNames -join ', ')
                    "Enabled"               = $or.Enabled
                    "FreeBusy Access"       = $or.FreeBusyAccessEnabled
                    "FreeBusy Level"        = $or.FreeBusyAccessLevel
                    "MailTips Access"       = $or.MailTipsAccessEnabled
                    "Calendar Sharing"      = $or.MailboxMoveEnabled
                    "Target Autodiscover"   = $or.TargetAutodiscoverEpr
                }
            }
            $fedHTML += "<h3>Organization Relationships (Frei/Gebucht-Freigabe)</h3>"
            $fedHTML += (ConvertTo-HTMLTable -Data $orData)
        }
        else {
            $fedHTML += "<h3>Organization Relationships</h3>"
            $fedHTML += "<p class='no-data'>Keine Organization Relationships konfiguriert.</p>"
        }
    }
    catch {
        Write-Log -Message "OrganizationRelationship-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Federated Domain Proof ---
    try {
        if ($fedTrust) {
            $fedDomainProof = Get-FederatedOrganizationIdentifier -ErrorAction SilentlyContinue
            if ($fedDomainProof) {
                $fdpData = [PSCustomObject]@{
                    "Federated Domains"         = ($fedDomainProof.Domains -join ', ')
                    "Enabled"                   = $fedDomainProof.Enabled
                    "Account Namespace"         = $fedDomainProof.AccountNamespace
                    "Delegation Trust"          = $fedDomainProof.DelegationTrustLink
                }
                $fedHTML += "<h3>Federated Organization Identifier</h3>"
                $fedHTML += (ConvertTo-HTMLTable -Data @($fdpData))
            }
        }
    }
    catch {
        Write-Log -Message "FederatedOrganizationIdentifier-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Hinweis ---
    $fedHTML += "<h3>Hinweis zu Federation</h3>"
    $fedHTML += "<p><em>Federation wird für folgende Szenarien benötigt:</em></p>"
    $fedHTML += "<ul>"
    $fedHTML += "<li>Frei/Gebucht-Informationen mit externen Organisationen teilen</li>"
    $fedHTML += "<li>Hybrid-Konfiguration mit Exchange Online (OAuth)</li>"
    $fedHTML += "<li>Postfach-Migrationen zwischen Organisationen</li>"
    $fedHTML += "<li>Kalender-Freigabe mit anderen Exchange-Organisationen</li>"
    $fedHTML += "</ul>"

    New-HTMLSection -Title "Exchange Federated Sharing" -Content $fedHTML
}

# ---------------------------------------------------------------
# 42. OAUTH / CERTIFICATE BASED AUTH
# ---------------------------------------------------------------
function Get-OAuthCertificateAuthInfo {
    <#
    .SYNOPSIS
        Dokumentiert die OAuth-Konfiguration, Auth Server, Zertifikate
        und Certificate-Based Authentication (CBA) für Exchange.
    #>
    Write-Log -Message "=== Sammle OAuth / Certificate Based Auth ===" -Level "INFO"

    $oauthHTML = ""

    # --- Auth Server ---
    try {
        $authServers = Get-AuthServer -ErrorAction SilentlyContinue
        if ($authServers) {
            $asData = foreach ($as in $authServers) {
                [PSCustomObject]@{
                    "Name"              = $as.Name
                    "Typ"               = $as.Type
                    "Enabled"           = $as.Enabled
                    "Auth Metadata URL" = if ($as.AuthMetadataUrl) { $as.AuthMetadataUrl } else { "-" }
                    "Realm"             = if ($as.Realm) { $as.Realm } else { "-" }
                    "Issuer Identifier" = if ($as.IssuerIdentifier) { $as.IssuerIdentifier } else { "-" }
                }
            }
            $oauthHTML += "<h3>Auth Server (OAuth-Konfiguration)</h3>"
            $oauthHTML += (ConvertTo-HTMLTable -Data $asData)
        }
        else {
            $oauthHTML += "<h3>Auth Server</h3>"
            $oauthHTML += "<p class='no-data'>Keine Auth Server konfiguriert.</p>"
        }
    }
    catch {
        Write-Log -Message "AuthServer-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Auth Config (OAuth-Zertifikate) ---
    try {
        $authConfig = Get-AuthConfig -ErrorAction SilentlyContinue
        if ($authConfig) {
            $acData = [PSCustomObject]@{
                "Name"                      = $authConfig.Name
                "Service Name"              = $authConfig.ServiceName
                "Realm"                     = $authConfig.Realm
                "Current Certificate"       = if ($authConfig.CurrentCertificateThumbprint) { $authConfig.CurrentCertificateThumbprint } else { "-" }
                "Previous Certificate"         = if ($authConfig.PreviousCertificateThumbprint) { $authConfig.PreviousCertificateThumbprint } else { "-" }
                "Next Certificate"           = if ($authConfig.NextCertificateThumbprint) { $authConfig.NextCertificateThumbprint } else { "-" }
                "Certificate Rollover"    = $authConfig.CertificateRolloverEnabled
            }
            $oauthHTML += "<h3>Auth Config (OAuth-Zertifikate)</h3>"
            $oauthHTML += (ConvertTo-HTMLTable -Data @($acData))

            # Zertifikat-Details
            $currentThumb = $authConfig.CurrentCertificateThumbprint
            if ($currentThumb) {
                try {
                    $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$currentThumb" -ErrorAction SilentlyContinue
                    if ($cert) {
                        $certInfo = [PSCustomObject]@{
                            "Subject"       = $cert.Subject
                            "Issuer"        = $cert.Issuer
                            "Not Before"    = $cert.NotBefore.ToString("dd.MM.yyyy")
                            "Not After"     = $cert.NotAfter.ToString("dd.MM.yyyy")
                            "Thumbprint"    = $cert.Thumbprint
                            "Serial Number" = $cert.SerialNumber
                        }
                        $oauthHTML += "<h4>Aktuelles OAuth-Zertifikat (Details)</h4>"
                        $oauthHTML += (ConvertTo-HTMLTable -Data @($certInfo))
                    }
                }
                catch {
                    Write-Log -Message "OAuth-Zertifikat nicht im lokalen Zertifikatsspeicher: $_" -Level "WARNING"
                }
            }
        }
    }
    catch {
        Write-Log -Message "AuthConfig-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Intra-Organization Connectors ---
    try {
        $intraOrg = Get-IntraOrganizationConnector -ErrorAction SilentlyContinue
        if ($intraOrg) {
            $iocData = foreach ($ioc in $intraOrg) {
                [PSCustomObject]@{
                    "Name"                  = $ioc.Name
                    "Target Address"        = ($ioc.TargetAddressDomains -join ', ')
                    "Discovery Endpoint"    = $ioc.DiscoveryEndpoint
                    "Enabled"               = $ioc.Enabled
                }
            }
            $oauthHTML += "<h3>Intra-Organization Connectors (OAuth)</h3>"
            $oauthHTML += (ConvertTo-HTMLTable -Data $iocData)
        }
    }
    catch {
        Write-Log -Message "IntraOrganizationConnector-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- OAuth 2.0 Client Profile ---
    try {
        $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($orgConfig) {
            $oauthProfile = [PSCustomObject]@{
                "OAuth2 Client Profile Enabled" = $orgConfig.OAuth2ClientProfileEnabled
                "OAuth2 Client Profile"           = if ($orgConfig.OAuth2ClientProfileEnabled) { "✅ Aktiviert" } else { "❌ Deaktiviert" }
            }
            $oauthHTML += "<h3>OAuth 2.0 Client Profile (Modern Auth)</h3>"
            $oauthHTML += (ConvertTo-HTMLTable -Data @($oauthProfile))
            if (-not $orgConfig.OAuth2ClientProfileEnabled) {
                $oauthHTML += "<p class='error'>⚠️ OAuth 2.0 ist nicht aktiviert! Dies ist für moderne Clients (Outlook 2016+, Mobile) und Hybrid-Szenarien erforderlich.</p>"
            }
        }
    }
    catch {
        Write-Log -Message "OAuth2ClientProfile-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Certificate Based Authentication (CBA) ---
    try {
        $cbaEnabled = $false
        $owaCba = Get-OwaVirtualDirectory -ErrorAction SilentlyContinue |
            Where-Object { $_.ClientAuthCleanupLevel -ne $null -or $_.IISAuthenticationMethods -match "Certificate" }
        if ($owaCba) {
            $cbaData = foreach ($owa in $owaCba) {
                [PSCustomObject]@{
                    "Server"                    = $owa.Server
                    "ClientAuthCleanupLevel"    = if ($owa.ClientAuthCleanupLevel) { $owa.ClientAuthCleanupLevel } else { "Nicht gesetzt" }
                    "IIS Auth Methods"          = ($owa.IISAuthenticationMethods -join ', ')
                }
            }
            $oauthHTML += "<h3>Certificate Based Authentication (CBA) - OWA</h3>"
            $oauthHTML += (ConvertTo-HTMLTable -Data $cbaData)
            $cbaEnabled = $true
        }

        # IIS Global CBA Check
        try {
            $isLocal = $env:COMPUTERNAME
            $iisConfigPath = "$env:windir\System32\inetsrv\config\applicationHost.config"
            if (Test-Path $iisConfigPath) {
                [xml]$iisXml = Get-Content $iisConfigPath -ErrorAction SilentlyContinue
                if ($iisXml) {
                    $oauthHTML += "<h3>IIS SSL-Einstellungen (Certificate Mapping)</h3>"
                    $oauthHTML += "<p><em>Für detaillierte CBA-Konfiguration prüfen Sie die IIS-Manager-Konsole "
                    $oauthHTML += "unter 'SSL-Einstellungen' > 'Clientzertifikate' > 'Erforderlich' oder 'Akzeptieren'.</em></p>"
                }
            }
        }
        catch {
            Write-Log -Message "IIS CBA-Check fehlgeschlagen: $_" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "CBA-Abfrage fehlgeschlagen: $_" -Level "WARNING"
    }

    # --- Zusammenfassung ---
    $oauthHTML += "<h3>OAuth/CBA Best Practices</h3>"
    $oauthHTML += "<table><thead><tr><th>Bereich</th><th>Empfehlung</th></tr></thead><tbody>"
    $oauthHTML += "<tr class='even'><td>OAuth 2.0</td><td>Aktivieren für moderne Clients und Hybrid</td></tr>"
    $oauthHTML += "<tr class='odd'><td>Auth-Zertifikat</td><td>Vor Ablauf erneuern (min. 1 Jahr gültig)</td></tr>"
    $oauthHTML += "<tr class='even'><td>Certificate Based Auth</td><td>Für sichere Client-Zugriffe ohne Passwort</td></tr>"
    $oauthHTML += "<tr class='odd'><td>Intra-Org Connector</td><td>Für OAuth-Kommunikation zwischen Exchange und SharePoint/Skype</td></tr>"
    $oauthHTML += "</tbody></table>"

    New-HTMLSection -Title "OAuth / Certificate Based Auth" -Content $oauthHTML
}

#endregion

#region ============================================================
# NEU IN v1.7 - ERWEITERTE SYSTEM- & SICHERHEITSCHECKS
#endregion ============================================================

# ---------------------------------------------------------------
# 43. WINDOWS FEATURES & ROLLEN
# ---------------------------------------------------------------
function Get-WindowsFeaturesInfo {
    <#
    .SYNOPSIS
        Dokumentiert installierte Windows Server Rollen und Features.
    #>
    Write-Log -Message "=== Sammle Windows Features & Rollen ===" -Level "INFO"

    $featHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Windows Features für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                $features = Get-CimInstance -CimSession $session -ClassName Win32_ServerFeature -ErrorAction SilentlyContinue
                if ($features) {
                    $featData = $features | Select-Object @{N='ID';E={$_.ID}}, @{N='Name';E={$_.Name}} | Sort-Object Name
                    $serverHTML += "<h4>Installierte Windows Server Rollen</h4>"
                    $serverHTML += "<p><strong>Anzahl:</strong> $($featData.Count)</p>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $featData)
                } else {
                    try {
                        $dismFeat = Get-WindowsFeature -ComputerName $server -ErrorAction SilentlyContinue | Where-Object { $_.Installed }
                        if ($dismFeat) {
                            $featData = $dismFeat | Select-Object @{N='Name';E={$_.Name}}, @{N='DisplayName';E={$_.DisplayName}} | Sort-Object Name
                            $serverHTML += "<h4>Installierte Windows Features (DISM)</h4>"
                            $serverHTML += "<p><strong>Anzahl:</strong> $($featData.Count)</p>"
                            $serverHTML += (ConvertTo-HTMLTable -Data $featData)
                        } else { $serverHTML += "<p class='no-data'>Keine Feature-Informationen verfügbar.</p>" }
                    } catch { $serverHTML += "<p class='no-data'>Feature-Informationen nicht verfügbar (CIM/DISM).</p>" }
                }
                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "Windows Features für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Windows Features konnten nicht abgerufen werden: $_</p>"
            }
            $featHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei Windows Features für ${server}: $_" -Level "ERROR"
            $featHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Windows Features & Rollen" -Content $featHTML
}

# ---------------------------------------------------------------
# 44. .NET FRAMEWORK VERSION & DLLS
# ---------------------------------------------------------------
function Get-DotNetFrameworkInfo {
    <#
    .SYNOPSIS
        Dokumentiert die installierte .NET Framework Version und relevante DLL-Versionen.
    #>
    Write-Log -Message "=== Sammle .NET Framework Informationen ===" -Level "INFO"

    $dotnetHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message ".NET Framework für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                # .NET Version aus Registry via CIM
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                $dotNetVersion = Get-CimInstance -CimSession $session -ClassName Win32_Registry |
                    Where-Object { $_.Key -eq "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" } |
                    Select-Object -ExpandProperty ValueNames -ErrorAction SilentlyContinue

                # Direct Registry access
                $regPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
                $releaseReg = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "Release"
                $versionReg = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $regPath -ValueName "Version"

                $dotNetReleaseMap = @{
                    528040 = "4.8 (Build 528040)"
                    528049 = "4.8 (Build 528049)"
                    528209 = "4.8 (Build 528209)"
                    528372 = "4.8 (Build 528372)"
                    528449 = "4.8 (Build 528449)"
                    533320 = "4.8.1"
                    533325 = "4.8.1 (Build 533325)"
                    9186481 = "4.8.1 (Windows 11)"
                }

                $dotNetDisplay = if ($releaseReg -and $dotNetReleaseMap.ContainsKey($releaseReg)) {
                    $dotNetReleaseMap[$releaseReg]
                } elseif ($versionReg) { $versionReg } else { "Nicht erkannt" }
                $dotNetRelease = if ($releaseReg) { $releaseReg } else { "-" }

                $netData = [PSCustomObject]@{
                    ".NET Framework Version" = $dotNetDisplay
                    "Release Registry Key"   = $dotNetRelease
                    "Version Registry"       = if ($versionReg) { $versionReg } else { "-" }
                    "CLR Version"            = if ((Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue).BuildNumber) { "4.0.30319" } else { "-" }
                }
                $serverHTML += "<h4>.NET Framework Version</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data @($netData))

                # System.Data.dll Version
                try {
                    $systemDataPath = "SOFTWARE\Microsoft\.NETFramework\AssemblyFolders\System.Data"
                    $serverHTML += "<h4>System.Data.dll Pfad</h4>"
                    $serverHTML += "<p>Standard-Speicherort: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Data.dll</p>"
                } catch { }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message ".NET Framework für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>.NET Framework Informationen konnten nicht abgerufen werden: $_</p>"
            }

            # Zusätzlich: Lokale .NET Assembly-Versionen
            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                if ($isLocal) {
                    $assemblies = @("System.Data", "System.Configuration", "System.Web", "System.DirectoryServices")
                    $asmData = foreach ($asm in $assemblies) {
                        try {
                            $a = [System.Reflection.Assembly]::LoadWithPartialName($asm)
                            if ($a) {
                                [PSCustomObject]@{
                                    "Assembly"    = $asm
                                    "Version"     = $a.GetName().Version.ToString()
                                    "Location"    = $a.Location
                                }
                            }
                        } catch { }
                    }
                    if ($asmData) {
                        $serverHTML += "<h4>.NET Assembly Versionen (lokal)</h4>"
                        $serverHTML += (ConvertTo-HTMLTable -Data $asmData)
                    }
                }
            } catch {
                Write-Log -Message "Lokale Assembly-Abfrage fehlgeschlagen: $_" -Level "WARNING"
            }

            $dotnetHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei .NET Framework für ${server}: $_" -Level "ERROR"
            $dotnetHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title ".NET Framework Version & DLLs" -Content $dotnetHTML
}

# ---------------------------------------------------------------
# 45. AUSSTEHENDE NEUSTARTS (PENDING REBOOT)
# ---------------------------------------------------------------
function Get-PendingRebootInfo {
    <#
    .SYNOPSIS
        Prüft auf ausstehende Neustarts pro Server via Registry und WMI.
    #>
    Write-Log -Message "=== Prüfe auf ausstehende Neustarts ===" -Level "INFO"

    $rebootHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Pending Reboot Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"
            $pendingReboot = $false
            $rebootReasons = @()

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # Check 1: HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations
                $pendingOp = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Control\Session Manager" -ValueName "PendingFileRenameOperations"
                if ($pendingOp) {
                    $pendingReboot = $true
                    $rebootReasons += "Ausstehende Datei-Umbenennungen (PendingFileRenameOperations)"
                }

                # Check 2: HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired
                $wuReboot = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -ValueName "RebootRequired"
                if ($wuReboot) {
                    $pendingReboot = $true
                    $rebootReasons += "Windows Update erfordert Neustart"
                }

                # Check 3: HKLM\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts
                $srvMgrReboot = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\ServerManager" -ValueName "CurrentRebootAttempts"
                if ($srvMgrReboot -and $srvMgrReboot -gt 0) {
                    $pendingReboot = $true
                    $rebootReasons += "Server Manager meldet ausstehenden Neustart"
                }

                # Check 4: CBS Reboot (Component Based Servicing)
                $cbsReboot = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" -ValueName "RebootPending"
                if ($cbsReboot) {
                    $pendingReboot = $true
                    $rebootReasons += "Component Based Servicing erfordert Neustart"
                }

                # Check 5: CIM - Win32_ComputerSystem.RebootRequired (Windows 10/2016+)
                try {
                    $cs = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
                    if ($cs -and $cs.PSSerializer) {
                        $rebootReqProp = $cs | Select-Object -Property RebootRequired -ErrorAction SilentlyContinue
                        if ($rebootReqProp -and $rebootReqProp.RebootRequired) {
                            $pendingReboot = $true
                            $rebootReasons += "Win32_ComputerSystem meldet RebootRequired"
                        }
                    }
                } catch { }

                # Check 6: SCCM/ConfigMgr Client
                $ccmReboot = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management" -ValueName "RebootBy" -ErrorAction SilentlyContinue
                if ($ccmReboot) {
                    $pendingReboot = $true
                    $rebootReasons += "ConfigMgr Client erfordert Neustart bis: $ccmReboot"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "Pending Reboot Check für $server fehlgeschlagen: $_" -Level "WARNING"
            }

            $statusIcon = if ($pendingReboot) { "🔴 NEUSTART AUSSTEHEND" } else { "✅ Kein Neustart erforderlich" }
            $serverHTML += "<h4>Status: $statusIcon</h4>"
            if ($pendingReboot -and $rebootReasons.Count -gt 0) {
                $serverHTML += "<ul>"
                foreach ($reason in $rebootReasons) {
                    $serverHTML += "<li>$reason</li>"
                }
                $serverHTML += "</ul>"
                $serverHTML += "<p class='error'>⚠️ Es sind ausstehende Neustarts auf diesem Server vorhanden! Bitte zeitnah durchführen.</p>"
            }

            $rebootHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei Pending Reboot für ${server}: $_" -Level "ERROR"
            $rebootHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Ausstehende Neustarts (Pending Reboot)" -Content $rebootHTML
}

# ---------------------------------------------------------------
# 46. CPU THROTTLING ERKENNUNG
# ---------------------------------------------------------------
function Get-CpuThrottlingInfo {
    <#
    .SYNOPSIS
        Erkennt CPU Throttling durch Vergleich von CurrentClockSpeed mit MaxClockSpeed.
    #>
    Write-Log -Message "=== Prüfe CPU Throttling ===" -Level "INFO"

    $cpuHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "CPU Throttling Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                $processors = Get-CimInstance -CimSession $session -ClassName Win32_Processor -ErrorAction SilentlyContinue
                if ($processors) {
                    $cpuData = foreach ($proc in $processors) {
                        $throttlingRatio = if ($proc.MaxClockSpeed -gt 0) {
                            [math]::Round(($proc.CurrentClockSpeed / $proc.MaxClockSpeed) * 100, 1)
                        } else { "N/A" }
                        $throttlingIcon = if ($throttlingRatio -ne "N/A" -and $throttlingRatio -lt 95) {
                            "⚠️ Throttling aktiv!"
                        } else { "✅ Normal" }

                        [PSCustomObject]@{
                            "CPU"                 = $proc.Name
                            "Device ID"           = $proc.DeviceID
                            "Max Clock Speed (MHz)" = $proc.MaxClockSpeed
                            "Current Clock (MHz)"   = $proc.CurrentClockSpeed
                            "Auslastung (%)"       = $throttlingRatio
                            "Status"              = $throttlingIcon
                            "Cores"               = $proc.NumberOfCores
                            "Logical Processors"  = $proc.NumberOfLogicalProcessors
                        }
                    }

                    $serverHTML += "<h4>CPU Throttling Analyse</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data $cpuData)

                    # Warnung bei Throttling
                    $throttled = $cpuData | Where-Object { $_."Status" -match "Throttling" }
                    if ($throttled) {
                        $serverHTML += "<p class='error'>⚠️ CPU Throttling erkannt! Die aktuelle Taktrate ist geringer als die maximale Taktrate. "
                        $serverHTML += "Mögliche Ursachen: ÜBerhitzung, Power Management, VM-Ressourcen-Limit.</p>"
                    } else {
                        $serverHTML += "<p class='success'>✅ Kein CPU Throttling festgestellt.</p>"
                    }
                } else {
                    $serverHTML += "<p class='no-data'>Keine Prozessor-Informationen verfügbar.</p>"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "CPU Throttling für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>CPU Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $cpuHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei CPU Throttling für ${server}: $_" -Level "ERROR"
            $cpuHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "CPU Throttling Analyse" -Content $cpuHTML
}

# ---------------------------------------------------------------
# 47. VISUAL C++ REDISTRIBUTABLE
# ---------------------------------------------------------------
function Get-VCRedistInfo {
    <#
    .SYNOPSIS
        Dokumentiert installierte Microsoft Visual C++ Redistributable Versionen.
    #>
    Write-Log -Message "=== Sammle Visual C++ Redistributable Versionen ===" -Level "INFO"

    $vcHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "VC++ Redist für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # 64-bit
                $vc64 = Get-CimInstance -CimSession $session -ClassName Win32_Product -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like "*Visual C++*" -or $_.Name -like "*Visual C++*Redistributable*" } |
                    Select-Object Name, Version, Vendor

                # Alternative via Registry (zuverlässiger als Win32_Product)
                try {
                    $vcKeys = @(
                        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                        "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    )
                    $vcItems = @()
                    foreach ($key in $vcKeys) {
                        # Read subkeys via remote registry
                        $subKeys = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $key -ValueName "" -ErrorAction SilentlyContinue
                        # Try via CIM StdRegProv
                        try {
                            $regProv = Get-CimInstance -CimSession $session -ClassName StdRegProv -Namespace root/default -ErrorAction SilentlyContinue
                            if ($regProv) {
                                $enumResult = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName EnumKey -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $key } -ErrorAction SilentlyContinue
                                if ($enumResult -and $enumResult.sNames) {
                                    foreach ($subKey in $enumResult.sNames) {
                                        $displayName = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = "$key\$subKey"; sValueName = "DisplayName" } -ErrorAction SilentlyContinue
                                        $displayVersion = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = "$key\$subKey"; sValueName = "DisplayVersion" } -ErrorAction SilentlyContinue
                                        if ($displayName.sValue -and $displayName.sValue -match "Visual C\+\+") {
                                            $vcItems += [PSCustomObject]@{
                                                "Name"    = $displayName.sValue
                                                "Version" = if ($displayVersion.sValue) { $displayVersion.sValue } else { "-" }
                                                "Arch"    = if ($key -match "WOW6432Node") { "32-bit" } else { "64-bit" }
                                            }
                                        }
                                    }
                                }
                            }
                        } catch { }
                    }

                    if ($vcItems.Count -gt 0) {
                        $serverHTML += "<h4>Installierte Visual C++ Redistributables</h4>"
                        $serverHTML += "<p><strong>Anzahl:</strong> $($vcItems.Count)</p>"
                        $serverHTML += (ConvertTo-HTMLTable -Data ($vcItems | Sort-Object Name))
                    } else {
                        $serverHTML += "<p class='no-data'>Keine Visual C++ Redistributables gefunden (via Registry).</p>"
                    }
                } catch {
                    Write-Log -Message "VC++ Registry-Abfrage fehlgeschlagen: $_" -Level "WARNING"
                    $serverHTML += "<p class='no-data'>VC++ Redistributables konnten nicht abgerufen werden.</p>"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "VC++ für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>VC++ Redistributable Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $vcHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei VC++ für ${server}: $_" -Level "ERROR"
            $vcHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Visual C++ Redistributable Versionen" -Content $vcHTML
}

# ---------------------------------------------------------------
# 48. CREDENTIAL GUARD STATUS
# ---------------------------------------------------------------
function Get-CredentialGuardInfo {
    <#
    .SYNOPSIS
        Prüft den Status von Windows Defender Credential Guard auf den Servern.
    #>
    Write-Log -Message "=== Prüfe Credential Guard Status ===" -Level "INFO"

    $cgHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Credential Guard Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # Check 1: LSA Credential Guard via Registry
                $lsaCfg = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Policies\Microsoft\Windows\DeviceGuard" -ValueName "LsaCfgFlags"

                $cgStatus = switch ($lsaCfg) {
                    0 { "❌ Deaktiviert" }
                    1 { "✅ Aktiviert - UEFI-Sperre aktiv" }
                    2 { "✅ Aktiviert - Ohne UEFI-Sperre" }
                    3 { "✅ Aktiviert - Mit UEFI-Sperre" }
                    default { "❌ Nicht konfiguriert / Deaktiviert" }
                }

                $cgData = [PSCustomObject]@{
                    "Credential Guard Status" = $cgStatus
                    "LsaCfgFlags (Registry)"  = if ($null -ne $lsaCfg) { $lsaCfg } else { "Nicht gesetzt" }
                    "Registry Pfad"           = "HKLM\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard\LsaCfgFlags"
                }
                $serverHTML += "<h4>Credential Guard Konfiguration</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data @($cgData))

                # Check 2: Virtualization Based Security
                try {
                    $vbs = Get-CimInstance -CimSession $session -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction SilentlyContinue
                    if ($vbs) {
                        $vbsData = [PSCustomObject]@{
                            "VirtualizationSecurity" = if ($vbs.VirtualizationBasedSecurityStatus -eq 2) { "✅ Aktiviert" } else { "❌ Deaktiviert" }
                            "CredentialGuard Configured" = if ($vbs.CredentialGuardConfigured) { "✅ Ja" } else { "❌ Nein" }
                        }
                        $serverHTML += "<h4>Virtualization Based Security Details</h4>"
                        $serverHTML += (ConvertTo-HTMLTable -Data @($vbsData))
                    }
                } catch {
                    Write-Log -Message "VBS CIM-Abfrage nicht möglich (Namespace nicht vorhanden): $_" -Level "WARNING"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "Credential Guard für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Credential Guard Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $cgHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei Credential Guard für ${server}: $_" -Level "ERROR"
            $cgHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Credential Guard Status" -Content $cgHTML
}

# ---------------------------------------------------------------
# 49. LOKALE ADMINISTRATOREN
# ---------------------------------------------------------------
function Get-LocalAdminInfo {
    <#
    .SYNOPSIS
        Listet die Mitglieder der lokalen Administratoren-Gruppe pro Server auf.
    #>
    Write-Log -Message "=== Sammle lokale Administratoren ===" -Level "INFO"

    $adminHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Lokale Administratoren für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # CIM: Win32_Group - Administratoren abrufen
                $adminGroup = Get-CimInstance -CimSession $session -ClassName Win32_Group -Filter "Name = 'Administrators' AND Domain = '$server'" -ErrorAction SilentlyContinue
                if (-not $adminGroup) {
                    $adminGroup = Get-CimInstance -CimSession $session -ClassName Win32_Group -Filter "Name = 'Administrators'" -ErrorAction SilentlyContinue
                }

                if ($adminGroup) {
                    $members = Get-CimInstance -CimSession $session -ClassName Win32_GroupUser -ErrorAction SilentlyContinue |
                        Where-Object { $_.GroupComponent -match "Administrators" }

                    $memberData = if ($members) {
                        $memberList = foreach ($member in $members) {
                            $part = $member.PartComponent -replace '.*Name="([^"]+)".*', '$1'
                            $domain = $member.PartComponent -replace '.*Domain="([^"]+)".*', '$1'
                            [PSCustomObject]@{
                                "Mitglied" = "$domain\$part"
                                "Typ"      = if ($member.PartComponent -match "Win32_UserAccount") { "Benutzer" } else { "Gruppe" }
                            }
                        }
                        $memberList | Sort-Object Mitglied
                    } else {
                        @()
                    }

                    $serverHTML += "<h4>Mitglieder der lokalen Administratoren-Gruppe</h4>"
                    if ($memberData.Count -gt 0) {
                        $serverHTML += "<p><strong>Anzahl Mitglieder:</strong> $($memberData.Count)</p>"
                        $serverHTML += (ConvertTo-HTMLTable -Data $memberData)
                    } else {
                        $serverHTML += "<p class='no-data'>Keine Mitglieder gefunden (oder Zugriff verweigert).</p>"
                    }

                    # Zusätzliche Information: Exchange Administratoren prüfen
                    try {
                        $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                        if ($isLocal) {
                            $groupObj = [ADSI]"WinNT://$server/Administrators"
                            $adsiMembers = @($groupObj.Invoke("Members"))
                            $adsiData = foreach ($member in $adsiMembers) {
                                $name = $member.GetType().InvokeMember("Name", 'GetProperty', $null, $member, $null)
                                $adsPath = $member.GetType().InvokeMember("AdsPath", 'GetProperty', $null, $member, $null)
                                [PSCustomObject]@{
                                    "Mitglied (ADSI)" = $name
                                    "ADS Path"         = $adsPath
                                }
                            }
                            if ($adsiData.Count -gt $memberData.Count) {
                                $serverHTML += "<h4>Erweiterte Administratoren-Liste (ADSI)</h4>"
                                $serverHTML += (ConvertTo-HTMLTable -Data $adsiData)
                            }
                        }
                    } catch {
                        Write-Log -Message "ADSI Administratoren-Abfrage fehlgeschlagen: $_" -Level "WARNING"
                    }
                } else {
                    $serverHTML += "<p class='error'>Administratoren-Gruppe konnte nicht gefunden werden.</p>"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "Lokale Administratoren für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Administratoren-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $adminHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei lokalen Administratoren für ${server}: $_" -Level "ERROR"
            $adminHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Lokale Administratoren" -Content $adminHTML
}

# ---------------------------------------------------------------
# 50. DOMAIN TRUSTS & VERSCHLÜSSELUNG
# ---------------------------------------------------------------
function Get-DomainTrustInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Domain Trusts und deren Verschlüsselungstypen.
    #>
    Write-Log -Message "=== Sammle Domain Trusts & Verschlüsselung ===" -Level "INFO"

    $trustHTML = ""

    try {
        # Prüfe ob AD-Modul verfügbar ist
        $adModule = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
        $adModuleLoaded = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue

        if ($adModule -or $adModuleLoaded) {
            if (-not $adModuleLoaded) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

            $trusts = Get-ADTrust -Filter * -ErrorAction SilentlyContinue
            if ($trusts) {
                $trustData = foreach ($trust in $trusts) {
                    $encTypes = switch ($trust.SupportedEncryptionTypes) {
                        0  { "Keine" }
                        1  { "RC4-HMAC" }
                        2  { "AES128-CTS-HMAC-SHA1" }
                        4  { "AES256-CTS-HMAC-SHA1" }
                        8  { "Future Encryption" }
                        3  { "RC4-HMAC, AES128" }
                        5  { "RC4-HMAC, AES256" }
                        7  { "RC4-HMAC, AES128, AES256" }
                        default { "Unbekannt ($($trust.SupportedEncryptionTypes))" }
                    }

                    [PSCustomObject]@{
                        "Name"                     = $trust.Name
                        "Direction"                = $trust.TrustDirection
                        "Type"                     = $trust.TrustType
                        "Attributes"               = $trust.TrustAttributes
                        "Supported Encryption"     = $encTypes
                        "Source"                   = $trust.Source
                        "Target"                   = $trust.Target
                        "ForestTransitive"         = $trust.ForestTransitive
                        "IntraForest"              = $trust.IntraForest
                        "IsTreeParent"             = $trust.IsTreeParent
                        "SID Filtering"            = $trust.SIDFilteringEnabled
                    }
                }
                $trustHTML += "<h3>Domain Trusts</h3>"
                $trustHTML += "<p><strong>Anzahl:</strong> $($trustData.Count)</p>"
                $trustHTML += (ConvertTo-HTMLTable -Data ($trustData | Sort-Object Name))
            } else {
                $trustHTML += "<h3>Domain Trusts</h3>"
                $trustHTML += "<p class='no-data'>Keine Domain Trusts vorhanden (oder Zugriff verweigert).</p>"
            }

            # Domain Info
            try {
                $domain = Get-ADDomain -ErrorAction SilentlyContinue
                if ($domain) {
                    $domainData = [PSCustomObject]@{
                        "DNS Name"              = $domain.DNSRoot
                        "NetBIOS Name"          = $domain.NetBIOSName
                        "Forest"                = $domain.Forest
                        "Domain Mode"           = $domain.DomainMode
                        "PDC Emulator"          = $domain.PDCEmulator
                        "Infrastructure Master" = $domain.InfrastructureMaster
                        "RID Master"            = $domain.RidMasterMaster
                    }
                    $trustHTML += "<h3>Aktuelle Domäne</h3>"
                    $trustHTML += (ConvertTo-HTMLTable -Data @($domainData))
                }
            } catch {
                Write-Log -Message "ADDomain-Abfrage fehlgeschlagen: $_" -Level "WARNING"
            }
        } else {
            $trustHTML += "<h3>Domain Trusts</h3>"
            $trustHTML += "<p class='no-data'>Active Directory PowerShell-Modul nicht verfügbar. Domain Trusts können nicht abgefragt werden.</p>"
        }
    } catch {
        Write-Log -Message "Fehler bei Domain Trusts: $_" -Level "ERROR"
        $trustHTML += "<p class='error'>Fehler: $_</p>"
    }

    New-HTMLSection -Title "Domain Trusts & Verschlüsselung" -Content $trustHTML
}

# ---------------------------------------------------------------
# 51. FIP-FS SCAN ENGINE VERSION
# ---------------------------------------------------------------
function Get-FIPFSInfo {
    <#
    .SYNOPSIS
        Dokumentiert die FIP-FS (Forefront Identity Protection) Anti-Malware Scan Engine.
    #>
    Write-Log -Message "=== Sammle FIP-FS Scan Engine Informationen ===" -Level "INFO"

    $fipfsHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "FIP-FS Scan Engine für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"

                # Exchange Install Path ermitteln
                $exchInstallPath = if ($isLocal) {
                    (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -Name "MsiInstallPath" -ErrorAction SilentlyContinue).MsiInstallPath
                } else {
                    Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ValueName "MsiInstallPath"
                }

                if ($exchInstallPath) {
                    $fipfsPath = Join-Path -Path $exchInstallPath -ChildPath "TransportRoles\Data\AntiSpam\Engine"
                    $scanEnginePath = Join-Path -Path $fipfsPath -ChildPath "ScanEngine.ini"
                    $updatePath = Join-Path -Path $fipfsPath -ChildPath "Update"

                    $serverHTML += "<h4>FIP-FS Scan Engine</h4>"

                    # ScanEngine.ini auslesen
                    if ($isLocal -and (Test-Path $scanEnginePath)) {
                        try {
                            $iniContent = Get-Content $scanEnginePath -ErrorAction SilentlyContinue
                            $engineVersion = ($iniContent | Select-String -Pattern "^EngineVersion=" -ErrorAction SilentlyContinue) -replace ".*=", ""
                            $engineName = ($iniContent | Select-String -Pattern "^EngineName=" -ErrorAction SilentlyContinue) -replace ".*=", ""
                            $patternVersion = ($iniContent | Select-String -Pattern "^PatternVersion=" -ErrorAction SilentlyContinue) -replace ".*=", ""

                            $engineData = [PSCustomObject]@{
                                "Engine Name"          = if ($engineName) { $engineName } else { "-" }
                                "Engine Version"       = if ($engineVersion) { $engineVersion } else { "-" }
                                "Pattern Version"      = if ($patternVersion) { $patternVersion } else { "-" }
                                "Engine Pfad"          = $fipfsPath
                            }
                            $serverHTML += (ConvertTo-HTMLTable -Data @($engineData))
                            $serverHTML += "<p><em>Letztes Update-Datum: $(Get-Date $((Get-Item $scanEnginePath).LastWriteTime) -Format 'dd.MM.yyyy HH:mm')</em></p>"
                        } catch {
                            $serverHTML += "<p class='no-data'>ScanEngine.ini konnte nicht gelesen werden: $_</p>"
                        }
                    } elseif (-not $isLocal) {
                        $serverHTML += "<p class='no-data'>FIP-FS Scan Engine Details nur lokal oder via UNC-Pfad verfügbar.</p>"
                    } else {
                        $serverHTML += "<p class='no-data'>FIP-FS Scan Engine nicht gefunden unter: $fipfsPath</p>"
                    }

                    # Anti-Malware Status
                    $serverHTML += "<h4>Exchange Anti-Malware Status</h4>"
                    try {
                        if ($isLocal) {
                            $malwareFilteringEnabled = (Get-TransportConfig -ErrorAction SilentlyContinue).AntispamAntimalwareEnabled
                            $malwareFilteringEnabled = if ($null -ne $malwareFilteringEnabled) { $malwareFilteringEnabled } else { $true }
                            $malwareStatus = [PSCustomObject]@{
                                "Anti-Malware aktiv"   = if ($malwareFilteringEnabled) { "✅ Aktiviert" } else { "❌ Deaktiviert" }
                                "FIP-FS Pfad"          = $fipfsPath
                            }
                            $serverHTML += (ConvertTo-HTMLTable -Data @($malwareStatus))
                        } else {
                            $serverHTML += "<p class='no-data'>Anti-Malware Status nur lokal abrufbar.</p>"
                        }
                    } catch {
                        $serverHTML += "<p class='no-data'>Anti-Malware Status nicht verfügbar.</p>"
                    }
                } else {
                    $serverHTML += "<p class='error'>Exchange Installationspfad nicht gefunden für Server $server.</p>"
                }
            } catch {
                Write-Log -Message "FIP-FS für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>FIP-FS Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $fipfsHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei FIP-FS für ${server}: $_" -Level "ERROR"
            $fipfsHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "FIP-FS Scan Engine Version" -Content $fipfsHTML
}

# ---------------------------------------------------------------
# 52. EXCHANGE SETTING OVERRIDES
# ---------------------------------------------------------------
function Get-SettingOverrideInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle konfigurierten Exchange Setting Overrides.
    #>
    Write-Log -Message "=== Sammle Exchange Setting Overrides ===" -Level "INFO"

    $overrideHTML = ""

    try {
        $overrides = Get-SettingOverride -Server $ExchangeServers[0] -ErrorAction SilentlyContinue
        if (-not $overrides) {
            # Fallback: Alle Server durchprobieren
            foreach ($server in $ExchangeServers) {
                $overrides = Get-SettingOverride -Server $server -ErrorAction SilentlyContinue
                if ($overrides) { break }
            }
        }

        if ($overrides) {
            $ovData = foreach ($ov in $overrides) {
                [PSCustomObject]@{
                    "Name"            = $ov.Name
                    "Component"       = $ov.Component
                    "Section"         = $ov.Section
                    "Parameters"      = ($ov.Parameters -join '; ')
                    "Reason"          = $ov.Reason
                    "Server"          = $ov.Server
                    "MinVersion"      = $ov.MinVersion
                    "MaxVersion"      = $ov.MaxVersion
                    "FixVersion"     = $ov.FixVersion
                    "Status"          = $ov.Status
                    "Erstellt"        = $ov.WhenCreated
                }
            }
            $overrideHTML += "<h3>Exchange Setting Overrides</h3>"
            $overrideHTML += "<p><strong>Anzahl:</strong> $($ovData.Count)</p>"
            $overrideHTML += "<p><em>Setting Overrides erlauben das Überschreiben von Exchange-Standardeinstellungen für Diagnose- und Fehlerbehebungszwecke. Sie sollten nur temporär aktiv sein.</em></p>"
            $overrideHTML += (ConvertTo-HTMLTable -Data ($ovData | Sort-Object Name))
        } else {
            $overrideHTML += "<h3>Exchange Setting Overrides</h3>"
            $overrideHTML += "<p class='success'>✅ Keine aktiven Setting Overrides gefunden.</p>"
        }
    } catch {
        Write-Log -Message "SettingOverride-Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $overrideHTML += "<h3>Exchange Setting Overrides</h3>"
        $overrideHTML += "<p class='no-data'>Setting Overrides konnten nicht abgerufen werden: $_</p>"
    }

    New-HTMLSection -Title "Exchange Setting Overrides" -Content $overrideHTML
}

# ---------------------------------------------------------------
# 53. EXCHANGE MAINTENANCE MODE (SERVER COMPONENT STATE)
# ---------------------------------------------------------------
function Get-ServerComponentStateInfo {
    <#
    .SYNOPSIS
        Prüft den Component State aller Exchange-Server (Maintenance Mode Erkennung).
    #>
    Write-Log -Message "=== Prüfe Exchange Server Component States ===" -Level "INFO"

    $compHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Component States für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $components = Get-ServerComponentState -Identity $server -ErrorAction SilentlyContinue
                if ($components) {
                    $inactiveComponents = $components | Where-Object { $_.State -eq "Inactive" }
                    $activeComponents = $components | Where-Object { $_.State -eq "Active" }

                    # Zusammenfassung
                    $serverHTML += "<h4>Component State Zusammenfassung</h4>"
                    $summaryData = [PSCustomObject]@{
                        "Server"                 = $server
                        "Gesamt Komponenten"     = $components.Count
                        "Aktiv"                  = $activeComponents.Count
                        "Inaktiv"                = $inactiveComponents.Count
                        "Maintenance Mode"       = if ($inactiveComponents.Count -gt 0 -and $inactiveComponents.Count -ge ($components.Count * 0.8)) { "🔴 JA (vermutlich)" } else { "✅ Nein" }
                    }
                    $serverHTML += (ConvertTo-HTMLTable -Data @($summaryData))

                    if ($inactiveComponents.Count -gt 0) {
                        $serverHTML += "<h4>Inaktive Komponenten</h4>"
                        $inactiveData = foreach ($ic in $inactiveComponents) {
                            [PSCustomObject]@{
                                "Komponente" = $ic.Component
                                "Status"     = "🔴 Inaktiv"
                                "Remote"     = $ic.Remote
                            }
                        }
                        $serverHTML += (ConvertTo-HTMLTable -Data ($inactiveData | Sort-Object Komponente))
                        $serverHTML += "<p class='error'>⚠️ $($inactiveComponents.Count) Komponente(n) sind inaktiv. Möglicherweise ist der Server im Wartungsmodus!</p>"
                    }

                    if ($activeComponents.Count -gt 0) {
                        $serverHTML += "<h4>Aktive Komponenten (Auszug)</h4>"
                        $activeData = foreach ($ac in $activeComponents) {
                            [PSCustomObject]@{
                                "Komponente" = $ac.Component
                                "Status"     = "✅ Aktiv"
                                "Remote"     = $ac.Remote
                            }
                        }
                        $serverHTML += (ConvertTo-HTMLTable -Data ($activeData | Sort-Object Komponente))
                    }
                } else {
                    $serverHTML += "<p class='no-data'>Keine Component States abrufbar für Server $server.</p>"
                }
            } catch {
                Write-Log -Message "Component State für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Component States konnten nicht abgerufen werden: $_</p>"
            }

            $compHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei Component States für ${server}: $_" -Level "ERROR"
            $compHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Exchange Server Component States" -Content $compHTML
}

# ---------------------------------------------------------------
# 54. SECURITY CVEs (CVE-2021-34470, CVE-2022-21978)
# ---------------------------------------------------------------
function Get-SecurityCVEInfo {
    <#
    .SYNOPSIS
        Prüft auf bekannte Sicherheitslücken (CVEs) in der Exchange-Konfiguration.
    #>
    Write-Log -Message "=== Prüfe Security CVEs ===" -Level "INFO"

    $cveHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "CVE-Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # --- CVE-2021-34470: Exchange-ACL-Permission (AD ACL) ---
                $cve34470 = [PSCustomObject]@{
                    "CVE"                     = "CVE-2021-34470"
                    "Beschreibung"            = "Exchange-ACL-Schwachstelle - Privilegienausweitung"
                    "Betroffene Versionen"    = "Exchange 2013, 2016, 2019"
                    "Fix verfügbar seit"      = "2021-06-01 (April 2021 SU)"
                    "Remediation"             = "Set-OrganizationConfig -ACLableSubscriptionEnabled $false"
                }

                # --- CVE-2022-21978: Exchange-Authentifizierungsumgehung ---
                $cve21978 = [PSCustomObject]@{
                    "CVE"                     = "CVE-2022-21978"
                    "Beschreibung"            = "Exchange Server Authentifizierungsumgehung via OWA"
                    "Betroffene Versionen"    = "Exchange 2013, 2016, 2019"
                    "Fix verfügbar seit"      = "2022-01-11 (January 2022 SU)"
                    "Remediation"             = "Installation des Januar 2022 Sicherheitsupdates"
                }

                # Prüfe Exchange Build-Nummer gegen bekannte Fix-Versionen
                $exchangeBuild = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ValueName "MsiProductMajor"
                $exchangeBuildMinor = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ValueName "MsiProductMinor"

                $supTable = @($cve34470, $cve21978)

                $serverHTML += "<h4>Security CVE Übersicht</h4>"
                $serverHTML += (ConvertTo-HTMLTable -Data $supTable)

                # Exchange-Version prüfen
                $serverHTML += "<h4>Exchange Update Level Prüfung</h4>"
                try {
                    $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                    if ($isLocal) {
                        $updateInfo = Get-ExchangeServerUpdateInfo -ErrorAction SilentlyContinue
                        if ($updateInfo) {
                            $serverHTML += "<p>$updateInfo</p>"
                        } else {
                            $serverHTML += "<p>Exchange Build: $exchangeBuild.$exchangeBuildMinor</p>"
                        }
                    } else {
                        $serverHTML += "<p>Exchange Build: $exchangeBuild.$exchangeBuildMinor</p>"
                    }
                } catch {
                    $serverHTML += "<p>Exchange Build: $exchangeBuild.$exchangeBuildMinor</p>"
                }

                $serverHTML += "<h4>Empfehlungen</h4>"
                $serverHTML += "<p>Führen Sie regelmäßige Security-Updates durch und prüfen Sie die aktuelle Exchange-Build-Nummer gegen die neueste verfügbare Version.</p>"
                $serverHTML += "<p>Aktuelle Exchange Security News: <a href='https://msrc.microsoft.com/update-guide' target='_blank'>Microsoft Security Response Center</a></p>"

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "CVE-Check für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>CVE-Prüfung konnte nicht durchgeführt werden: $_</p>"
            }

            $cveHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei CVE-Check für ${server}: $_" -Level "ERROR"
            $cveHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    # AD-spezifische CVE-Prüfung (CVE-2021-34470 betrifft AD ACL)
    try {
        $cveHTML += "<h3>Active Directory Security Prüfung</h3>"
        $adModule = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
        if ($adModule) {
            try {
                Import-Module ActiveDirectory -ErrorAction SilentlyContinue
                $adRoot = Get-ADRootDSE -ErrorAction SilentlyContinue
                if ($adRoot) {
                    $adSchema = $adRoot.schemaNamingContext
                    $cveHTML += "<p><strong>AD Schema:</strong> $adSchema</p>"
                    $cveHTML += "<p><em>Hinweis: CVE-2021-34470 betrifft die Exchange-ACL-Konfiguration. Prüfen Sie mit:</em></p>"
                    $cveHTML += "<code>Get-OrganizationConfig | fl ACLableSubscriptionEnabled</code>"
                }
            } catch {
                Write-Log -Message "AD-CVE Prüfung fehlgeschlagen: $_" -Level "WARNING"
            }
        } else {
            $cveHTML += "<p class='no-data'>AD-Modul nicht verfügbar für erweiterte CVE-Prüfung.</p>"
        }
    } catch {
        Write-Log -Message "AD-CVE Prüfung fehlgeschlagen: $_" -Level "WARNING"
    }

    New-HTMLSection -Title "Security CVE Prüfung" -Content $cveHTML
}

# ---------------------------------------------------------------
# 55. HTTP PROXY KONFIGURATION (WINHTTP)
# ---------------------------------------------------------------
function Get-HttpProxyInfo {
    <#
    .SYNOPSIS
        Dokumentiert die HTTP-Proxy Konfiguration (WinHTTP) pro Server.
    #>
    Write-Log -Message "=== Sammle HTTP Proxy Konfiguration ===" -Level "INFO"

    $proxyHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "HTTP Proxy Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # WinHTTP Proxy aus Registry (CurrentUser + LocalMachine)
                $proxyKeys = @(
                    @{ Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"; Name = "DefaultConnectionSettings" },
                    @{ Path = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"; Name = "DefaultConnectionSettings" },
                    @{ Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"; Name = "ProxyEnable" },
                    @{ Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"; Name = "ProxyServer" },
                    @{ Path = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings"; Name = "ProxyEnable" },
                    @{ Path = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings"; Name = "ProxyServer" }
                )

                $proxyDataList = @()
                foreach ($key in $proxyKeys) {
                    try {
                        $value = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $key.Path -ValueName $key.Name -ErrorAction SilentlyContinue
                        if ($null -ne $value) {
                            $displayValue = if ($value -is [byte[]]) { "Binärdaten ($($value.Length) Bytes)" } else { $value.ToString() }
                            $proxyDataList += [PSCustomObject]@{
                                "Registry Pfad"  = "HKLM\$($key.Path)\$($key.Name)"
                                "Wert"           = $displayValue
                            }
                        }
                    } catch { }
                }

                if ($proxyDataList.Count -gt 0) {
                    $serverHTML += "<h4>WinHTTP / Internet Settings Registry</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data ($proxyDataList | Sort-Object "Registry Pfad"))
                }

                # ProxyEnable boolesch interpretieren
                $proxyEnabled = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" -ValueName "ProxyEnable" -ErrorAction SilentlyContinue
                $proxyServer = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" -ValueName "ProxyServer" -ErrorAction SilentlyContinue

                $serverHTML += "<h4>HTTP Proxy Status</h4>"
                if ($proxyEnabled -eq 1) {
                    $proxyStatus = "🔴 Proxy aktiviert"
                    $proxyServerValue = if ($proxyServer) { $proxyServer } else { "Kein Proxy-Server konfiguriert" }
                    $serverHTML += "<p class='error'>$proxyStatus</p>"
                    $serverHTML += "<p><strong>Proxy-Server:</strong> $proxyServerValue</p>"
                } elseif ($null -ne $proxyEnabled) {
                    $serverHTML += "<p>✅ Proxy deaktiviert (ProxyEnable = $proxyEnabled)</p>"
                } else {
                    $serverHTML += "<p class='no-data'>Keine Proxy-Konfiguration in Registry gefunden.</p>"
                }

                # WinHTTP Proxy via netshell (nur lokal)
                try {
                    $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                    if ($isLocal) {
                        $netshProxy = netsh winhttp show proxy 2>&1 | Out-String
                        if ($netshProxy) {
                            $serverHTML += "<h4>WinHTTP Proxy (netsh)</h4>"
                            $serverHTML += "<pre style='background:#F5F5F5;padding:10px;font-size:9pt;'>$([System.Web.HttpUtility]::HtmlEncode($netshProxy.Trim()))</pre>"
                        }
                    }
                } catch {
                    Write-Log -Message "Netsh WinHTTP Proxy Check fehlgeschlagen: $_" -Level "WARNING"
                }

                # Proxy-Konfiguration für Exchange-spezifische Dienste
                $serverHTML += "<h4>Exchange & Proxy</h4>"
                $serverHTML += "<p><em>Ein HTTP-Proxy kann Exchange-Funktionen beeinträchtigen:</em></p>"
                $serverHTML += "<ul>"
                $serverHTML += "<li>Hybridkonfiguration mit Exchange Online</li>"
                $serverHTML += "<li>EEMS (Emergency Mitigation Service) - Pattern-Updates</li>"
                $serverHTML += "<li>Anti-Spam Definition Updates</li>"
                $serverHTML += "<li>Autodiscover externe Abfragen</li>"
                $serverHTML += "</ul>"

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "HTTP Proxy für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>HTTP Proxy Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $proxyHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei HTTP Proxy für ${server}: $_" -Level "ERROR"
            $proxyHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "HTTP Proxy Konfiguration" -Content $proxyHTML
}

# ---------------------------------------------------------------
# 56. INSTALLIERTE ANTIVIRENLÖSUNG
# ---------------------------------------------------------------
function Get-AntivirusProductInfo {
    <#
    .SYNOPSIS
        Erkennt installierte Antivirenlösungen und deren Status auf den Servern.
    #>
    Write-Log -Message "=== Sammle installierte Antivirenlösungen ===" -Level "INFO"

    $avHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "Antivirus-Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # Methode 1: SecurityCenter2 WMI (Windows Defender / Security Center)
                $avProducts = @()
                try {
                    $securityCenter = Get-CimInstance -CimSession $session -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
                    if ($securityCenter) {
                        $avProducts = foreach ($av in $securityCenter) {
                            $stateText = switch ($av.productState) {
                                {$_ -band 0x1000} { "Aktiv" }
                                {$_ -band 0x2000} { "Aktiv (Überwachen)" }
                                default { "Unbekannt" }
                            }
                            $enabledText = if ($av.productState -band 0x100) { "Ja" } else { "Nein" }
                            $upToDateText = if ($av.productState -band 0x10) { "Ja" } else { "Nein" }
                            $realTimeText = if (($av.productState -band 0x1000) -eq 0x1000) { "✅ Aktiv" } else { "❌ Inaktiv" }

                            [PSCustomObject]@{
                                "Antivirenprodukt"   = $av.displayName
                                "Hersteller"         = if ($av.companyName) { $av.companyName } else { "-" }
                                "Echtzeitschutz"     = $realTimeText
                                "Aktiviert"          = $enabledText
                                "Definition aktuell" = $upToDateText
                                "Product State (Hex)"= "0x$('{0:X}' -f $av.productState)"
                            }
                        }
                    }
                } catch {
                    Write-Log -Message "SecurityCenter2 WMI nicht verfügbar auf $server (Namespace existiert nicht): $_" -Level "WARNING"
                }

                # Methode 2: Registry - Installierte AV-Produkte auslesen
                if ($avProducts.Count -eq 0) {
                    try {
                        $avKeys = Get-CimInstance -CimSession $session -ClassName StdRegProv -Namespace root/default -ErrorAction SilentlyContinue
                        if ($avKeys) {
                            $avPaths = @(
                                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                                "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                            )
                            $avFromReg = @()
                            foreach ($regPath in $avPaths) {
                                $enumResult = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName EnumKey -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $regPath } -ErrorAction SilentlyContinue
                                if ($enumResult -and $enumResult.sNames) {
                                    foreach ($subKey in $enumResult.sNames) {
                                        $fullPath = "$regPath\$subKey"
                                        $displayName = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullPath; sValueName = "DisplayName" } -ErrorAction SilentlyContinue
                                        if ($displayName.sValue -and ($displayName.sValue -match "Antivirus|Anti-Virus|Anti Virus|Virenschutz|Defender|McAfee|Trend Micro|Sophos|Symantec|Norton|Kaspersky|Bitdefender|Avast|AVG|Malwarebytes|CrowdStrike|SentinelOne|Palo Alto|Panda|F-Secure|ESET|G Data|VIPRE|Webroot|Fortinet|Check Point" -or $subKey -match "antivirus|defender|mcafee|trend|sophos|Symantec|kaspersky|bitdefender|avast|avg|malwarebytes|crowdstrike|sentinel|paloalto|panda|fsecure|eset" )) {
                                            $displayVersion = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullPath; sValueName = "DisplayVersion" } -ErrorAction SilentlyContinue
                                            $publisher = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullPath; sValueName = "Publisher" } -ErrorAction SilentlyContinue
                                            $avFromReg += [PSCustomObject]@{
                                                "Antivirenprodukt" = $displayName.sValue
                                                "Version"          = if ($displayVersion.sValue) { $displayVersion.sValue } else { "-" }
                                                "Hersteller"       = if ($publisher.sValue) { $publisher.sValue } else { "-" }
                                            }
                                        }
                                    }
                                }
                            }
                            if ($avFromReg.Count -gt 0) {
                                $serverHTML += "<h4>Installierte Antiviren-/Security-Software (Registry)</h4>"
                                $serverHTML += "<p><strong>Anzahl:</strong> $($avFromReg.Count)</p>"
                                $serverHTML += (ConvertTo-HTMLTable -Data ($avFromReg | Sort-Object Hersteller))
                            }
                        }
                    } catch {
                        Write-Log -Message "AV Registry-Abfrage (CIM) fehlgeschlagen: $_" -Level "WARNING"
                    }
                }

                # Ergebnisse aus SecurityCenter anzeigen
                if ($avProducts.Count -gt 0) {
                    $serverHTML += "<h4>Windows Security Center - Erkannte Antivirenprodukte</h4>"
                    $serverHTML += (ConvertTo-HTMLTable -Data ($avProducts | Sort-Object Hersteller))
                }

                # Methode 3: Windows Defender spezifisch (lokal)
                try {
                    $isLocal = $server -match "^($([regex]::Escape($env:COMPUTERNAME))|localhost|\.)$"
                    if ($isLocal) {
                        try {
                            $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
                            if ($defenderStatus) {
                                $defData = [PSCustomObject]@{
                                    "RealTime Protection Enabled" = if ($defenderStatus.RealTimeProtectionEnabled) { "✅ Aktiv" } else { "❌ Inaktiv" }
                                    "AM Service Enabled"           = $defenderStatus.AMServiceEnabled
                                    "Antispyware Enabled"          = $defenderStatus.AntispywareEnabled
                                    "Antivirus Enabled"            = $defenderStatus.AntivirusEnabled
                                    "Behavior Monitor Enabled"     = $defenderStatus.BehaviorMonitorEnabled
                                    "Defender Version"             = $defenderStatus.AMProductVersion
                                    "Engine Version"               = $defenderStatus.AMEngineVersion
                                    "Definitions Update"           = if ($defenderStatus.AntivirusSignatureLastUpdated) { $defenderStatus.AntivirusSignatureLastUpdated } else { "-" }
                                    "Definitions Version"          = if ($defenderStatus.AntivirusSignatureVersion) { $defenderStatus.AntivirusSignatureVersion } else { "-" }
                                }
                                $serverHTML += "<h4>Microsoft Defender Status (lokal)</h4>"
                                $serverHTML += (ConvertTo-HTMLTable -Data @($defData))
                            }
                        } catch {
                            Write-Log -Message "Get-MpComputerStatus nicht verfügbar (kein Defender oder Win10+): $_" -Level "WARNING"
                        }
                    }
                } catch {
                    Write-Log -Message "Lokaler Defender-Check fehlgeschlagen: $_" -Level "WARNING"
                }

                # Kein AV gefunden
                if ($avProducts.Count -eq 0) {
                    try {
                        $regCheckAvCount = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName EnumKey -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" } -ErrorAction SilentlyContinue
                        $serverHTML += "<p class='no-data'>Keine Antivirenprodukte via SecurityCenter gefunden (möglicherweise kein SecurityCenter oder AV nicht registriert).</p>"
                        $serverHTML += "<p><em>Hinweis: Der installierte Virenscanner kann auch über die Registry unter den Uninstall-Keys gefunden werden (siehe Tabelle oben).</em></p>"
                    } catch { }
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
            } catch {
                Write-Log -Message "Antivirus-Check für $server fehlgeschlagen: $_" -Level "WARNING"
                $serverHTML += "<p class='error'>Antivirus-Informationen konnten nicht abgerufen werden: $_</p>"
            }

            $avHTML += $serverHTML
        } catch {
            Write-Log -Message "Fehler bei Antivirus für ${server}: $_" -Level "ERROR"
            $avHTML += "<h3 class='server-break'>Server: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }
    New-HTMLSection -Title "Installierte Antivirenlösung" -Content $avHTML
}

function Get-NICReceiveBufferInfo {
    <#
    .SYNOPSIS
        Prüft die Receive Buffer Größen von Netzwerkadaptern (10/25/40 Gbit/s) und gibt Empfehlungen.
    .DESCRIPTION
        Analysiert die Receive Buffer Werte aller Netzwerkadapter über CIM.
        Bei Hochgeschwindigkeits-Adaptern (10/25/40 Gbit/s) wird geprüft, ob
        die Receive Buffer ausreichend dimensioniert sind (Intel-Empfehlung).
        Zu kleine Werte können zu StalledDueToTarget_MdbReplication führen.
    #>
    Write-Log -Message "=== Sammle NIC Receive Buffer Informationen ===" -Level "INFO"

    $nicHTML = ""

    foreach ($server in $ExchangeServers) {
        try {
            Write-Log -Message "NIC Receive Buffer Check für Server: $server" -Level "INFO"
            $serverHTML = "<h3 class='server-break'>Server: $server</h3>"

            try {
                $session = New-ServerCimSession -ComputerName $server
                if (-not $session) { throw "Keine CIM-Session verfügbar." }

                # Netzwerkadapter abfragen
                $adapters = Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapter -Filter "NetEnabled = TRUE AND Name IS NOT NULL" -ErrorAction SilentlyContinue
                if (-not $adapters) {
                    $adapters = Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapter -Filter "Name IS NOT NULL" -ErrorAction SilentlyContinue
                }

                if ($adapters) {
                    $nicData = @()
                    foreach ($adapter in $adapters) {
                        $name = $adapter.Name
                        $speed = $adapter.Speed
                        $speedGbps = if ($speed -and $speed -gt 0) { [math]::Round($speed / 1e9, 1) } else { 0 }
                        $mac = $adapter.MACAddress
                        $manufacturer = $adapter.Manufacturer

                        if ($speedGbps -ge 1 -or $speed -eq 0) {
                            $nicKey = $null
                            $rssQueues = $null
                            $recvBuffers = $null
                            $sendBuffers = $null

                            # Registry-Zugriff auf NIC-Parameter
                            try {
                                $regPath = "SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
                                $regResult = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName EnumKey -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $regPath } -ErrorAction SilentlyContinue
                                if ($regResult -and $regResult.sNames) {
                                    $adapterPnpId = $adapter.PNPDeviceID
                                    if ($adapterPnpId) {
                                        $pnpIdClean = $adapterPnpId -replace '\\', '#'
                                        foreach ($subKey in $regResult.sNames) {
                                            $fullKeyPath = "$regPath\$subKey"
                                            $keyPnpId = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetStringValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullKeyPath; sValueName = "*PNPDeviceID" } -ErrorAction SilentlyContinue
                                            if ($keyPnpId.sValue -and $keyPnpId.sValue -eq $adapterPnpId) {
                                                $recvBuf = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetDWordValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullKeyPath; sValueName = "ReceiveBuffers" } -ErrorAction SilentlyContinue
                                                $sendBuf = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetDWordValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullKeyPath; sValueName = "TransmitBuffers" } -ErrorAction SilentlyContinue
                                                $rssQueues = Invoke-CimMethod -CimSession $session -ClassName StdRegProv -MethodName GetDWordValue -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $fullKeyPath; sValueName = "*RssQueues" } -ErrorAction SilentlyContinue

                                                if ($recvBuf) { $recvBuffers = $recvBuf.uValue }
                                                if ($sendBuf) { $sendBuffers = $sendBuf.uValue }
                                                break
                                            }
                                        }
                                    }
                                }
                            } catch {
                                Write-Log -Message "Registry-Zugriff auf NIC-Parameter nicht möglich: $_" -Level "DEBUG"
                            }

                            # Empfehlung basierend auf Speed
                            $recommendation = ""
                            $statusIcon = ""
                            if ($speedGbps -ge 40) {
                                if ($recvBuffers -ge 4096) { $statusIcon = "✅"; $recommendation = "OK (≥4096 empfohlen für 40 Gbit/s)" }
                                elseif ($recvBuffers -gt 0) { $statusIcon = "⚠️"; $recommendation = "Mindestens 4096 empfohlen für 40 Gbit/s (aktuell: $recvBuffers)" }
                                else { $statusIcon = "⚠️"; $recommendation = "4096 empfohlen für 40 Gbit/s" }
                            } elseif ($speedGbps -ge 25) {
                                if ($recvBuffers -ge 4096) { $statusIcon = "✅"; $recommendation = "OK (≥4096 empfohlen für 25 Gbit/s)" }
                                elseif ($recvBuffers -gt 0) { $statusIcon = "⚠️"; $recommendation = "Mindestens 4096 empfohlen für 25 Gbit/s (aktuell: $recvBuffers)" }
                                else { $statusIcon = "⚠️"; $recommendation = "4096 empfohlen für 25 Gbit/s" }
                            } elseif ($speedGbps -ge 10) {
                                if ($recvBuffers -ge 2048) { $statusIcon = "✅"; $recommendation = "OK (≥2048 empfohlen für 10 Gbit/s)" }
                                elseif ($recvBuffers -gt 0) { $statusIcon = "⚠️"; $recommendation = "Mindestens 2048 empfohlen für 10 Gbit/s (aktuell: $recvBuffers)" }
                                else { $statusIcon = "⚠️"; $recommendation = "2048 empfohlen für 10 Gbit/s" }
                            } elseif ($speedGbps -ge 1) {
                                $statusIcon = "ℹ️"; $recommendation = "Keine spezifische Empfehlung (1 Gbit/s)"
                            }

                            if ($speedGbps -gt 0 -or ($recvBuffers -gt 0)) {
                                $nicData += [PSCustomObject]@{
                                    "Adapter-Name"          = $name
                                    "Geschwindigkeit"       = if ($speedGbps -gt 0) { "$speedGbps Gbit/s" } else { "Unbekannt" }
                                    "MAC-Adresse"           = if ($mac) { $mac } else { "-" }
                                    "Hersteller"            = if ($manufacturer) { $manufacturer } else { "-" }
                                    "Receive Buffers"       = if ($recvBuffers) { $recvBuffers } else { "Standard (512)" }
                                    "Transmit Buffers"      = if ($sendBuffers) { $sendBuffers } else { "Standard" }
                                    "RSS Queues"            = if ($rssQueues) { $rssQueues.uValue } else { "Standard" }
                                    "Empfehlung"            = "$statusIcon $recommendation"
                                }
                            }
                        }
                    }

                    if ($nicData.Count -gt 0) {
                        $serverHTML += ConvertTo-HTMLTable -Data $nicData -HeaderColor "#005a9e"
                        $serverHTML += "<p style='font-size:0.9em; color:#666;'><strong>Hinweis:</strong> Microsoft CSS-Exchange empfiehlt bei 25 Gbit/s Adaptern Receive Buffers von 4096 (Default: 512) um StalledDueToTarget_MdbReplication zu vermeiden. Siehe <a href='https://github.com/microsoft/CSS-Exchange/issues/2560'>GitHub Issue #2560</a>.</p>"
                    } else {
                        $serverHTML += "<p>Keine Hochgeschwindigkeits-Netzwerkadapter gefunden.</p>"
                    }
                } else {
                    $serverHTML += "<p class='error'>Keine Netzwerkadapter abrufbar.</p>"
                }

                if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }

            } catch {
                Write-Log -Message "Fehler bei NIC-Analyse für ${server}: $_" -Level "ERROR"
                $serverHTML += "<p class='error'>Fehler: $_</p>"
            }

            $nicHTML += $serverHTML

        } catch {
            Write-Log -Message "Fehler bei NIC-Analyse für ${server}: $_" -Level "ERROR"
        }
    }
    New-HTMLSection -Title "NIC Receive Buffer Analyse" -Content $nicHTML
}

#endregion

#region ============================================================
# HTML-DOKUMENT GENERIERUNG
#endregion ============================================================

function Build-HTMLDocument {
    <#
    .SYNOPSIS
        Baut das finale HTML-Dokument mit Deckblatt, Inhaltsverzeichnis und allen Sektionen.
    #>
    Write-Log -Message "=== Erstelle HTML-Dokument ===" -Level "INFO"

    $cssStyle = @"
    <style>
        body {
            font-family: 'Segoe UI', Calibri, Arial, sans-serif;
            font-size: 11pt;
            color: #333333;
            margin: 40px;
            line-height: 1.5;
        }
        .cover-page {
            text-align: center;
            padding: 100px 0;
            page-break-after: always;
            border-bottom: 3px solid #0078D4;
        }
        .cover-page h1 {
            font-size: 28pt;
            color: #0078D4;
            margin-bottom: 10px;
        }
        .cover-page h2 {
            font-size: 18pt;
            color: #555555;
            font-weight: normal;
        }
        .cover-page .meta {
            margin-top: 50px;
            font-size: 12pt;
            color: #777777;
        }
        h2 {
            color: #0078D4;
            font-size: 16pt;
            border-bottom: 2px solid #0078D4;
            padding-bottom: 5px;
            margin-top: 30px;
        }
        h3 {
            color: #333333;
            font-size: 13pt;
            border-bottom: 1px solid #CCCCCC;
            padding-bottom: 3px;
            margin-top: 20px;
        }
        h3.server-break {
            page-break-before: avoid;
        }
        h4 {
            color: #555555;
            font-size: 11pt;
            margin-top: 15px;
        }
        .toc {
            page-break-after: always;
        }
        .toc h2 {
            page-break-before: avoid;
        }
        .toc ul {
            list-style-type: none;
            padding-left: 0;
        }
        .toc li {
            padding: 4px 0;
        }
        .toc a {
            text-decoration: none;
            color: #0078D4;
        }
        .toc a:hover {
            text-decoration: underline;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0 20px 0;
            font-size: 9pt;
        }
        th {
            background-color: #0078D4;
            color: white;
            text-align: left;
            padding: 8px 10px;
            font-weight: bold;
            border: 1px solid #005A9E;
        }
        td {
            padding: 6px 10px;
            border: 1px solid #DDDDDD;
            vertical-align: top;
            word-wrap: break-word;
            max-width: 300px;
        }
        tr.even { background-color: #F5F5F5; }
        tr.odd { background-color: #FFFFFF; }
        tr:hover { background-color: #E8F4FD; }
        .section { margin-bottom: 20px; }
        .no-data {
            color: #888888;
            font-style: italic;
            padding: 10px;
            background-color: #F9F9F9;
            border-left: 3px solid #CCCCCC;
        }
        .error {
            color: #D32F2F;
            padding: 10px;
            background-color: #FFEBEE;
            border-left: 3px solid #D32F2F;
        }
        .summary-box {
            background-color: #F0F7FF;
            border: 1px solid #0078D4;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }
        .summary-box h3 {
            border: none;
            color: #0078D4;
            margin-top: 0;
        }
        .footer {
            text-align: center;
            font-size: 8pt;
            color: #999999;
            border-top: 1px solid #CCCCCC;
            padding-top: 10px;
            margin-top: 40px;
        }
        @media print {
            body { margin: 20px; line-height: 1.4; font-size: 9pt; }
            h1, h2, h3, h4 { orphans: 3; widows: 3; }
            h2 { page-break-before: auto; margin-top: 0; }
            h3.server-break { page-break-inside: avoid; page-break-before: auto; }
            .section { page-break-inside: avoid; margin-bottom: 15px; }
            table { page-break-inside: avoid; margin: 8px 0 15px 0; }
            tr { page-break-inside: avoid; }
            thead { display: table-header-group; }
            tfoot { display: table-footer-group; }
            .toc { page-break-after: always; }
            .cover-page { page-break-after: always; }
        }
    </style>
"@

    # Inhaltsverzeichnis
    $tocHTML = "<div class='toc'>`n<h2>Inhaltsverzeichnis</h2>`n<ul>"
    foreach ($entry in $script:TOCEntries) {
        $tocHTML += "<li><a href='#$($entry.Anchor)'>$($entry.Title)</a></li>`n"
    }
    $tocHTML += "</ul></div>"

    # Zusammenfassung
    $summaryHTML = @"
    <div class="summary-box">
        <h3>Dokumentations-Zusammenfassung</h3>
        <p><strong>Dokumentierte Server:</strong> $($ExchangeServers -join ', ')</p>
        <p><strong>Erstellt am:</strong> $script:Timestamp</p>
        <p><strong>Erstellt von:</strong> $script:DocAuthor auf $script:DocComputerName</p>
        <p><strong>Sektionen:</strong> $($script:SectionCounter)</p>
        <p><strong>Fehler:</strong> $($script:ErrorCount)</p>
        <p><strong>Warnungen:</strong> $($script:WarningCount)</p>
    </div>
"@

    # Gesamtes HTML-Dokument
    $fullHTML = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="author" content="$script:DocAuthor">
    <meta name="generator" content="Exchange Dokumentations-Skript v1.7">
    <title>$script:DocTitle - $script:DocSubTitle</title>
    $cssStyle
</head>
<body>

    <div class="cover-page">
        <h1>$script:DocTitle</h1>
        <h2>$script:DocSubTitle</h2>
        <div class="meta">
            <p>Erstellt am: $script:DateOnly</p>
            <p>Erstellt von: $script:DocAuthor</p>
            <p>Dokumentierte Server: $($ExchangeServers -join ', ')</p>
            <p>Version: 1.7 (Exchange 2019 &amp; SE / CIM-DCOM Fallback)</p>
        </div>
    </div>

    $tocHTML

    $summaryHTML

    $($script:HTMLSections -join "`n`n")

    <div class="footer">
        <p>$script:DocTitle | $script:DocSubTitle | Erstellt: $script:Timestamp | PowerShell Dokumentation v1.7</p>
    </div>

</body>
</html>
"@

    return $fullHTML
}

#endregion

#region ============================================================
# SEKTIONS-REGISTRY, MARKDOWN, EXPORT & GUI
#endregion ============================================================

function Get-DocSectionRegistry {
    <#
    .SYNOPSIS
        Liefert die geordnete Liste aller verfügbaren Dokumentations-Sektionen
        (Schlüssel, Anzeigename, Kategorie, zugehörige Funktion). Basis für GUI und Runner.
    #>
    return @(
        # === HARDWARE & OS ===
        [PSCustomObject]@{ Key = "Hardware";           Label = "Hardware & Server-Details";        Category = "Hardware & OS";      Function = "Get-HardwareInformation" }
        [PSCustomObject]@{ Key = "ProcessorCores";     Label = "Processor Core Analyse";            Category = "Hardware & OS";      Function = "Get-ProcessorCoreAnalysis" }
        [PSCustomObject]@{ Key = "RAM";                Label = "RAM Requirements Check";            Category = "Hardware & OS";      Function = "Get-RAMRequirementsInfo" }
        [PSCustomObject]@{ Key = "Patch";              Label = "Exchange Version & Build";          Category = "Hardware & OS";      Function = "Get-PatchInformation" }
        [PSCustomObject]@{ Key = "Software";           Label = "Installierte Software";             Category = "Hardware & OS";      Function = "Get-InstalledSoftwareInfo" }
        [PSCustomObject]@{ Key = "Network";            Label = "Netzwerk-Konfiguration (NIC Speed)"; Category = "Hardware & OS";    Function = "Get-NetworkConfigurationInfo" }
        [PSCustomObject]@{ Key = "PowerPlan";          Label = "Power Plan & Performance";          Category = "Hardware & OS";      Function = "Get-PowerPlanInfo" }
        [PSCustomObject]@{ Key = "SMBv1";              Label = "SMBv1 Status (Sicherheit)";         Category = "Hardware & OS";      Function = "Get-SMBv1StatusInfo" }
        [PSCustomObject]@{ Key = "EventLog";           Label = "Event Logs (7 Tage)";               Category = "Hardware & OS";      Function = "Get-EventLogInfo" }
        [PSCustomObject]@{ Key = "Firewall";           Label = "Firewall-Regeln & Ports";           Category = "Hardware & OS";      Function = "Get-FirewallInfo" }
        [PSCustomObject]@{ Key = "DiskSpace";          Label = "Disk Space (DB/Logs/Temp)";         Category = "Hardware & OS";      Function = "Get-DiskSpaceInfo" }
        [PSCustomObject]@{ Key = "RPCPort";            Label = "RPC Port 135 Status";               Category = "Hardware & OS";      Function = "Get-RPCPortStatusInfo" }
        [PSCustomObject]@{ Key = "LDAP";               Label = "LDAP Konnektivität (DC)";           Category = "Hardware & OS";      Function = "Get-LDAPConnectivityInfo" }
        [PSCustomObject]@{ Key = "Licensing";          Label = "Lizenzierung";                      Category = "Hardware & OS";      Function = "Get-LicensingInfo" }

        # === ACTIVE DIRECTORY ===
        [PSCustomObject]@{ Key = "FSMO";               Label = "FSMO-Rollen";                       Category = "Active Directory";   Function = "Get-FSMORoles" }
        [PSCustomObject]@{ Key = "AD";                 Label = "Active Directory & Schema";         Category = "Active Directory";   Function = "Get-ADInformation" }

        # === EXCHANGE GRUNDKONFIGURATION ===
        [PSCustomObject]@{ Key = "ExchangeServer";     Label = "Exchange Server Übersicht";         Category = "Exchange Basis";     Function = "Get-ExchangeServerOverview" }
        [PSCustomObject]@{ Key = "ExchangeServices";   Label = "Exchange Service Status";           Category = "Exchange Basis";     Function = "Get-ExchangeServiceStatusInfo" }
        [PSCustomObject]@{ Key = "OrgConfig";          Label = "Organisations-Konfiguration";       Category = "Exchange Basis";     Function = "Get-OrganizationConfigInfo" }
        [PSCustomObject]@{ Key = "URLs";               Label = "URLs / VirtualDirectories / EEMS";  Category = "Exchange Basis";     Function = "Get-ExchangeURLs" }
        [PSCustomObject]@{ Key = "IISAppPools";        Label = "IIS Application Pool Konfiguration"; Category = "Exchange Basis";   Function = "Get-IISAppPoolSettingsInfo" }
        [PSCustomObject]@{ Key = "Database";           Label = "Datenbanken & DAG";                 Category = "Exchange Basis";     Function = "Get-DatabaseAndDAGInfo" }
        [PSCustomObject]@{ Key = "DAGHealth";          Label = "DAG Replication Health";            Category = "Exchange Basis";     Function = "Get-DAGReplicationHealthInfo" }
        [PSCustomObject]@{ Key = "DBCopies";           Label = "Database Copies & Redundanz";       Category = "Exchange Basis";     Function = "Get-DatabaseCopiesInfo" }
        [PSCustomObject]@{ Key = "Certificate";        Label = "Zertifikate & Ablauf-Status";       Category = "Exchange Basis";     Function = "Get-CertificateInfo" }
        [PSCustomObject]@{ Key = "CertExpiration";     Label = "Certificate Expiration Check";      Category = "Exchange Basis";     Function = "Get-CertificateExpirationInfo" }
        [PSCustomObject]@{ Key = "MAPI";               Label = "MAPI Connectivity (Outlook)";       Category = "Exchange Basis";     Function = "Get-MAPIConnectivityInfo" }
        [PSCustomObject]@{ Key = "RBAC";               Label = "RBAC-Rollen";                       Category = "Exchange Basis";     Function = "Get-RBACInfo" }

        # === MAIL-FLOW & TRANSPORT ===
        [PSCustomObject]@{ Key = "SendConnector";      Label = "Sendeconnectoren";                  Category = "Mail-Flow";          Function = "Get-SendConnectorInfo" }
        [PSCustomObject]@{ Key = "ReceiveConnector";   Label = "Empfangsconnectoren";               Category = "Mail-Flow";          Function = "Get-ReceiveConnectorInfo" }
        [PSCustomObject]@{ Key = "RemoteDomain";       Label = "Remote Domains";                    Category = "Mail-Flow";          Function = "Get-RemoteDomainInfo" }
        [PSCustomObject]@{ Key = "AcceptedDomain";     Label = "Accepted Domains";                  Category = "Mail-Flow";          Function = "Get-AcceptedDomainInfo" }
        [PSCustomObject]@{ Key = "TransportRule";      Label = "Transport-Regeln";                  Category = "Mail-Flow";          Function = "Get-TransportRuleInfo" }
        [PSCustomObject]@{ Key = "TransportStorage";   Label = "Transport - Speicherorte";          Category = "Mail-Flow";          Function = "Get-TransportComponentStorageInfo" }
        [PSCustomObject]@{ Key = "SMTPRelay";          Label = "SMTP-Relay Konfiguration";          Category = "Mail-Flow";          Function = "Get-SMTPRelayInfo" }

        # === DNS & EXTERNE RECORDS ===
        [PSCustomObject]@{ Key = "MXRecord";           Label = "MX / SPF / DMARC / Autodiscover";   Category = "DNS & Records";      Function = "Get-MXRecordInfo" }
        [PSCustomObject]@{ Key = "DKIM";               Label = "DKIM-Konfiguration";                Category = "DNS & Records";      Function = "Get-DKIMInfo" }

        # === EMPFÄNGER & POSTFÄCHER ===
        [PSCustomObject]@{ Key = "MailboxStats";       Label = "Mailbox-Statistiken";               Category = "Empfänger";          Function = "Get-MailboxStatisticsOverview" }
        [PSCustomObject]@{ Key = "Quota";              Label = "Mailbox-Quotas";                    Category = "Empfänger";          Function = "Get-MailboxQuotaInfo" }
        [PSCustomObject]@{ Key = "EmailAddressPolicy"; Label = "E-Mail-Adressrichtlinien";          Category = "Empfänger";          Function = "Get-EmailAddressPolicyInfo" }
        [PSCustomObject]@{ Key = "AddressList";        Label = "Adresslisten / GAL / ABP";          Category = "Empfänger";          Function = "Get-AddressListInfo" }
        [PSCustomObject]@{ Key = "OAB";                Label = "Offline-Adressbuch";                Category = "Empfänger";          Function = "Get-OABInfo" }
        [PSCustomObject]@{ Key = "PublicFolder";       Label = "Public Folder";                     Category = "Empfänger";          Function = "Get-PublicFolderInfo" }

        # === SICHERHEIT ===
        [PSCustomObject]@{ Key = "Security";           Label = "Sicherheit & Authentifizierung";    Category = "Sicherheit";         Function = "Get-SecurityAndAuthInfo" }
        [PSCustomObject]@{ Key = "AntiSpam";           Label = "Anti-Spam / Anti-Malware";          Category = "Sicherheit";         Function = "Get-AntiSpamMalwareInfo" }
        [PSCustomObject]@{ Key = "AntiVirus";          Label = "Antivirus Ausschlüsse";             Category = "Sicherheit";         Function = "Get-AntivirusExclusionsInfo" }
        [PSCustomObject]@{ Key = "MobileDevice";       Label = "Mobile Device Policies";            Category = "Sicherheit";         Function = "Get-MobileDeviceInfo" }
        [PSCustomObject]@{ Key = "Throttling";         Label = "Throttling Policies";               Category = "Sicherheit";         Function = "Get-ThrottlingPolicyInfo" }

        # === COMPLIANCE ===
        [PSCustomObject]@{ Key = "Compliance";         Label = "Compliance & DLP";                  Category = "Compliance";         Function = "Get-ComplianceInfo" }
        [PSCustomObject]@{ Key = "Retention";          Label = "Retention & Journal";               Category = "Compliance";         Function = "Get-RetentionPolicyInfo" }

        # === TLS/SSL & SICHERHEIT ===
        [PSCustomObject]@{ Key = "TLS";                Label = "TLS/SSL Konfiguration";             Category = "Sicherheit & TLS";   Function = "Get-TLSSSLConfigurationInfo" }

        # === HYBRID ===
        [PSCustomObject]@{ Key = "Hybrid";             Label = "Hybrid mit Exchange Online";        Category = "Hybrid / Cloud";     Function = "Get-HybridConfigurationInfo" }

        # === NEU IN v1.7 ===
        [PSCustomObject]@{ Key = "WindowsFeatures";    Label = "Windows Features & Rollen";         Category = "Hardware & OS";      Function = "Get-WindowsFeaturesInfo" }
        [PSCustomObject]@{ Key = "DotNetFramework";    Label = ".NET Framework Version & DLLs";     Category = "Hardware & OS";      Function = "Get-DotNetFrameworkInfo" }
        [PSCustomObject]@{ Key = "PendingReboot";      Label = "Ausstehende Neustarts";             Category = "Hardware & OS";      Function = "Get-PendingRebootInfo" }
        [PSCustomObject]@{ Key = "CpuThrottling";      Label = "CPU Throttling Analyse";            Category = "Hardware & OS";      Function = "Get-CpuThrottlingInfo" }
        [PSCustomObject]@{ Key = "VCRedist";           Label = "Visual C++ Redistributable";        Category = "Hardware & OS";      Function = "Get-VCRedistInfo" }
        [PSCustomObject]@{ Key = "CredentialGuard";    Label = "Credential Guard Status";           Category = "Sicherheit";         Function = "Get-CredentialGuardInfo" }
        [PSCustomObject]@{ Key = "LocalAdmin";         Label = "Lokale Administratoren";            Category = "Sicherheit";         Function = "Get-LocalAdminInfo" }
        [PSCustomObject]@{ Key = "DomainTrust";        Label = "Domain Trusts & Verschlüsselung";   Category = "Active Directory";   Function = "Get-DomainTrustInfo" }
        [PSCustomObject]@{ Key = "FIPFS";              Label = "FIP-FS Scan Engine Version";        Category = "Sicherheit";         Function = "Get-FIPFSInfo" }
        [PSCustomObject]@{ Key = "SettingOverride";    Label = "Exchange Setting Overrides";        Category = "Exchange Basis";     Function = "Get-SettingOverrideInfo" }
        [PSCustomObject]@{ Key = "ServerComponent";    Label = "Server Component State";            Category = "Exchange Basis";     Function = "Get-ServerComponentStateInfo" }
        [PSCustomObject]@{ Key = "SecurityCVE";        Label = "Security CVE Prüfung";              Category = "Sicherheit";         Function = "Get-SecurityCVEInfo" }
        [PSCustomObject]@{ Key = "HttpProxy";          Label = "HTTP Proxy Konfiguration";          Category = "Sicherheit";         Function = "Get-HttpProxyInfo" }
        [PSCustomObject]@{ Key = "AntivirusProduct";  Label = "Installierte Antivirenlösung";      Category = "Sicherheit";         Function = "Get-AntivirusProductInfo" }
        [PSCustomObject]@{ Key = "NICReceiveBuffer";   Label = "NIC Receive Buffer Analyse";        Category = "Hardware & OS";      Function = "Get-NICReceiveBufferInfo" }

        # === NEU IN v1.6 ===
        [PSCustomObject]@{ Key = "MessageQueue";       Label = "Message Queue Analyse";             Category = "Mail-Flow";          Function = "Get-MessageQueueInfo" }
        [PSCustomObject]@{ Key = "MessageSizeLimits";  Label = "Message Size Limits";               Category = "Mail-Flow";          Function = "Get-MessageSizeLimitsInfo" }
        [PSCustomObject]@{ Key = "PartnerApps";        Label = "Partner Applications";              Category = "Exchange Basis";     Function = "Get-PartnerApplicationsInfo" }
        [PSCustomObject]@{ Key = "FederatedSharing";   Label = "Federated Sharing";                 Category = "Exchange Basis";     Function = "Get-FederatedSharingInfo" }
        [PSCustomObject]@{ Key = "OAuthCBA";           Label = "OAuth / Certificate Based Auth";    Category = "Sicherheit";         Function = "Get-OAuthCertificateAuthInfo" }
        [PSCustomObject]@{ Key = "ArchiveConfig";     Label = "Archive Konfiguration";              Category = "Empfänger";          Function = "Get-ArchiveConfigurationInfo" }
        [PSCustomObject]@{ Key = "CalendarResource";   Label = "Calendar & Resource Mailbox";       Category = "Empfänger";          Function = "Get-CalendarResourceConfigInfo" }
    )
}

function ConvertTo-MarkdownFromHtml {
    <#
    .SYNOPSIS
        Konvertiert die im Skript verwendeten HTML-Fragmente in Markdown.
    .PARAMETER Html
        Der zu konvertierende HTML-String.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Html
    )

    $md = $Html

    # Tabellen separat behandeln
    $tableRegex = [regex]'(?s)<table.*?>(.*?)</table>'
    $md = $tableRegex.Replace($md, {
        param($m)
        $tableContent = $m.Groups[1].Value
        $rowRegex = [regex]'(?s)<tr.*?>(.*?)</tr>'
        $rows = $rowRegex.Matches($tableContent)
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine()
        $rowIndex = 0
        foreach ($row in $rows) {
            $cellRegex = [regex]'(?s)<t[hd].*?>(.*?)</t[hd]>'
            $cells = $cellRegex.Matches($row.Groups[1].Value)
            $cellTexts = foreach ($c in $cells) {
                ($c.Groups[1].Value -replace '<.*?>', '' -replace '\|', '\|').Trim()
            }
            if ($cellTexts.Count -gt 0) {
                [void]$sb.AppendLine("| " + ($cellTexts -join " | ") + " |")
                if ($rowIndex -eq 0) {
                    [void]$sb.AppendLine("| " + (($cellTexts | ForEach-Object { "---" }) -join " | ") + " |")
                }
                $rowIndex++
            }
        }
        [void]$sb.AppendLine()
        return $sb.ToString()
    })

    # Überschriften
    $md = $md -replace '(?s)<h1[^>]*>(.*?)</h1>', "`n# `$1`n"
    $md = $md -replace '(?s)<h2[^>]*>(.*?)</h2>', "`n## `$1`n"
    $md = $md -replace '(?s)<h3[^>]*>(.*?)</h3>', "`n### `$1`n"
    $md = $md -replace '(?s)<h4[^>]*>(.*?)</h4>', "`n#### `$1`n"

    # Inline-Formatierung
    $md = $md -replace '(?s)<strong>(.*?)</strong>', '**$1**'
    $md = $md -replace '(?s)<em>(.*?)</em>', '*$1*'
    $md = $md -replace '<br\s*/?>', "`n"
    $md = $md -replace '(?s)<p[^>]*>(.*?)</p>', "`n`$1`n"

    # Restliche Tags entfernen
    $md = $md -replace '<[^>]+>', ''

    # HTML-Entities dekodieren
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    $md = [System.Web.HttpUtility]::HtmlDecode($md)

    # Mehrfache Leerzeilen reduzieren
    $md = $md -replace '(\r?\n){3,}', "`n`n"

    return $md.Trim()
}

function Build-MarkdownDocument {
    <#
    .SYNOPSIS
        Baut ein Markdown-Dokument aus den gesammelten HTML-Sektionen.
    #>
    Write-Log -Message "=== Erstelle Markdown-Dokument ===" -Level "INFO"

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("# $script:DocTitle")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("**$script:DocSubTitle**")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("- Erstellt am: $script:DateOnly")
    [void]$sb.AppendLine("- Erstellt von: $script:DocAuthor auf $script:DocComputerName")
    [void]$sb.AppendLine("- Dokumentierte Server: $($ExchangeServers -join ', ')")
    [void]$sb.AppendLine("- Erkannte Edition: $script:ExchangeEdition")
    [void]$sb.AppendLine("- Fehler: $($script:ErrorCount) | Warnungen: $($script:WarningCount)")
    [void]$sb.AppendLine()

    # Inhaltsverzeichnis
    [void]$sb.AppendLine("## Inhaltsverzeichnis")
    [void]$sb.AppendLine()
    foreach ($entry in $script:TOCEntries) {
        [void]$sb.AppendLine("- $($entry.Title)")
    }
    [void]$sb.AppendLine()

    # Sektionen konvertieren
    foreach ($section in $script:HTMLSections) {
        [void]$sb.AppendLine((ConvertTo-MarkdownFromHtml -Html $section))
        [void]$sb.AppendLine()
    }

    return $sb.ToString()
}

function Export-DocumentToPdf {
    <#
    .SYNOPSIS
        Konvertiert eine HTML-Datei in PDF. Nutzt zuerst Microsoft Word (COM),
        bei Nichtverfügbarkeit Microsoft Edge / Chrome im Headless-Modus.
    .PARAMETER HtmlPath
        Pfad zur HTML-Quelldatei.
    .PARAMETER PdfPath
        Ziel-Pfad der PDF-Datei.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$HtmlPath,
        [Parameter(Mandatory = $true)] [string]$PdfPath
    )

    # --- Methode 1: Microsoft Word COM ---
    $word = $null
    $doc = $null
    
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        $doc = $word.Documents.Open($HtmlPath)
        # 17 = wdFormatPDF
        $doc.SaveAs([ref]$PdfPath, [ref]17)
        $doc.Close($false)
        Write-Log -Message "PDF über Microsoft Word erstellt: $PdfPath" -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Word-COM PDF-Export nicht möglich: $($_.Exception.Message). Versuche Browser-Headless..." -Level "WARNING"
    }
    finally {
        # Cleanup - in separaten Try-Catch um unerwartete Fehler abzufangen
        try {
            if ($doc) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            if ($word) {
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        }
        catch {
            # Fehler beim Cleanup - ignorieren (Word-Prozess läuft trotzdem weiter)
            [void]0
        }
        # Erzwinge Garbage Collection um Word-Prozesse freizugeben
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # --- Methode 2: Edge / Chrome Headless ---
    $browsers = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    foreach ($browser in $browsers) {
        if (Test-Path $browser) {
            try {
                $uri = ([System.Uri]$HtmlPath).AbsoluteUri
                $arguments = "--headless --disable-gpu --print-to-pdf=`"$PdfPath`" --no-margins `"$uri`""
                Start-Process -FilePath $browser -ArgumentList $arguments -Wait -NoNewWindow
                if (Test-Path $PdfPath) {
                    Write-Log -Message "PDF über Browser-Headless erstellt: $PdfPath" -Level "INFO"
                    return $true
                }
            }
            catch {
                Write-Log -Message "Browser-Headless PDF-Export fehlgeschlagen ($browser): $_" -Level "WARNING"
            }
        }
    }

    Write-Log -Message "PDF-Export nicht möglich (weder Word noch Edge/Chrome verfügbar)." -Level "ERROR"
    return $false
}

function Show-DocumentationGui {
    <#
    .SYNOPSIS
        Zeigt eine WPF-GUI zur Auswahl von Servern, Ausgabepfad, Sektionen und
        Ausgabeformaten. Server werden automatisch erkannt und als Checkboxen angeboten.
        Sektionen sind nach Themenbereichen gruppiert.
    #>
    [CmdletBinding()]
    param(
        [string]$DefaultServers,
        [string]$DefaultCompany,
        [string]$DefaultOutputPath
    )

    try {
        # WPF Assemblies laden
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction Stop
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    }
    catch {
        Write-Host "FEHLER: WPF-Assemblies konnten nicht geladen werden: $_" -ForegroundColor Red
        return $null
    }

    # --- Exchange Server automatisch erkennen ---
    $detectedServers = @()
    try {
        $snapInLoaded = Get-PSSnapin -Name "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue
        $moduleLoaded = Get-Module -Name "ExchangeManagementTools" -ErrorAction SilentlyContinue
        if (-not $snapInLoaded -and -not $moduleLoaded) {
            if (Get-PSSnapin -Registered -Name "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue) {
                Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue
            }
            elseif (Get-Module -ListAvailable -Name "ExchangeManagementTools" -ErrorAction SilentlyContinue) {
                Import-Module "ExchangeManagementTools" -ErrorAction SilentlyContinue
            }
        }
        $detectedServers = @(Get-ExchangeServer -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
    }
    catch {
        # Keine Server gefunden - manuelle Eingabe nötig
    }

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Server Dokumentation v1.7" Height="920" Width="1120"
        WindowStartupLocation="CenterScreen" Background="#EDEDF2" ResizeMode="CanResize" MinWidth="900" MinHeight="700"
        FontFamily="Segoe UI, Arial, sans-serif">
    <Window.Resources>
        <!-- Farben -->
        <Color x:Key="PrimaryColor">#1A237E</Color>
        <Color x:Key="PrimaryLight">#3949AB</Color>
        <Color x:Key="AccentColor">#00BCD4</Color>
        <Color x:Key="SuccessColor">#43A047</Color>
        <Color x:Key="ErrorColor">#E53935</Color>

        <SolidColorBrush x:Key="PrimaryBrush" Color="#1A237E"/>
        <SolidColorBrush x:Key="PrimaryLightBrush" Color="#3949AB"/>
        <SolidColorBrush x:Key="PrimaryGradientEnd" Color="#283593"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#00BCD4"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#43A047"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#E53935"/>
        <SolidColorBrush x:Key="CardBg" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PageBg" Color="#EDEDF2"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#212121"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#757575"/>

        <!-- Modernes GroupBox-Style (Card) -->
        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,14"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="16" ShadowDepth="2" Color="#1A000000" Opacity="0.12"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Button Primary Style -->
        <Style x:Key="PrimaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="44"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="8" Background="#1A237E" BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#283593"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="12" ShadowDepth="3" Color="#FF1A237E" Opacity="0.4"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#0D154A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Button Secondary Style -->
        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="44"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="Foreground" Value="#1A237E"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="8" Background="White" BorderBrush="#1A237E" BorderThickness="1.5">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F0F0FF"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#1A000000" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E0E0F0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern CheckBox Style -->
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="#212121"/>
            <Setter Property="Margin" Value="0,3,0,3"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Border x:Name="CheckBorder" Width="20" Height="20" CornerRadius="4" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                                    <Path x:Name="CheckMark" Width="12" Height="12" Stretch="Uniform" Fill="#1A237E" Data="M 0 6 L 4 10 L 10 2" Visibility="Hidden" Margin="0,0,0,0"/>
                                </Border>
                            </BulletDecorator.Bullet>
                            <ContentPresenter Margin="8,0,0,0" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#E8EAF6"/>
                                <Setter TargetName="CheckBorder" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckBorder" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="CheckBorder" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#1A000000" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Toggle Button Style für Formate -->
        <Style x:Key="FormatToggleStyle" TargetType="ToggleButton">
            <Setter Property="Height" Value="42"/>
            <Setter Property="Width" Value="110"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,14,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border x:Name="border" CornerRadius="8" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                                <TextBlock x:Name="Icon" Text="●" FontSize="10" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="#BDBDBD"/>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E8EAF6"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="Icon" Property="Foreground" Value="#1A237E"/>
                                <Setter TargetName="Icon" Property="Text" Value="◆"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#1A000000" Opacity="0.08"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Modern Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="10,0"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="BorderBrush" Value="#BDBDBD"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="10,0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="1" Color="#FF1A237E" Opacity="0.15"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ScrollViewer Modern Style -->
        <Style x:Key="ModernScrollViewer" TargetType="ScrollViewer">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid>
                            <ScrollContentPresenter Margin="{TemplateBinding Padding}"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar" HorizontalAlignment="Right" Width="8" Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}">
                                <ScrollBar.Template>
                                    <ControlTemplate TargetType="ScrollBar">
                                        <Grid>
                                            <Track x:Name="PART_Track">
                                                <Track.Thumb>
                                                    <Thumb>
                                                        <Thumb.Template>
                                                            <ControlTemplate TargetType="Thumb">
                                                                <Border CornerRadius="4" Background="#C0C0C0" Margin="0,2"/>
                                                            </ControlTemplate>
                                                        </Thumb.Template>
                                                    </Thumb>
                                                </Track.Thumb>
                                            </Track>
                                        </Grid>
                                    </ControlTemplate>
                                </ScrollBar.Template>
                            </ScrollBar>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- ===== HEADER ===== -->
        <Border Grid.Row="0" CornerRadius="12" Margin="0,0,0,20" BorderThickness="0">
            <Border.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#1A237E" Offset="0.0"/>
                    <GradientStop Color="#283593" Offset="0.5"/>
                    <GradientStop Color="#3949AB" Offset="1.0"/>
                </LinearGradientBrush>
            </Border.Background>
            <Border.Effect>
                <DropShadowEffect BlurRadius="20" ShadowDepth="4" Color="#FF1A237E" Opacity="0.35"/>
            </Border.Effect>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <!-- Icon Bereich -->
                <Border Grid.Column="0" Width="56" Height="56" CornerRadius="10" Background="#FFFFFF22" Margin="16,16,0,16">
                    <TextBlock Text="📊" FontSize="28" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <!-- Titel -->
                <StackPanel Grid.Column="1" Margin="16,16,0,16" VerticalAlignment="Center">
                    <TextBlock Text="Exchange Server Dokumentation" FontSize="26" FontWeight="SemiBold" Foreground="White"/>
                    <TextBlock Text="System-Checks · TLS/SSL · Multi-Format Export · 13 neue Prüfungen" Foreground="#90CAF9" FontSize="13" Margin="0,4,0,0"/>
                </StackPanel>
                <!-- Version Badge -->
                <Border Grid.Column="2" CornerRadius="6" Background="#FFFFFF22" Padding="10,6" Margin="0,0,16,0" VerticalAlignment="Center">
                    <TextBlock Text="v1.7" Foreground="Black" FontWeight="Bold" FontSize="13"/>
                </Border>
            </Grid>
        </Border>

        <!-- ===== SERVER AUSWAHL ===== -->
        <Border Grid.Row="1" Style="{StaticResource CardStyle}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="🖥️" FontSize="18" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBlock Text="Exchange Server" FontSize="16" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" Orientation="Horizontal">
                        <CheckBox Name="ChkSelectAllServers" Content="Alle auswählen" Style="{StaticResource ModernCheckBox}" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    </StackPanel>
                </Grid>
                <Border Grid.Row="1" CornerRadius="8" Background="#F5F5F8" Padding="14,12" MinHeight="40">
                    <Grid>
                        <WrapPanel Grid.Row="0" Name="ServerPanel" Orientation="Horizontal"/>
                        <TextBox Grid.Row="0" Name="TxtServersManual" Height="32" Visibility="Collapsed" VerticalContentAlignment="Center" Margin="0,0,0,0" Style="{StaticResource ModernTextBox}"/>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- ===== ORGANISATION & AUSGABEPFAD ===== -->
        <Border Grid.Row="2" Style="{StaticResource CardStyle}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <!-- Organisation -->
                <StackPanel Grid.Column="0">
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="🏢" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="Organisation" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBox Name="TxtCompany" Style="{StaticResource ModernTextBox}"/>
                </StackPanel>
                <!-- Ausgabepfad -->
                <StackPanel Grid.Column="2">
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="📁" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="Ausgabepfad" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Name="TxtOutputPath" Style="{StaticResource ModernTextBox}"/>
                        <Button Grid.Column="1" Name="BtnBrowse" Content="📂" Width="40" Height="36" Margin="10,0,0,0" Padding="0"
                                FontSize="18" Cursor="Hand" BorderThickness="1.5" BorderBrush="#BDBDBD" Background="White"
                                ToolTip="Ausgabeverzeichnis wählen">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border x:Name="border" CornerRadius="6" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="border" Property="Background" Value="#F5F5FF"/>
                                            <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter TargetName="border" Property="Background" Value="#E8EAF6"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                    </Grid>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ===== AUSGABEFORMATE ===== -->
        <Border Grid.Row="3" Style="{StaticResource CardStyle}">
            <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Text="📄" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Text="Ausgabeformate" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <ToggleButton Name="ChkHtml" Content="HTML" Style="{StaticResource FormatToggleStyle}" IsChecked="True"/>
                    <ToggleButton Name="ChkPdf" Content="PDF" Style="{StaticResource FormatToggleStyle}"/>
                    <ToggleButton Name="ChkMarkdown" Content="Markdown" Style="{StaticResource FormatToggleStyle}"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- =Category Header== -->
        <Border Grid.Row="4" Style="{StaticResource CardStyle}" Padding="16,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="📋" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Text="Dokumentationsbereiche" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                </StackPanel>
                <CheckBox Grid.Column="2" Name="ChkSelectAll" Content="Alle auswählen" Style="{StaticResource ModernCheckBox}" FontWeight="SemiBold" IsChecked="True"/>
            </Grid>
        </Border>

        <!-- ===== DOKUMENTATIONSBEREICHE ===== -->
        <Border Grid.Row="5" Style="{StaticResource CardStyle}" Padding="0" Margin="0,0,0,14">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Style="{StaticResource ModernScrollViewer}" Padding="16,12">
                <StackPanel Name="CategoryPanel" Orientation="Vertical"/>
            </ScrollViewer>
        </Border>

        <!-- ===== STATUS ===== -->
        <Border Grid.Row="6" CornerRadius="8" Padding="12,10" Margin="0,0,0,12" Background="#FFF3E0" BorderThickness="0" Visibility="Collapsed" Name="StatusBorder">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="ℹ️" FontSize="16" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Grid.Column="1" Name="TxtStatus" Foreground="#E65100" FontWeight="SemiBold" TextWrapping="Wrap" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- ===== BUTTONS ===== -->
        <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,0,0">
            <Button Name="BtnCancel" Content="Abbrechen" Width="130" Style="{StaticResource SecondaryButtonStyle}" Margin="0,0,14,0"/>
            <Button Name="BtnStart" Content="  🚀  Dokumentation starten" Width="260" Style="{StaticResource PrimaryButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
'@

    try {
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        Write-Host "FEHLER beim Laden der XAML: $_" -ForegroundColor Red
        return $null
    }

    # Steuerelemente referenzieren
    try {
        $chkSelectAllServers = $window.FindName("ChkSelectAllServers")
        $serverPanel         = $window.FindName("ServerPanel")
        $txtServersManual    = $window.FindName("TxtServersManual")
        $txtCompany          = $window.FindName("TxtCompany")
        $txtOutputPath       = $window.FindName("TxtOutputPath")
        $btnBrowse           = $window.FindName("BtnBrowse")
        $chkHtml             = $window.FindName("ChkHtml")
        $chkPdf              = $window.FindName("ChkPdf")
        $chkMarkdown         = $window.FindName("ChkMarkdown")
        $chkSelectAll        = $window.FindName("ChkSelectAll")
        $categoryPanel       = $window.FindName("CategoryPanel")
        $txtStatus           = $window.FindName("TxtStatus")
        $statusBorder        = $window.FindName("StatusBorder")
        $btnStart            = $window.FindName("BtnStart")
        $btnCancel           = $window.FindName("BtnCancel")
    }
    catch {
        Write-Host "FEHLER beim Zugriff auf GUI-Steuerelemente: $_" -ForegroundColor Red
        return $null
    }

    # Vorbelegung
    $txtCompany.Text    = $DefaultCompany
    $txtOutputPath.Text = $DefaultOutputPath

    # Server-Checkboxen
    $serverCheckboxes = @{}
    $modernCheckboxStyle = $window.FindResource("ModernCheckBox")
    if ($detectedServers.Count -gt 0) {
        foreach ($srv in $detectedServers) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $srv
            $cb.Tag = $srv
            $cb.IsChecked = $true
            $cb.Margin = "0,0,20,6"
            $cb.FontSize = 13
            $cb.Style = $modernCheckboxStyle
            [void]$serverPanel.Children.Add($cb)
            $serverCheckboxes[$srv] = $cb
        }
        $chkSelectAllServers.IsChecked = $true
    }
    else {
        $serverPanel.Visibility = "Collapsed"
        $txtServersManual.Visibility = "Visible"
        $txtServersManual.Text = $DefaultServers
        $chkSelectAllServers.Visibility = "Collapsed"
    }

    $chkSelectAllServers.Add_Click({
        $state = $chkSelectAllServers.IsChecked
        foreach ($cb in $serverPanel.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox]) { $cb.IsChecked = $state }
        }
    })

    # Sektions-Checkboxen
    $sectionCheckboxes = @{}
    $registry = Get-DocSectionRegistry
    $categories = $registry | Group-Object -Property Category

    foreach ($cat in $categories) {
        $border = New-Object System.Windows.Controls.Border
        $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $border.Background = [System.Windows.Media.Brushes]::White
        $border.BorderThickness = [System.Windows.Thickness]::new(0)
        $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
        $border.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{
            BlurRadius = 12
            ShadowDepth = 1
            Color = [System.Windows.Media.Color]::FromArgb(26, 0, 0, 0)
            Opacity = 0.1
        }

        $outerSp = New-Object System.Windows.Controls.StackPanel
        $outerSp.Orientation = "Vertical"

        # Category Header
        $headerBorder = New-Object System.Windows.Controls.Border
        $headerBorder.CornerRadius = [System.Windows.CornerRadius]::new(8, 8, 0, 0)
        $headerBorder.Background = [System.Windows.Media.Brushes]::White
        $headerBorder.BorderThickness = [System.Windows.Thickness]::new(0)
        $headerBorder.Padding = [System.Windows.Thickness]::new(12, 10, 12, 8)

        $headerGrid = New-Object System.Windows.Controls.Grid
        $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))
        $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

        $iconMap = @{
            "Hardware & OS"      = "🖥️"
            "Active Directory"   = "🏛️"
            "Exchange Basis"     = "⚙️"
            "Mail-Flow"          = "📨"
            "DNS & Records"      = "🌐"
            "Empfänger"          = "👤"
            "Sicherheit"         = "🔒"
            "Sicherheit & TLS"   = "🔐"
            "Compliance"         = "📋"
            "Hybrid / Cloud"     = "☁️"
        }
        $catIcon = if ($iconMap.ContainsKey($cat.Name)) { $iconMap[$cat.Name] } else { "📋" }

        $iconText = New-Object System.Windows.Controls.TextBlock
        $iconText.Text = "$catIcon  $($cat.Name)"
        $iconText.FontSize = 14
        $iconText.FontWeight = [System.Windows.FontWeights]::SemiBold
        $iconText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 33, 33, 33))
        $iconText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$headerGrid.Children.Add($iconText)

        $headerBorder.Child = $headerGrid
        [void]$outerSp.Children.Add($headerBorder)

        # Separator
        $sep = New-Object System.Windows.Controls.Border
        $sep.Height = 1
        $sep.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(40, 0, 0, 0))
        $sep.Margin = [System.Windows.Thickness]::new(12, 0, 12, 0)
        [void]$outerSp.Children.Add($sep)

        # Content
        $contentBorder = New-Object System.Windows.Controls.Border
        $contentBorder.CornerRadius = [System.Windows.CornerRadius]::new(0, 0, 8, 8)
        $contentBorder.Background = [System.Windows.Media.Brushes]::White
        $contentBorder.Padding = [System.Windows.Thickness]::new(12, 8, 12, 10)

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = "Vertical"

        foreach ($section in $cat.Group) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $section.Label
            $cb.Tag = $section.Key
            $cb.IsChecked = $true
            $cb.Margin = "0,3,0,3"
            $cb.FontSize = 12
            $cb.Style = $modernCheckboxStyle
            [void]$sp.Children.Add($cb)
            $sectionCheckboxes[$section.Key] = $cb
        }
        $contentBorder.Child = $sp
        [void]$outerSp.Children.Add($contentBorder)
        $border.Child = $outerSp
        [void]$categoryPanel.Children.Add($border)
    }

    $chkSelectAll.Add_Click({
        $state = $chkSelectAll.IsChecked
        foreach ($key in $sectionCheckboxes.Keys) {
            $sectionCheckboxes[$key].IsChecked = $state
        }
    })

    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Ausgabeverzeichnis wählen"
        if ($txtOutputPath.Text -and (Test-Path $txtOutputPath.Text)) { $dlg.SelectedPath = $txtOutputPath.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutputPath.Text = $dlg.SelectedPath
        }
    })

    $script:GuiResult = $null

    # Hilfsfunktion für Status-Anzeige
    $script:ShowStatus = {
        param([string]$Message, [string]$Type = "error")
        $txtStatus.Text = $Message
        $statusBorder.Visibility = "Visible"
        switch ($Type) {
            "error"   { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 255, 235, 238)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 198, 40, 40)) }
            "success" { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 232, 245, 233)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 46, 125, 50)) }
            "info"    { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 227, 242, 253)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 21, 101, 192)) }
        }
    }
    $script:HideStatus = {
        $statusBorder.Visibility = "Collapsed"
    }

    $btnStart.Add_Click({
        # Status verstecken
        $statusBorder.Visibility = "Collapsed"

        $selectedServers = @()
        if ($detectedServers.Count -gt 0) {
            foreach ($srv in $serverCheckboxes.Keys) {
                if ($serverCheckboxes[$srv].IsChecked) { $selectedServers += $srv }
            }
        }
        else {
            $selectedServers = @($txtServersManual.Text -split '[,;\s]+' | Where-Object { $_ })
        }

        if ($selectedServers.Count -eq 0) {
            & $script:ShowStatus -Message "Fehler: Bitte mindestens einen Exchange-Server auswählen." -Type "error"
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
            & $script:ShowStatus -Message "Fehler: Bitte einen Ausgabepfad angeben." -Type "error"
            return
        }

        $selectedSections = @()
        foreach ($key in $sectionCheckboxes.Keys) {
            if ($sectionCheckboxes[$key].IsChecked) { $selectedSections += $key }
        }
        if ($selectedSections.Count -eq 0) {
            & $script:ShowStatus -Message "Fehler: Bitte mindestens einen Dokumentationsbereich auswählen." -Type "error"
            return
        }

        $formats = @()
        if ($chkHtml.IsChecked)     { $formats += "HTML" }
        if ($chkPdf.IsChecked)      { $formats += "PDF" }
        if ($chkMarkdown.IsChecked) { $formats += "Markdown" }
        if ($formats.Count -eq 0) {
            & $script:ShowStatus -Message "Fehler: Bitte mindestens ein Ausgabeformat wählen." -Type "error"
            return
        }

        $script:GuiResult = @{
            Servers    = $selectedServers
            Company    = $txtCompany.Text
            OutputPath = $txtOutputPath.Text
            Sections   = $selectedSections
            Formats    = $formats
        }
        $window.Close()
    })

    $btnCancel.Add_Click({
        $script:GuiResult = $null
        $window.Close()
    })

    [void]$window.ShowDialog()
    return $script:GuiResult
}

#endregion

#region ============================================================
# HAUPTPROGRAMM
#endregion ============================================================

try {
    # ============================================================
    # 0. EINGABEN ERMITTELN (GUI ODER PARAMETER)
    # ============================================================

    # System.Web Assembly laden (für HtmlEncode in ConvertTo-HTMLTable)
    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch {}

    # Entscheiden ob GUI angezeigt wird:
    # - explizit via -ShowGui ODER
    # - keine Server angegeben und nicht -NoGui
    $useGui = $ShowGui -or (-not $ExchangeServers -and -not $NoGui)

    # Standardauswahl der Sektionen (alle), falls keine angegeben
    $selectedSectionKeys = if ($Sections) { $Sections } else { (Get-DocSectionRegistry).Key }

    if ($useGui) {
        $guiConfig = Show-DocumentationGui `
            -DefaultServers ($ExchangeServers -join ", ") `
            -DefaultCompany $CompanyName `
            -DefaultOutputPath $OutputPath

        if (-not $guiConfig) {
            Write-Host "Abgebrochen durch Benutzer (GUI)." -ForegroundColor Yellow
            return
        }

        # GUI-Auswahl übernehmen
        $ExchangeServers     = $guiConfig.Servers
        $CompanyName         = $guiConfig.Company
        $OutputPath          = $guiConfig.OutputPath
        $selectedSectionKeys = $guiConfig.Sections
        $OutputFormats       = $guiConfig.Formats
    }

    # Validierung
    if (-not $ExchangeServers -or $ExchangeServers.Count -eq 0) {
        throw "Keine Exchange-Server angegeben. Bitte -ExchangeServers verwenden oder die GUI nutzen."
    }

    # ============================================================
    # 0b. PFADE UND VARIABLEN (NEU) BERECHNEN
    # ============================================================
    $script:LogPath        = $OutputPath
    $stamp                 = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile        = Join-Path -Path $LogPath -ChildPath "Exchange-Dokumentation_$stamp.log"
    $script:HTMLOutputFile = Join-Path -Path $LogPath -ChildPath "Exchange_Dokumentation_$stamp.html"
    $script:PDFOutputFile  = Join-Path -Path $LogPath -ChildPath "Exchange_Dokumentation_$stamp.pdf"
    $script:MDOutputFile   = Join-Path -Path $LogPath -ChildPath "Exchange_Dokumentation_$stamp.md"
    $script:DocSubTitle    = $CompanyName

    # ============================================================
    # 1. VORBEREITUNG
    # ============================================================
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        Write-Host "Ausgabeverzeichnis erstellt: $LogPath" -ForegroundColor Cyan
    }

    # Logging starten
    Write-Log -Message "=============================================" -Level "INFO"
    Write-Log -Message "Exchange Dokumentation gestartet (v1.7 - 2019/SE)" -Level "INFO"
    Write-Log -Message "Zielserver: $($ExchangeServers -join ', ')" -Level "INFO"
    Write-Log -Message "Ausgabepfad: $LogPath" -Level "INFO"
    Write-Log -Message "Ausgabeformate: $($OutputFormats -join ', ')" -Level "INFO"
    Write-Log -Message "Gewählte Sektionen: $($selectedSectionKeys.Count)" -Level "INFO"
    Write-Log -Message "Verbindungsmodus: CIM mit automatischem DCOM Fallback" -Level "INFO"
    Write-Log -Message "=============================================" -Level "INFO"

    # ============================================================
    # 2. EXCHANGE SNAP-IN / MODUL LADEN
    # ============================================================
    $exchangeLoaded = Initialize-ExchangeEnvironment
    if (-not $exchangeLoaded) {
        Write-Log -Message "Exchange Management Tools nicht verfügbar. Skript wird beendet." -Level "ERROR"
        throw "Exchange Management Tools nicht verfügbar."
    }

    # Edition erkennen (2019 vs. SE)
    Get-ExchangeEdition

    # DocTitle dynamisch anpassen
    $script:DocTitle = "Exchange Server $($script:ExchangeEdition) - Umgebungsdokumentation"

    # ============================================================
    # 3. KONNEKTIVITÄTSTEST
    # ============================================================
    Write-Log -Message "=== Prüfe Erreichbarkeit der Server ===" -Level "INFO"
    foreach ($server in $ExchangeServers) {
        try {
            if (-not (Test-Connection -ComputerName $server -Count 2 -Quiet -ErrorAction Stop)) {
                Write-Log -Message "Server $server ist NICHT erreichbar!" -Level "WARNING"
            }
            else {
                Write-Log -Message "Server $server ist erreichbar (Ping OK)." -Level "INFO"
            }
        }
        catch {
            Write-Log -Message "Ping-Test für ${server} fehlgeschlagen: $_" -Level "WARNING"
        }
    }

    # ============================================================
    # 4. DATENSAMMLUNG - NUR AUSGEWÄHLTE SEKTIONEN
    # ============================================================
    Write-Log -Message "=== Starte Datensammlung ($($selectedSectionKeys.Count) Sektionen) ===" -Level "INFO"

    $registry = Get-DocSectionRegistry
    foreach ($section in $registry) {
        if ($selectedSectionKeys -contains $section.Key) {
            try {
                Write-Log -Message "Sektion: $($section.Label) [$($section.Key)]" -Level "INFO"
                & $section.Function
            }
            catch {
                Write-Log -Message "Fehler in Sektion $($section.Key): $_" -Level "ERROR"
            }
        }
    }

    # ============================================================
    # 5. DOKUMENTE ERZEUGEN UND SPEICHERN (PRO FORMAT)
    # ============================================================
    $createdFiles = [System.Collections.ArrayList]::new()

    # HTML wird immer für PDF benötigt - bei PDF ohne HTML temporär erzeugen
    $needHtml = ($OutputFormats -contains "HTML") -or ($OutputFormats -contains "PDF")

    if ($needHtml) {
        Write-Log -Message "=== Generiere HTML-Dokument ===" -Level "INFO"
        $finalHTML = Build-HTMLDocument
        $finalHTML | Out-File -FilePath $script:HTMLOutputFile -Encoding UTF8 -Force
        if ($OutputFormats -contains "HTML") {
            [void]$createdFiles.Add($script:HTMLOutputFile)
        }
    }

    if ($OutputFormats -contains "PDF") {
        Write-Log -Message "=== Generiere PDF-Dokument ===" -Level "INFO"
        if (Export-DocumentToPdf -HtmlPath $script:HTMLOutputFile -PdfPath $script:PDFOutputFile) {
            [void]$createdFiles.Add($script:PDFOutputFile)
        }
        # Temporäres HTML entfernen, falls HTML nicht gewünscht war
        if ($OutputFormats -notcontains "HTML" -and (Test-Path $script:HTMLOutputFile)) {
            Remove-Item -Path $script:HTMLOutputFile -Force -ErrorAction SilentlyContinue
        }
    }

    if ($OutputFormats -contains "Markdown") {
        Write-Log -Message "=== Generiere Markdown-Dokument ===" -Level "INFO"
        $finalMD = Build-MarkdownDocument
        $finalMD | Out-File -FilePath $script:MDOutputFile -Encoding UTF8 -Force
        [void]$createdFiles.Add($script:MDOutputFile)
    }

    Write-Log -Message "=============================================" -Level "INFO"
    Write-Log -Message "Dokumentation erfolgreich erstellt!" -Level "INFO"
    foreach ($f in $createdFiles) { Write-Log -Message "Datei: $f" -Level "INFO" }
    Write-Log -Message "Log-Datei:  $($script:LogFile)" -Level "INFO"
    Write-Log -Message "Fehler: $($script:ErrorCount) | Warnungen: $($script:WarningCount)" -Level "INFO"
    Write-Log -Message "=============================================" -Level "INFO"

    # Konsolenausgabe
    Write-Host "`n" -NoNewline
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "  Exchange Dokumentation abgeschlossen! (v1.7 - 2019/SE)" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan
    foreach ($f in $createdFiles) {
        Write-Host "  Datei: $f" -ForegroundColor White
    }
    Write-Host "  Log-Datei:  $($script:LogFile)" -ForegroundColor White
    Write-Host "  Fehler:     $($script:ErrorCount) | Warnungen: $($script:WarningCount)" -ForegroundColor White
    Write-Host "===============================================================" -ForegroundColor Cyan

    # Datei optional öffnen
    if ($createdFiles.Count -gt 0) {
        $openFile = Read-Host "`nDokumentation jetzt oeffnen? (J/N)"
        if ($openFile -eq "J" -or $openFile -eq "j") {
            Start-Process $createdFiles[0]
        }
    }
}
catch {
    Write-Log -Message "KRITISCHER FEHLER: $_" -Level "ERROR"
    Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Host "`nKritischer Fehler! Details: $($script:LogFile)" -ForegroundColor Red
}
finally {
    Write-Log -Message "Skript beendet um $(Get-Date -Format 'HH:mm:ss')" -Level "INFO"
}

#endregion


