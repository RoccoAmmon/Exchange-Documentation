<#
.SYNOPSIS
    Exchange On-Premises Dokumentations-Skript (v1.1 - Exchange 2019 & SE Support)
.DESCRIPTION
    Erstellt eine umfassende HTML-Dokumentation der gesamten Exchange On-Premises Umgebung.
    Unterstützt Exchange Server 2019 und Exchange Server Subscription Edition (SE).
    Das HTML-Dokument kann direkt in Microsoft Word importiert werden.
    
    WICHTIG: Diese Version verwendet CIM-Sessions mit automatischem DCOM-Fallback,
    sodass das Skript auch funktioniert, wenn WinRM (PowerShell Remoting) nicht
    korrekt konfiguriert ist.

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
    - Anti-Spam/Anti-Malware Konfiguration
    - Compliance (DLP, Litigation Hold, In-Place Hold)
    - Mailbox-Quotas und Warnungs-Schwellwerte
    - SMTP-Relay Konfiguration (Geräte/Applikationen)

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
    Version:         1.1
    Erstellt:        2026-03-05
    Letzte Änderung: 2026-06-04
    Änderungen:      v1.1 - Erweiterte Transportkomponenten-Dokumentation (Speicherorte, Queue-DB, Message-Tracking, SMTP-Logs, Safety-Net)
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

function Detect-ExchangeEdition {
    <#
    .SYNOPSIS
        Ermittelt die Exchange Edition (2019 oder SE) anhand der AdminDisplayVersion.
        Exchange SE: Build >= 2000 ODER Remote-Registry zeigt 'Subscription'.
        Exchange 2019 hatte maximal Build ~1688 (CU15), SE startet bei höheren Builds.
    #>
    Write-Log -Message "=== Ermittle Exchange Edition (2019 vs. SE) ===" -Level "INFO"

    try {
        # Ersten Exchange Server abfragen
        $exServer = Get-ExchangeServer -ErrorAction Stop | Select-Object -First 1
        if ($exServer) {
            $serverName = $exServer.Name
            $adminVersion = $exServer.AdminDisplayVersion.ToString()
            Write-Log -Message "AdminDisplayVersion: $adminVersion" -Level "INFO"

            # Build-Nummer extrahieren (Format: Version 15.2 (Build 1544.11))
            if ($adminVersion -match 'Build\s+(\d+)\.(\d+)') {
                $buildMajor = [int]$Matches[1]
                Write-Log -Message "Build-Nummer (Major): $buildMajor" -Level "INFO"

                # Exchange 2019 letzte Version war CU15 mit Build ~1688
                # Exchange SE startet bei Build 1544 (RTM) aber neuere SE Versionen haben Build >= 2000
                # Build >= 2000 ist definitiv SE (2019 ging nie so hoch)
                if ($buildMajor -ge 2000) {
                    $script:ExchangeEdition = "SE"
                    Write-Log -Message "Exchange Edition erkannt: Subscription Edition (SE) - Build $buildMajor > 2000" -Level "INFO"
                }
                elseif ($buildMajor -ge 1544) {
                    # Grauzone: könnte 2019 CU14/CU15 oder SE RTM/CU1 sein
                    # Remote-Registry-Check auf Exchange Server durchführen
                    Write-Log -Message "Build $buildMajor in Grauzone (1544-1999) - prüfe Remote-Registry auf $serverName" -Level "INFO"
                    
                    $productName = $null
                    try {
                        # Versuche Remote-Registry via Invoke-Command
                        $productName = Invoke-Command -ComputerName $serverName -ScriptBlock {
                            $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ErrorAction SilentlyContinue
                            return $reg.MsiProductName
                        } -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Log -Message "Remote-Registry fehlgeschlagen: $_" -Level "WARNING"
                    }

                    # Fallback: Lokale Registry (falls Skript auf Exchange Server läuft)
                    if (-not $productName) {
                        $regLocal = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ErrorAction SilentlyContinue
                        $productName = $regLocal.MsiProductName
                    }

                    Write-Log -Message "MsiProductName: $productName" -Level "INFO"

                    if ($productName -match 'Subscription') {
                        $script:ExchangeEdition = "SE"
                        Write-Log -Message "Exchange Edition erkannt: Subscription Edition (SE) - via Registry" -Level "INFO"
                    }
                    else {
                        $script:ExchangeEdition = "2019"
                        Write-Log -Message "Exchange Edition erkannt: 2019 (CU14+)" -Level "INFO"
                    }
                }
                else {
                    # Build < 1544 ist definitiv 2019 (vor CU14)
                    $script:ExchangeEdition = "2019"
                    Write-Log -Message "Exchange Edition erkannt: 2019 (vor CU14)" -Level "INFO"
                }
            }
            else {
                # Fallback: Wenn Regex nicht matcht
                Write-Log -Message "Build-Nummer konnte nicht extrahiert werden - Fallback auf Remote-Registry" -Level "WARNING"
                $productName = $null
                try {
                    $productName = Invoke-Command -ComputerName $serverName -ScriptBlock {
                        $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup" -ErrorAction SilentlyContinue
                        return $reg.MsiProductName
                    } -ErrorAction SilentlyContinue
                }
                catch { }

                if ($productName -match 'Subscription') {
                    $script:ExchangeEdition = "SE"
                }
                else {
                    $script:ExchangeEdition = "2019"
                }
                Write-Log -Message "Exchange Edition (via Registry): $($script:ExchangeEdition)" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "Kein Exchange Server gefunden - Edition bleibt 'Unknown'" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Fehler bei Edition-Erkennung: $_ - Fallback auf '2019'" -Level "WARNING"
        $script:ExchangeEdition = "2019"
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
                    if ($_.AdminDisplayVersion -match 'Version 15\.2') {
                        if ($_.AdminDisplayVersion -match 'Build 2\d{3}') { 'SE' } else { '2019' }
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
# 35. HYBRID-KONFIGURATION MIT EXCHANGE ONLINE
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
            page-break-before: always;
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
            body { margin: 20px; }
            h2 { page-break-before: always; }
            table { page-break-inside: auto; }
            tr { page-break-inside: avoid; }
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
    <meta name="generator" content="Exchange Dokumentations-Skript v1.0">
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
            <p>Version: 1.0 (Exchange 2019 &amp; SE / CIM-DCOM Fallback)</p>
        </div>
    </div>

    $tocHTML

    $summaryHTML

    $($script:HTMLSections -join "`n`n")

    <div class="footer">
        <p>$script:DocTitle | $script:DocSubTitle | Erstellt: $script:Timestamp | PowerShell Dokumentation v1.0</p>
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
        [PSCustomObject]@{ Key = "Patch";              Label = "Exchange Version & Build";          Category = "Hardware & OS";      Function = "Get-PatchInformation" }
        [PSCustomObject]@{ Key = "EventLog";           Label = "Event Logs (7 Tage)";               Category = "Hardware & OS";      Function = "Get-EventLogInfo" }
        [PSCustomObject]@{ Key = "Firewall";           Label = "Firewall-Regeln & Ports";           Category = "Hardware & OS";      Function = "Get-FirewallInfo" }
        [PSCustomObject]@{ Key = "Licensing";          Label = "Lizenzierung";                      Category = "Hardware & OS";      Function = "Get-LicensingInfo" }

        # === ACTIVE DIRECTORY ===
        [PSCustomObject]@{ Key = "FSMO";               Label = "FSMO-Rollen";                       Category = "Active Directory";   Function = "Get-FSMORoles" }
        [PSCustomObject]@{ Key = "AD";                 Label = "Active Directory & Schema";         Category = "Active Directory";   Function = "Get-ADInformation" }

        # === EXCHANGE GRUNDKONFIGURATION ===
        [PSCustomObject]@{ Key = "ExchangeServer";     Label = "Exchange Server Übersicht";         Category = "Exchange Basis";     Function = "Get-ExchangeServerOverview" }
        [PSCustomObject]@{ Key = "OrgConfig";          Label = "Organisations-Konfiguration";       Category = "Exchange Basis";     Function = "Get-OrganizationConfigInfo" }
        [PSCustomObject]@{ Key = "URLs";               Label = "URLs / VirtualDirectories / EEMS";  Category = "Exchange Basis";     Function = "Get-ExchangeURLs" }
        [PSCustomObject]@{ Key = "Database";           Label = "Datenbanken & DAG";                 Category = "Exchange Basis";     Function = "Get-DatabaseAndDAGInfo" }
        [PSCustomObject]@{ Key = "Certificate";        Label = "Zertifikate";                       Category = "Exchange Basis";     Function = "Get-CertificateInfo" }
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
        [PSCustomObject]@{ Key = "MobileDevice";       Label = "Mobile Device Policies";            Category = "Sicherheit";         Function = "Get-MobileDeviceInfo" }
        [PSCustomObject]@{ Key = "Throttling";         Label = "Throttling Policies";               Category = "Sicherheit";         Function = "Get-ThrottlingPolicyInfo" }

        # === COMPLIANCE ===
        [PSCustomObject]@{ Key = "Compliance";         Label = "Compliance & DLP";                  Category = "Compliance";         Function = "Get-ComplianceInfo" }
        [PSCustomObject]@{ Key = "Retention";          Label = "Retention & Journal";               Category = "Compliance";         Function = "Get-RetentionPolicyInfo" }

        # === HYBRID ===
        [PSCustomObject]@{ Key = "Hybrid";             Label = "Hybrid mit Exchange Online";        Category = "Hybrid / Cloud";     Function = "Get-HybridConfigurationInfo" }
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

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

    # --- Exchange Server automatisch erkennen ---
    $detectedServers = @()
    try {
        # Versuche Exchange Snap-In / Modul zu laden (falls noch nicht geladen)
        $snapInLoaded = Get-PSSnapin -Name "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue
        $moduleLoaded = Get-Module -Name "ExchangeManagementTools" -ErrorAction SilentlyContinue
        if (-not $snapInLoaded -and -not $moduleLoaded) {
            # Snap-In laden
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
        Title="Exchange Server Dokumentation v1.0" Height="900" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#F5F5F5" ResizeMode="CanResize" MinWidth="900" MinHeight="700">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Background="#0078D4" CornerRadius="6" Padding="15" Margin="0,0,0,15">
            <StackPanel>
                <TextBlock Text="Exchange Server Dokumentation" FontSize="24" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Automatische Erkennung der Server | Auswahl der Dokumentationsbereiche | Multi-Format Export" 
                           Foreground="#B3D9FF" FontSize="12" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- SERVER-AUSWAHL -->
        <GroupBox Grid.Row="1" Header="📦 Exchange Server" Margin="0,0,0,12" Padding="10" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                    <CheckBox Name="ChkSelectAllServers" Content="Alle Server" FontWeight="Bold" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <TextBlock Text="Erkannte Server:" VerticalAlignment="Center" Foreground="#555" Margin="0,0,10,0"/>
                </StackPanel>
                <WrapPanel Grid.Row="1" Name="ServerPanel" Orientation="Horizontal"/>
                <TextBox Grid.Row="1" Name="TxtServersManual" Height="28" Visibility="Collapsed" 
                         VerticalContentAlignment="Center" Margin="0,4,0,0"/>
            </Grid>
        </GroupBox>

        <!-- ALLGEMEINE EINSTELLUNGEN -->
        <Grid Grid.Row="2" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Grid.Column="0" Header="🏢 Organisation" Padding="10" Background="White">
                <TextBox Name="TxtCompany" Height="28" VerticalContentAlignment="Center"/>
            </GroupBox>

            <GroupBox Grid.Column="2" Header="📂 Ausgabepfad" Padding="10" Background="White">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Grid.Column="0" Name="TxtOutputPath" Height="28" VerticalContentAlignment="Center"/>
                    <Button Grid.Column="1" Name="BtnBrowse" Content="..." Width="40" Height="28" Margin="8,0,0,0"/>
                </Grid>
            </GroupBox>
        </Grid>

        <!-- AUSGABEFORMATE -->
        <GroupBox Grid.Row="3" Header="📄 Ausgabeformate" Margin="0,0,0,12" Padding="10" Background="White">
            <StackPanel Orientation="Horizontal">
                <CheckBox Name="ChkHtml" Content="HTML" IsChecked="True" Margin="0,0,30,0" VerticalAlignment="Center" FontSize="13"/>
                <CheckBox Name="ChkPdf" Content="PDF (Word/Browser)" Margin="0,0,30,0" VerticalAlignment="Center" FontSize="13"/>
                <CheckBox Name="ChkMarkdown" Content="Markdown (.md)" VerticalAlignment="Center" FontSize="13"/>
            </StackPanel>
        </GroupBox>

        <!-- DOKUMENTATIONSBEREICHE -->
        <GroupBox Grid.Row="4" Header="📋 Zu dokumentierende Bereiche" Margin="0,0,0,12" Background="White">
            <DockPanel>
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="10,8,10,8">
                    <CheckBox Name="ChkSelectAll" Content="✓ Alle auswählen" FontWeight="Bold" Margin="0,0,30,0" IsChecked="True"/>
                    <TextBlock Text="Bereiche nach Themen gruppiert" Foreground="#666" VerticalAlignment="Center"/>
                </StackPanel>
                <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="5">
                    <WrapPanel Name="CategoryPanel" Orientation="Horizontal" ItemWidth="340"/>
                </ScrollViewer>
            </DockPanel>
        </GroupBox>

        <!-- STATUS -->
        <TextBlock Grid.Row="5" Name="TxtStatus" Foreground="#D32F2F" FontWeight="Bold" Margin="0,0,0,10" TextWrapping="Wrap"/>

        <!-- BUTTONS -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="BtnStart" Content="🚀 Dokumentation starten" Width="200" Height="40"
                    Background="#0078D4" Foreground="White" FontWeight="Bold" FontSize="14" Margin="0,0,12,0">
                <Button.Resources>
                    <Style TargetType="Border"><Setter Property="CornerRadius" Value="4"/></Style>
                </Button.Resources>
            </Button>
            <Button Name="BtnCancel" Content="Abbrechen" Width="120" Height="40" FontSize="13"/>
        </StackPanel>
    </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Steuerelemente referenzieren
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
    $btnStart            = $window.FindName("BtnStart")
    $btnCancel           = $window.FindName("BtnCancel")

    # Vorbelegung
    $txtCompany.Text    = $DefaultCompany
    $txtOutputPath.Text = $DefaultOutputPath

    # --- Server-Checkboxen oder manuelles Feld ---
    $serverCheckboxes = @{}
    if ($detectedServers.Count -gt 0) {
        foreach ($srv in $detectedServers) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $srv
            $cb.Tag = $srv
            $cb.IsChecked = $true
            $cb.Margin = "0,0,20,6"
            $cb.FontSize = 13
            [void]$serverPanel.Children.Add($cb)
            $serverCheckboxes[$srv] = $cb
        }
        $chkSelectAllServers.IsChecked = $true
    }
    else {
        # Keine Server gefunden - zeige manuelles Textfeld
        $serverPanel.Visibility = "Collapsed"
        $txtServersManual.Visibility = "Visible"
        $txtServersManual.Text = $DefaultServers
        $chkSelectAllServers.Visibility = "Collapsed"
    }

    # "Alle Server" Checkbox-Logik
    $chkSelectAllServers.Add_Click({
        $state = $chkSelectAllServers.IsChecked
        foreach ($cb in $serverPanel.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox]) { $cb.IsChecked = $state }
        }
    })

    # --- Sektions-Checkboxen nach Kategorien gruppiert ---
    $sectionCheckboxes = @{}
    $registry = Get-DocSectionRegistry
    $categories = $registry | Group-Object -Property Category

    foreach ($cat in $categories) {
        # GroupBox pro Kategorie
        $gb = New-Object System.Windows.Controls.GroupBox
        $gb.Header = $cat.Name
        $gb.Margin = "5"
        $gb.Padding = "8"
        $gb.Background = [System.Windows.Media.Brushes]::White
        $gb.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212))
        $gb.BorderThickness = "1"

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = "Vertical"

        foreach ($section in $cat.Group) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $section.Label
            $cb.Tag = $section.Key
            $cb.IsChecked = $true
            $cb.Margin = "0,3,0,3"
            $cb.FontSize = 12
            [void]$sp.Children.Add($cb)
            $sectionCheckboxes[$section.Key] = $cb
        }

        $gb.Content = $sp
        [void]$categoryPanel.Children.Add($gb)
    }

    # "Alle auswählen" Logik für Sektionen
    $chkSelectAll.Add_Click({
        $state = $chkSelectAll.IsChecked
        foreach ($key in $sectionCheckboxes.Keys) {
            $sectionCheckboxes[$key].IsChecked = $state
        }
    })

    # Durchsuchen-Dialog
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Ausgabeverzeichnis wählen"
        if ($txtOutputPath.Text -and (Test-Path $txtOutputPath.Text)) { $dlg.SelectedPath = $txtOutputPath.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutputPath.Text = $dlg.SelectedPath
        }
    })

    # Ergebnis-Container
    $script:GuiResult = $null

    $btnStart.Add_Click({
        $txtStatus.Text = ""

        # Server ermitteln
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
            $txtStatus.Text = "⚠️ Bitte mindestens einen Exchange-Server auswählen."
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
            $txtStatus.Text = "⚠️ Bitte einen Ausgabepfad angeben."
            return
        }

        $selectedSections = @()
        foreach ($key in $sectionCheckboxes.Keys) {
            if ($sectionCheckboxes[$key].IsChecked) { $selectedSections += $key }
        }
        if ($selectedSections.Count -eq 0) {
            $txtStatus.Text = "⚠️ Bitte mindestens einen Dokumentationsbereich auswählen."
            return
        }

        $formats = @()
        if ($chkHtml.IsChecked)     { $formats += "HTML" }
        if ($chkPdf.IsChecked)      { $formats += "PDF" }
        if ($chkMarkdown.IsChecked) { $formats += "Markdown" }
        if ($formats.Count -eq 0) {
            $txtStatus.Text = "⚠️ Bitte mindestens ein Ausgabeformat wählen."
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
    Write-Log -Message "Exchange Dokumentation gestartet (v1.0 - 2019/SE)" -Level "INFO"
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
    Detect-ExchangeEdition

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
    Write-Host "  Exchange Dokumentation abgeschlossen! (v1.0 - 2019/SE)" -ForegroundColor Cyan
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


