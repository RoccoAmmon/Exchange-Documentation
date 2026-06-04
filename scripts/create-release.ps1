<#
.SYNOPSIS
    Erstellt ein Git-Tag und pusht es zum Remote-Repository, damit ein Release-Workflow ausgelöst wird.
.NOTES
    Autor: Rocco Ammon
    Version: 1.0
    Erstellt: 2026-06-04
#>

param(
    [string]$Version = 'v1.0',
    [string]$Remote = 'origin'
)

try {
    $ErrorActionPreference = 'Stop'

    Write-Host "Erstelle Tag $Version und pushe zu $Remote..."

    # Prüfen, ob Git verfügbar ist
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw 'git ist nicht installiert oder nicht im PATH.'
    }

    # Sauberen Arbeitsbaum sicherstellen
    $status = git status --porcelain
    if ($status) {
        throw 'Arbeitsbaum nicht sauber. Bitte committe oder stash deine Änderungen.'
    }

    # Tag anlegen (falls noch nicht vorhanden)
    $existing = git tag --list $Version
    if (-not $existing) {
        git tag -a $Version -m "Release $Version"
        Write-Host "Tag $Version erstellt."
    }
    else {
        Write-Host "Tag $Version existiert bereits."
    }

    # Push Tag
    git push $Remote $Version
    Write-Host "Tag $Version zu $Remote gepusht. GitHub Release-Workflow wird ausgelöst (wenn konfiguriert)."
}
catch {
    Write-Host "FEHLER: $_" -ForegroundColor Red
    exit 1
}
