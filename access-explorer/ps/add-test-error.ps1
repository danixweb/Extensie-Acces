# add-test-error.ps1 — standalone dev/test helper, NOT part of the shipped extension.
#
# Inserts a single simulated row into the linked SQL Server error log table
# `dbo_Erori program`, matching exactly what a real TRATARE_ERORI Case Else /
# ScrieEroare call would produce (modul/rutina populated directly), so the
# /vba-corectare-erori skill's full pipeline can be validated end-to-end without
# waiting for a real production crash.
#
# Usage:
#   powershell -NoProfile -File ps\add-test-error.ps1 -DbPath "C:\...\Something.accdb" `
#       -Modul CodBare -Rutina Barcode_128 -Numar 11 -Mesaj "11 Division by zero" `
#       -Context "Netratata CodBare rutina Barcode_128"

param(
    [Parameter(Mandatory = $true)] [string]$DbPath,
    [Parameter(Mandatory = $true)] [string]$Modul,
    [Parameter(Mandatory = $true)] [string]$Rutina,
    [Parameter(Mandatory = $true)] [int]$Numar,
    [Parameter(Mandatory = $true)] [string]$Mesaj,
    [Parameter(Mandatory = $true)] [string]$Context,
    [string]$Utilizator = 'Test Claude',
    [string]$CredentialsFile,
    [string]$SettingsFile,
    [switch]$Visible
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $CredentialsFile) { $CredentialsFile = Join-Path $scriptDir 'db-credentials.local.json' }
if (-not $SettingsFile) { $SettingsFile = Join-Path $scriptDir 'settings.local.json' }

. (Join-Path $scriptDir 'db-link-helper.ps1')
. (Join-Path $scriptDir 'visibility-settings.ps1')

if (-not (Test-Path -LiteralPath $DbPath)) {
    throw "Database not found: $DbPath"
}
$creds = Get-LinkCredentials $CredentialsFile
$isVisible = $Visible -or (Get-VisibleOperationsSetting $SettingsFile)

$app = $null
try {
    $app = New-Object -ComObject Access.Application
    $app.Visible = $isVisible
    # UserControl defaults to False (automation-owned): explicitly True when visible, so
    # Access survives after this script's COM client disconnects, instead of auto-quitting.
    try { $app.UserControl = $isVisible } catch { }

    $app.OpenCurrentDatabase($DbPath, $false)
    $db = $app.CurrentDb()
    Connect-LinkedTables $db $creds | Out-Null

    if ($isVisible) { Show-DataObject $app 'dbo_Erori program' }

    $rs = Open-LinkedRecordset $db 'dbo_Erori program' $dbOpenDynaset
    try {
        $rs.AddNew()
        Set-FieldValue $rs.Fields.Item('ora') (Get-Date)
        Set-FieldValue $rs.Fields.Item('numar') $Numar
        Set-FieldValue $rs.Fields.Item('mesaj') $Mesaj
        Set-FieldValue $rs.Fields.Item('modul') $Modul
        Set-FieldValue $rs.Fields.Item('rutina') $Rutina
        Set-FieldValue $rs.Fields.Item('utilizator') $Utilizator
        Set-FieldValue $rs.Fields.Item('vazut') $false
        Set-FieldValue $rs.Fields.Item('tratata') $false
        Set-FieldValue $rs.Fields.Item('context') $Context
        $rs.Update()
        if ($isVisible) { try { $app.DoCmd.Requery() } catch { } }
    } finally {
        $rs.Close()
    }

    $rsId = Open-LinkedRecordset $db "SELECT MAX(id) AS maxid FROM [dbo_Erori program]" $dbOpenSnapshot
    $newId = $rsId.Fields.Item('maxid').Value
    $rsId.Close()

    Write-Output "Inserted test error row with id=$newId (modul=$Modul, rutina=$Rutina, numar=$Numar)."
} catch {
    throw "add-test-error.ps1 failed: $(Describe-Error $_)"
} finally {
    if ($app) {
        if (-not $isVisible) {
            try { $app.CloseCurrentDatabase() } catch { }
            try { $app.Quit() } catch { }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
