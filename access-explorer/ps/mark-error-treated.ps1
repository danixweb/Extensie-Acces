# mark-error-treated.ps1 — standalone dev/test helper, NOT part of the shipped extension.
#
# Marks a single row in the linked SQL Server error log table `dbo_Erori program` as
# treated (tratata=True), by id. Used by the /vba-corectare-erori skill after a
# proposed fix has been confirmed and applied to the VBA source.
#
# Usage:
#   powershell -NoProfile -File ps\mark-error-treated.ps1 -DbPath "C:\...\Something.accdb" -Id 7186

param(
    [Parameter(Mandatory = $true)] [string]$DbPath,
    [Parameter(Mandatory = $true)] [int]$Id,
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

    Open-DatabaseWithStartupBypass $app $DbPath
    $db = $app.CurrentDb()
    Connect-LinkedTables $db $creds | Out-Null

    if ($isVisible) { Show-DataObject $app 'dbo_Erori program' }

    $sql = "SELECT * FROM [dbo_Erori program] WHERE id=$Id"
    $rs = Open-LinkedRecordset $db $sql $dbOpenDynaset
    try {
        if ($rs.EOF) {
            throw "No row with id=$Id found in [dbo_Erori program]."
        }
        $rs.Edit()
        Set-FieldValue $rs.Fields.Item('tratata') $true
        $rs.Update()
        if ($isVisible) { try { $app.DoCmd.Requery() } catch { } }
        Write-Output "Marked id=$Id as tratata=True."
    } finally {
        $rs.Close()
    }
} catch {
    throw "mark-error-treated.ps1 failed: $(Describe-Error $_)"
} finally {
    if ($app) {
        if (-not $isVisible) {
            try { $app.CloseCurrentDatabase() } catch { }
            try { $app.Quit() } catch { }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
