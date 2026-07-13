# get-untreated-errors.ps1 — standalone dev/test helper, NOT part of the shipped extension.
#
# Reads up to -Limit untreated rows (tratata=False) from the linked SQL Server error
# log table `dbo_Erori program`, ordered by id, and prints them as JSON — input for
# the /vba-corectare-erori skill.
#
# Usage:
#   powershell -NoProfile -File ps\get-untreated-errors.ps1 -DbPath "C:\...\Something.accdb" -Limit 5

param(
    [Parameter(Mandatory = $true)] [string]$DbPath,
    [Parameter(Mandatory = $true)] [int]$Limit,
    [string]$CredentialsFile,
    [string]$OutFile,
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

    $sql = "SELECT TOP $Limit * FROM [dbo_Erori program] WHERE tratata=False ORDER BY id"
    $rs = Open-LinkedRecordset $db $sql $dbOpenSnapshot
    try {
        $fieldNames = @()
        foreach ($f in $rs.Fields) { $fieldNames += $f.Name }

        $rows = @()
        while (-not $rs.EOF) {
            $row = [ordered]@{}
            foreach ($name in $fieldNames) { $row[$name] = $rs.Fields.Item($name).Value }
            $rows += [pscustomobject]$row
            $rs.MoveNext()
        }

        $json = $rows | ConvertTo-Json -Depth 5
        if ($rows.Count -eq 1) { $json = "[$json]" }  # ConvertTo-Json drops the array wrapper for a single object
        if ($rows.Count -eq 0) { $json = '[]' }       # ...and returns $null (empty output) for an empty array
        if ($OutFile) {
            $json | Out-File -LiteralPath $OutFile -Encoding utf8
        } else {
            Write-Output $json
        }
    } finally {
        $rs.Close()
    }
} catch {
    throw "get-untreated-errors.ps1 failed: $(Describe-Error $_)"
} finally {
    if ($app) {
        if (-not $isVisible) {
            try { $app.CloseCurrentDatabase() } catch { }
            try { $app.Quit() } catch { }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
