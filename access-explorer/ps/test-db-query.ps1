# test-db-query.ps1 — standalone dev/test helper, NOT part of the shipped extension.
#
# Opens a .accdb directly via Access COM automation, supplies SQL Server credentials
# to any ODBC-linked TableDefs (so RefreshLink authenticates silently instead of
# popping the native modal login dialog, which would hang a headless caller), then
# dumps a few rows from a table/query so real linked data can be inspected.
#
# Usage:
#   powershell -NoProfile -File ps\test-db-query.ps1 -DbPath "C:\...\Something.accdb" -Table "SomeLinkedTable"
#
# Credentials come from db-credentials.local.json (gitignored) next to this script;
# copy db-credentials.local.json.example and fill in real values first.

param(
    [Parameter(Mandatory = $true)] [string]$DbPath,
    [Parameter(Mandatory = $true)] [string]$Table,
    [int]$Top = 20,
    [string]$CredentialsFile,
    [string]$OutFile,
    [string]$SettingsFile,
    [switch]$Visible
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

# $PSScriptRoot is not reliably populated inside param-block default expressions
# on this PowerShell 5.1 setup, so defaults referencing it are resolved here instead.
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

    Write-Host "Opening $DbPath ..."
    Open-DatabaseWithStartupBypass $app $DbPath

    $db = $app.CurrentDb()
    $relinked = Connect-LinkedTables $db $creds
    Write-Host "Relinked $relinked ODBC-linked table(s) with supplied credentials."

    if ($isVisible) { Show-DataObject $app $Table }

    Write-Host "Reading top $Top row(s) from '$Table' ..."
    $rs = Open-LinkedRecordset $db $Table
    try {
        $fieldNames = @()
        foreach ($f in $rs.Fields) { $fieldNames += $f.Name }

        $rows = @()
        $i = 0
        while (-not $rs.EOF -and $i -lt $Top) {
            $row = [ordered]@{}
            foreach ($name in $fieldNames) { $row[$name] = $rs.Fields.Item($name).Value }
            $rows += [pscustomobject]$row
            $rs.MoveNext()
            $i++
        }

        if ($OutFile) {
            $rows | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $OutFile -Encoding utf8
            Write-Host "Wrote $($rows.Count) row(s) to $OutFile"
        } else {
            $rows | Format-Table -AutoSize | Out-String | Write-Host
        }
    } finally {
        $rs.Close()
    }
} catch {
    throw "test-db-query.ps1 failed: $(Describe-Error $_)"
} finally {
    if ($app) {
        if (-not $isVisible) {
            try { $app.CloseCurrentDatabase() } catch { }
            try { $app.Quit() } catch { }
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
