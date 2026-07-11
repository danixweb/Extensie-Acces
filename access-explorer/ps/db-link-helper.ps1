# db-link-helper.ps1 — shared, dot-sourced helper for the standalone dev/test scripts
# in this folder (test-db-query.ps1, get-untreated-errors.ps1, mark-error-treated.ps1).
# NOT part of the shipped extension/bridge — these scripts talk to Access COM directly
# with SQL Server credentials supplied locally, purely for ad-hoc testing.

$dbOpenDynaset = 2
$dbOpenSnapshot = 4
$dbSeeChanges = 512

function Get-LinkCredentials([string]$credentialsFile) {
    if (-not (Test-Path -LiteralPath $credentialsFile)) {
        throw "Credentials file not found: $credentialsFile (copy db-credentials.local.json.example and fill in real values)"
    }
    $creds = Get-Content -LiteralPath $credentialsFile -Raw | ConvertFrom-Json
    if (-not $creds.uid -or -not $creds.pwd) {
        throw "Credentials file must have non-empty 'uid' and 'pwd' fields: $credentialsFile"
    }
    return $creds
}

# Maps a caught exception to a short message, same categories as access-bridge.ps1's
# Get-ErrorInfo, kept local since these scripts are intentionally independent of the bridge.
function Describe-Error($err) {
    $ex = $err.Exception
    $hr = 0
    try { $hr = $ex.HResult } catch { }
    $num = 0
    if (($hr -band 0xFFFF0000) -eq 0x800A0000) { $num = $hr -band 0xFFFF }
    $hint = switch ($num) {
        { $_ -in 3151, 3059, 3078, 3024 } { 'likely a linked-table authentication failure (wrong UID/PWD or DSN unreachable)'; break }
        { $_ -in 3051, 3033 } { 'permission denied on the linked table'; break }
        { $_ -eq 3622 } { 'SQL Server table has an IDENTITY column - retry needs dbSeeChanges (should be handled automatically)'; break }
        default { $null }
    }
    $msg = ($ex.Message -replace '\s+', ' ').Trim()
    if ($hint) { return "$msg  [error $num - $hint]" }
    return "$msg  [error $num, hresult 0x$($hr.ToString('X8'))]"
}

# Rebuilds an ODBC linked-table Connect string with UID/PWD from the credentials file,
# stripping any existing UID/PWD parts first so this is safe to re-run.
function Add-Credentials([string]$connect, [string]$uid, [string]$pwd) {
    $parts = $connect -split ';' | Where-Object {
        $_ -and $_ -notmatch '^\s*(UID|PWD)\s*='
    }
    $parts += "UID=$uid"
    $parts += "PWD=$pwd"
    return ($parts -join ';') + ';'
}

# Relinks every ODBC-linked TableDef in $db with the supplied credentials, so
# RefreshLink authenticates silently instead of popping the native modal login
# dialog (which would hang a headless caller). Returns the number relinked.
function Connect-LinkedTables($db, $creds) {
    $db.TableDefs.Refresh()
    $relinked = 0
    foreach ($t in $db.TableDefs) {
        if ($t.Connect -and $t.Connect -like 'ODBC;*') {
            $t.Connect = Add-Credentials $t.Connect $creds.uid $creds.pwd
            try {
                $t.RefreshLink()
                $relinked++
            } catch {
                Write-Warning "RefreshLink failed for linked table '$($t.Name)': $(Describe-Error $_)"
            }
        }
    }
    return $relinked
}

# Opens a Recordset, retrying with dbSeeChanges when the source is a SQL Server
# table/query with an IDENTITY column (DAO error 3622).
function Open-LinkedRecordset($db, [string]$source, [int]$type = $dbOpenDynaset) {
    try {
        return $db.OpenRecordset($source, $type)
    } catch {
        if ((Describe-Error $_) -match 'error 3622') {
            return $db.OpenRecordset($source, $type, $dbSeeChanges)
        }
        throw
    }
}

# PowerShell's COM property-set binder for a DAO Field's "Value" property gets stuck
# on the .NET type of an earlier value assigned to any property named "Value" in the
# process, then throws InvalidCastException for later, differently-typed fields (e.g.
# a Date field fails with "Unable to cast ... DateTime to ... String"). Routing the
# assignment through raw reflection instead of PowerShell's binder sidesteps that cache.
function Set-FieldValue($field, $value) {
    $field.GetType().InvokeMember('Value', [System.Reflection.BindingFlags]::SetProperty, $null, $field, @($value)) | Out-Null
}
