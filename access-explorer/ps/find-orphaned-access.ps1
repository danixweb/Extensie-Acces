# find-orphaned-access.ps1 — lists MSACCESS.EXE processes likely left over from a crashed/killed
# automation session (this extension's own bridge, or a standalone dev script), as opposed to a
# database the user opened normally.
#
# Heuristic: Access launched via COM automation (New-Object -ComObject Access.Application) never
# gets a file path on its command line — OpenCurrentDatabase is called afterward, over COM, not
# passed as an argument. A user double-clicking an .accdb (or Windows reopening one on login)
# always launches MSACCESS.EXE with the file path as a command-line argument. So "no .accdb in the
# command line" is a strong signal the process is an orphaned automation instance, not real user work.
#
# Usage:
#   powershell -NoProfile -File ps\find-orphaned-access.ps1
# Output: JSON array of { pid, commandLine, startTime }, one entry per candidate. Empty array ([])
# when there are none or MSACCESS.EXE isn't running at all.

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$candidates = @()
$procs = Get-CimInstance -ClassName Win32_Process -Filter "Name='MSACCESS.EXE'" -ErrorAction SilentlyContinue
foreach ($p in $procs) {
    $cmdLine = [string]$p.CommandLine
    if ($cmdLine -notmatch '(?i)\.accdb') {
        $startTime = $null
        try { $startTime = $p.CreationDate.ToString('yyyy-MM-dd HH:mm:ss') } catch { }
        $candidates += [pscustomobject]@{
            pid         = [int]$p.ProcessId
            commandLine = $cmdLine
            startTime   = $startTime
        }
    }
}

$json = $candidates | ConvertTo-Json -Depth 3
if ($candidates.Count -eq 1) { $json = "[$json]" }  # ConvertTo-Json drops the array wrapper for a single object
if ($candidates.Count -eq 0) { $json = '[]' }
Write-Output $json
