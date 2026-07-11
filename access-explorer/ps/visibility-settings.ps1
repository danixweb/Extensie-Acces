# visibility-settings.ps1 — shared, dot-sourced by BOTH access-bridge.ps1 (shipped engine)
# and the standalone dev scripts in this folder. Kept separate from db-link-helper.ps1
# (which holds SQL Server credential/relink logic, dev-only) since visibility is a
# concern of the shipped bridge too.
#
# Reads the "visibleOperations" opt-in flag from a local, gitignored settings file, so
# Access (and the specific table/module/form/report being acted on) can be shown on
# screen in real time instead of running fully hidden.

function Get-VisibleOperationsSetting([string]$settingsFile) {
    if (-not $settingsFile -or -not (Test-Path -LiteralPath $settingsFile)) { return $false }
    try {
        return [bool]((Get-Content -LiteralPath $settingsFile -Raw | ConvertFrom-Json).visibleOperations)
    } catch {
        return $false
    }
}

# Best-effort: opens a table (or, failing that, a query) datasheet on screen so the
# user watches the actual data while a script reads/writes it. Never fatal — this is
# a visual aid, not something the script's own result depends on.
function Show-DataObject($app, [string]$name) {
    try {
        $app.DoCmd.OpenTable($name)
        return
    } catch { }
    try {
        $app.DoCmd.OpenQuery($name)
    } catch { }
}
