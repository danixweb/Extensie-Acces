# access-bridge.ps1 — COM bridge between the VS Code extension and Microsoft Access.
#
# Protocol: JSON lines over stdin/stdout.
#   Request : {"id":1,"op":"open","args":{"path":"C:\\db.accdb"}}
#   Response: {"id":1,"ok":true,"data":{...}}  or  {"id":1,"ok":false,"error":{"code":"...","message":"...","number":3045}}
# stdout carries ONLY protocol lines; diagnostics go to stderr.
# Large payloads (module code, SQL) are exchanged through temp files named in args.file —
# never inline in JSON — so framing can't break on huge or exotic module bodies.

# powershell.exe defaults stdio to the OEM code page, which mangles diacritics (ăâîșț).
# Do NOT assign [Console]::InputEncoding here — on redirected stdin the setter reopens
# the handle and silently drops already-buffered input. Use explicit UTF-8 streams instead.
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$stdin  = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $Utf8NoBom)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $Utf8NoBom)
$stdout.AutoFlush = $true

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

# Access object type constants (AcObjectType)
$acTable = 0; $acQuery = 1; $acForm = 2; $acReport = 3; $acMacro = 4; $acModule = 5
$acQuitSaveNone = 2
$acCmdCompileAndSaveAllModules = 126
$acDesignView = 1

. (Join-Path $PSScriptRoot 'visibility-settings.ps1')

Add-Type -Name Win32 -Namespace Bridge -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(System.IntPtr hWnd);
[DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, System.UIntPtr dwExtraInfo);
'@

$VK_SHIFT = 0x10
$KEYEVENTF_KEYUP = 0x2

$script:app = $null
$script:accessPid = 0
$script:accessHwnd = [System.IntPtr]::Zero
$script:dbPath = $null
$script:visibleOperations = $false
$script:bypassStartup = $false
$script:bypassKeyEnsuredFor = $null

# Opens the database SHARED in $script:app, bypassing the Startup form / AutoExec when
# $script:bypassStartup is set: ensures AllowBypassKey=True via DAO first (the Shift trick
# is silently ignored otherwise), simulates a held Shift during the open, then closes any
# form that still managed to load (e.g. a physical keypress broke the simulated Shift).
# Every OpenCurrentDatabase in this bridge must go through here — a Startup form waiting
# for login, or a MsgBox in a Form_Load, is a modal window that blocks all later COM ops.
# Returns $true when the Startup form ran anyway (reported to the caller as startupRan).
function Open-DbShared([string]$path) {
    if (-not $script:bypassStartup) {
        $script:app.OpenCurrentDatabase($path, $false)
        return $false
    }
    # Only probe/set AllowBypassKey once per path per bridge session (e.g. not again on the
    # reopen-after-compact) — opening a SEPARATE DAO.Database handle on a file Access.Application
    # already has open elsewhere in this same run was observed to leave the Jet/ACE engine with a
    # residual lock that made a later CompactRepair fail with "already opened by user ... on
    # machine ...", even after the extra handle was explicitly Close()'d and released. Reusing
    # $script:app.DBEngine (the SAME engine instance Access itself uses, rather than spinning up
    # a second one via `New-Object -ComObject DAO.DBEngine`) avoids a second engine entirely.
    if ($script:bypassKeyEnsuredFor -ne $path) {
        try {
            $dbDao = $script:app.DBEngine.OpenDatabase($path)
            try {
                $prop = $dbDao.Properties.Item('AllowBypassKey')
                if (-not $prop.Value) { $prop.Value = $true }
            } catch {
                # dbBoolean = 1
                $dbDao.Properties.Append($dbDao.CreateProperty('AllowBypassKey', 1, $true))
            }
            $dbDao.Close()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbDao)
            $script:bypassKeyEnsuredFor = $path
        } catch {
            [Console]::Error.WriteLine("bridge warning: could not ensure AllowBypassKey: $($_.Exception.Message)")
        }
    }
    [Bridge.Win32]::keybd_event($VK_SHIFT, 0, 0, [System.UIntPtr]::Zero)
    Start-Sleep -Milliseconds 100
    try {
        $script:app.OpenCurrentDatabase($path, $false)
    } finally {
        # Always release Shift, even on failure — otherwise it stays "stuck" system-wide.
        [Bridge.Win32]::keybd_event($VK_SHIFT, 0, $KEYEVENTF_KEYUP, [System.UIntPtr]::Zero)
    }
    $startupRan = $false
    try {
        $startupRan = ($script:app.Forms.Count -gt 0)
        $guard = 0
        while ($script:app.Forms.Count -gt 0 -and $guard -lt 20) {
            # acForm = 2, acSaveNo = 2
            $script:app.DoCmd.Close(2, $script:app.Forms.Item(0).Name, 2)
            $guard++
        }
    } catch {
        [Console]::Error.WriteLine("bridge warning: could not close startup form(s): $($_.Exception.Message)")
    }
    return $startupRan
}

function Write-Response([hashtable]$obj) {
    $json = ($obj | ConvertTo-Json -Compress -Depth 8)
    $stdout.WriteLine($json)
}

# SaveAsText emits ANSI (current code page) for modules and UTF-16 for some macro/form
# exports. Detect by BOM, fall back to the system ANSI code page.
function Read-ExportedText([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @{ text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2); enc = 'utf16' }
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @{ text = $Utf8NoBom.GetString($bytes, 3, $bytes.Length - 3); enc = 'utf8' }
    }
    return @{ text = [System.Text.Encoding]::Default.GetString($bytes); enc = 'ansi' }
}

function Write-ImportText([string]$path, [string]$text, [string]$enc) {
    switch ($enc) {
        'utf16' { [System.IO.File]::WriteAllText($path, $text, [System.Text.Encoding]::Unicode) }
        default { [System.IO.File]::WriteAllText($path, $text, [System.Text.Encoding]::Default) }
    }
}

# Maps a caught exception to a stable error code the extension can localize.
function Get-ErrorInfo($err) {
    $ex = $err.Exception
    # PowerShell wraps COM failures (e.g. MethodInvocationException around a COMException
    # from SaveAsText/LoadFromText) — the outer HResult is then a generic CLR code and the
    # DAO/Access number lives on an inner exception. Walk the chain and prefer the first
    # FACILITY_CONTROL (0x800A....) or class-not-registered HResult found.
    $hr = 0
    $cur = $ex
    while ($null -ne $cur) {
        $h = 0
        try { $h = $cur.HResult } catch { }
        if ($h -eq 0x80040154 -or $h -eq -2147221164 -or (($h -band 0xFFFF0000) -eq 0x800A0000)) { $hr = $h; break }
        if ($hr -eq 0 -and $h -ne 0) { $hr = $h }
        $cur = $cur.InnerException
    }
    $num = 0
    # DAO/Access automation errors use FACILITY_CONTROL (0xA); low word = error number.
    if (($hr -band 0xFFFF0000) -eq 0x800A0000) { $num = $hr -band 0xFFFF }
    $code = 'COM_ERROR'
    if ($hr -eq 0x80040154 -or $hr -eq -2147221164) { $code = 'ACCESS_NOT_INSTALLED' }
    elseif ($num -in 3045, 3050, 3704, 3734, 3009, 3211, 3260) { $code = 'DB_LOCKED' }
    elseif ($num -in 3051, 3033) { $code = 'NO_PERMISSION' }
    elseif ($num -in 3024, 3044, 3055) { $code = 'DB_NOT_FOUND' }
    elseif ($num -in 2544, 3265, 7874, 2103, 29068, 32584) { $code = 'OBJECT_NOT_FOUND' }
    elseif ($num -in 3129, 3141, 3075, 3067, 3131, 3134, 3144, 3061) { $code = 'SQL_SYNTAX' }
    elseif ($num -in 2501, 29054) { $code = 'MACRO_SECURITY' }
    return @{
        code    = $code
        number  = $num
        hresult = $hr
        message = ($ex.Message -replace '\s+', ' ').Trim()
    }
}

function Assert-Open {
    if ($null -eq $script:app -or $null -eq $script:dbPath) {
        throw 'No database is open in this bridge (op requires a prior "open").'
    }
}

# Releases a short-lived per-request COM RCW (DAO collection/item, VBComponent, CodeModule).
# Never call this on $script:app itself or anything reused across requests — Close-App is
# the only place that releases the application object, on the way out.
function Clear-ComObject([object]$obj) {
    if ($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

function Get-Db { return $script:app.CurrentDb() }

function Op-Open([hashtable]$args_) {
    $path = [string]$args_.path
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ _error = @{ code = 'DB_NOT_FOUND'; number = 0; hresult = 0; message = "File not found: $path" } }
    }
    $script:app = New-Object -ComObject Access.Application
    # Explicit arg (from the extension's own setting) wins; otherwise fall back to the
    # local settings file, so direct/manual protocol use (e.g. a skill driving the bridge
    # by hand) still honors the same shared "visible operations" preference.
    $script:visibleOperations = if ($args_.ContainsKey('visibleOperations')) {
        [bool]$args_.visibleOperations
    } else {
        Get-VisibleOperationsSetting (Join-Path $PSScriptRoot 'settings.local.json')
    }
    $script:app.Visible = $script:visibleOperations
    # UserControl defaults to False (automation-owned): explicitly True when visible, so
    # Access survives even if the bridge process dies unexpectedly instead of auto-quitting
    # (normal shutdown still goes through Close-App's explicit Quit() either way).
    try { $script:app.UserControl = $script:visibleOperations } catch { }
    # Resolve the real MSACCESS.EXE PID (and main window handle, reused later to bring the
    # window to the front for VBA project unlock) so the extension can kill it if the bridge hangs.
    # hWndAccessApp surfaces as a method through PowerShell COM late binding.
    try {
        $script:accessHwnd = [System.IntPtr][int]$script:app.hWndAccessApp()
        $procId = [uint32]0
        [Bridge.Win32]::GetWindowThreadProcessId($script:accessHwnd, [ref]$procId) | Out-Null
        $script:accessPid = $procId
    } catch {
        [Console]::Error.WriteLine("bridge warning: could not resolve Access PID: $($_.Exception.Message)")
        $script:accessPid = 0
        $script:accessHwnd = [System.IntPtr]::Zero
    }
    # Bypass Startup/AutoExec — for a code-reading/editing tool, running the app's Startup
    # form or AutoExec macro never helps and only risks a blocking modal dialog (a Startup
    # form waiting for login, a MsgBox in Form_Load) or an app that quits itself. The flag
    # is remembered so every later reopen (compact) uses the same bypass.
    $script:bypassStartup = $args_.ContainsKey('bypassStartup') -and $args_.bypassStartup
    # Shared mode: coexists with the Access UI having the file open normally.
    $startupRan = Open-DbShared $path
    $script:dbPath = $path
    # A password-locked VBA project (Tools > VBAProject Properties > Protection) is a different
    # condition from the Trust Center's "Trust access to the VBA project object model" setting —
    # reading .Protection never itself triggers the password prompt, it just reports lock state.
    $vbaProtected = $false
    try { $vbaProtected = ($script:app.VBE.ActiveVBProject.Protection -eq 1) } catch { $vbaProtected = $true }
    return @{ accessPid = $script:accessPid; path = $path; vbaProtected = $vbaProtected; startupRan = $startupRan }
}

function Op-List {
    Assert-Open
    $db = Get-Db
    $tables = @()
    $hasLinkedTables = $false
    $tableDefs = $db.TableDefs
    $tableDefs.Refresh()
    foreach ($t in $tableDefs) {
        $n = $t.Name
        $isOdbc = [bool]($t.Connect -and $t.Connect -like 'ODBC;*')
        Clear-ComObject $t
        if ($isOdbc) { $hasLinkedTables = $true }
        if ($n -like 'MSys*' -or $n -like '~*' -or $n -like 'f_*ADO*') { continue }
        $tables += $n
    }
    Clear-ComObject $tableDefs

    $queries = @()
    $queryDefs = $db.QueryDefs
    $queryDefs.Refresh()
    foreach ($q in $queryDefs) {
        $qn = $q.Name
        Clear-ComObject $q
        if ($qn -like '~*') { continue }
        $queries += $qn
    }
    Clear-ComObject $queryDefs

    $proj = $script:app.CurrentProject
    $forms = @();   foreach ($o in $proj.AllForms)   { $forms   += $o.Name; Clear-ComObject $o }
    $reports = @(); foreach ($o in $proj.AllReports) { $reports += $o.Name; Clear-ComObject $o }
    $macros = @();  foreach ($o in $proj.AllMacros)  { $macros  += $o.Name; Clear-ComObject $o }
    $modules = @(); foreach ($o in $proj.AllModules) { $modules += $o.Name; Clear-ComObject $o }
    Clear-ComObject $proj
    Clear-ComObject $db

    return @{
        tables  = @($tables  | Sort-Object)
        queries = @($queries | Sort-Object)
        forms   = @($forms   | Sort-Object)
        reports = @($reports | Sort-Object)
        macros  = @($macros  | Sort-Object)
        modules = @($modules | Sort-Object)
        hasLinkedTables = $hasLinkedTables
    }
}

# Rebuilds an ODBC linked-table Connect string with UID/PWD, stripping any existing UID/PWD
# parts first so re-relinking (e.g. after the extension prompts for new credentials) is safe
# to repeat. Note this persists into the .accdb's own TableDef metadata — the standard Access
# "saved link" mechanism — not just into this bridge session.
function Add-LinkCredentials([string]$connect, [string]$uid, [string]$pwd) {
    $parts = $connect -split ';' | Where-Object { $_ -and $_ -notmatch '^\s*(UID|PWD)\s*=' }
    $parts += "UID=$uid"
    $parts += "PWD=$pwd"
    return ($parts -join ';') + ';'
}

# Applies the given credentials to every ODBC-linked TableDef and calls RefreshLink, so opening
# a linked table's data authenticates silently instead of popping the native modal login dialog
# (invisible to COM automation, which would hang the bridge).
function Op-RelinkCredentials([hashtable]$args_) {
    Assert-Open
    $uid = [string]$args_.uid
    $pwd = [string]$args_.pwd
    $db = Get-Db
    $db.TableDefs.Refresh()
    $relinked = 0
    $failedNames = @()
    foreach ($t in $db.TableDefs) {
        if ($t.Connect -and $t.Connect -like 'ODBC;*') {
            $t.Connect = Add-LinkCredentials $t.Connect $uid $pwd
            try {
                $t.RefreshLink()
                $relinked++
            } catch {
                $failedNames += $t.Name
            }
        }
        Clear-ComObject $t
    }
    Clear-ComObject $db
    if ($relinked -eq 0 -and $failedNames.Count -gt 0) {
        return @{ _error = @{
            code    = 'LINKED_AUTH_FAILED'
            number  = 0
            hresult = 0
            message = "Failed to authenticate linked table(s) with the given credentials: $($failedNames -join ', ')"
        } }
    }
    return @{ relinked = $relinked; failed = @($failedNames) }
}

# Export an object with SaveAsText into args.file (re-encoded UTF-8 for the extension).
function Export-ObjectText([int]$objType, [string]$name, [string]$outFile) {
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        $script:app.SaveAsText($objType, $name, $tmp)
        $r = Read-ExportedText $tmp
        [System.IO.File]::WriteAllText($outFile, $r.text, $Utf8NoBom)
        return $r
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

# Import UTF-8 text from args.file into an object with LoadFromText (re-encoded for Access).
function Import-ObjectText([int]$objType, [string]$name, [string]$inFile, [string]$enc) {
    $text = [System.IO.File]::ReadAllText($inFile, $Utf8NoBom)
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        Write-ImportText $tmp $text $enc
        $script:app.LoadFromText($objType, $name, $tmp)
    } finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

# VBE access ("Trust access to the VBA project object model") is optional: SaveAsText
# covers standard modules everywhere, but on some Access versions (e.g. 2013) it
# exports CLASS modules as empty files — those need the VBE CodeModule path.
function Get-VbComponent([string]$name) {
    return $script:app.VBE.ActiveVBProject.VBComponents.Item($name)
}

# Best-effort, non-fatal: brings the actual code/object being read or written on screen
# when visibleOperations is on, so the user watches the action happen in real time.
function Show-CodePane([string]$name) {
    if (-not $script:visibleOperations) { return }
    try {
        $script:app.VBE.MainWindow.Visible = $true
        $vbc = Get-VbComponent $name
        $cm = $vbc.CodeModule
        $cm.CodePane.Show()
        Clear-ComObject $cm
        Clear-ComObject $vbc
    } catch { }
}

function Show-DesignObject([string]$kind, [string]$name) {
    if (-not $script:visibleOperations) { return }
    try {
        if ($kind -eq 'form') { $script:app.DoCmd.OpenForm($name, $acDesignView) }
        else { $script:app.DoCmd.OpenReport($name, $acDesignView) }
    } catch { }
}

function Op-GetModule([hashtable]$args_) {
    Assert-Open
    $name = [string]$args_.name
    Show-CodePane $name
    $r = Export-ObjectText $acModule $name $args_.file
    if ($r.text.Length -gt 0) {
        # Newer Access versions include the VERSION/Attribute header; detect class-ness
        # from it, refining via VBE when available.
        $isClass = $r.text -match '(?m)^\s*VERSION \d+\.\d+ CLASS\s*$'
        try {
            $vbc = Get-VbComponent $name
            $isClass = ($vbc.Type -eq 2)
            Clear-ComObject $vbc
        } catch { }
        return @{ isClass = [bool]$isClass; viaVbe = $false }
    }
    # Empty export: either a genuinely empty module or a class module this Access
    # version cannot SaveAsText. Read through the VBE instead.
    try {
        $c = Get-VbComponent $name
        $cm = $c.CodeModule
        $text = ''
        if ($cm.CountOfLines -gt 0) { $text = $cm.Lines(1, $cm.CountOfLines) }
        [System.IO.File]::WriteAllText($args_.file, $text, $Utf8NoBom)
        $isClassVia = ($c.Type -eq 2)
        Clear-ComObject $cm
        Clear-ComObject $c
        return @{ isClass = $isClassVia; viaVbe = $true }
    } catch {
        return @{ _error = @{
            code    = 'VBE_TRUST_REQUIRED'
            number  = 0
            hresult = 0
            message = "SaveAsText returned no content for module '$name' and the VBA project object model is not trusted."
        } }
    }
}

function Op-SaveModule([hashtable]$args_) {
    Assert-Open
    $name = [string]$args_.name
    $viaVbe = $false
    if ($args_.ContainsKey('viaVbe') -and $args_.viaVbe) { $viaVbe = $true }
    if (-not $viaVbe) {
        Import-ObjectText $acModule $name $args_.file 'ansi'
        # LoadFromText only updates the in-memory VBA project; without an explicit,
        # synchronous Save here the change only reaches disk via the async post-save
        # compile check (fsProvider.scheduleCompileCheck), and is lost if the database
        # is closed (Quit acQuitSaveNone) before that check finishes.
        $script:app.DoCmd.Save($acModule, $name)
        Show-CodePane $name
        return @{ saved = $true }
    }
    # Modules read through the VBE must be written back through it too —
    # LoadFromText would replace the class module with a broken standard module.
    $text = [System.IO.File]::ReadAllText($args_.file, $Utf8NoBom)
    $vbc = Get-VbComponent $name
    $cm = $vbc.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    if ($text.Length -gt 0) { $cm.AddFromString($text) }
    $script:app.DoCmd.Save($acModule, $name)
    Show-CodePane $name
    Clear-ComObject $cm
    Clear-ComObject $vbc
    return @{ saved = $true }
}

function Op-GetMacro([hashtable]$args_) {
    Assert-Open
    $r = Export-ObjectText $acMacro $args_.name $args_.file
    return @{ enc = $r.enc }
}

function Op-SaveMacro([hashtable]$args_) {
    Assert-Open
    $enc = if ($args_.ContainsKey('enc') -and $args_.enc) { [string]$args_.enc } else { 'utf16' }
    Import-ObjectText $acMacro $args_.name $args_.file $enc
    return @{ saved = $true }
}

function Op-GetQuerySql([hashtable]$args_) {
    Assert-Open
    $db = Get-Db
    $queryDefs = $db.QueryDefs
    $queryDefs.Refresh()
    $qdef = $queryDefs.Item($args_.name)
    $sql = $qdef.SQL
    Clear-ComObject $qdef
    Clear-ComObject $queryDefs
    Clear-ComObject $db
    [System.IO.File]::WriteAllText($args_.file, $sql, $Utf8NoBom)
    return @{ }
}

function Op-SaveQuerySql([hashtable]$args_) {
    Assert-Open
    $sql = [System.IO.File]::ReadAllText($args_.file, $Utf8NoBom)
    $db = Get-Db
    $queryDefs = $db.QueryDefs
    $queryDefs.Refresh()
    $qdef = $queryDefs.Item($args_.name)
    try {
        # DAO validates the SQL at assignment and throws (e.g. 3129) on bad syntax.
        $qdef.SQL = $sql
    } finally {
        Clear-ComObject $qdef
        Clear-ComObject $queryDefs
        Clear-ComObject $db
    }
    return @{ saved = $true }
}

function Op-GetTableDef([hashtable]$args_) {
    Assert-Open
    $db = Get-Db
    $t = $db.TableDefs($args_.name)
    $daoTypes = @{ 1='Boolean'; 2='Byte'; 3='Integer'; 4='Long'; 5='Currency'; 6='Single'; 7='Double';
                   8='Date/Time'; 9='Binary'; 10='Text'; 11='OLE Object'; 12='Memo/Long Text';
                   15='GUID'; 16='BigInt'; 18='Char'; 20='Decimal'; 23='VarChar'; 101='Attachment';
                   102='Complex Byte'; 103='Complex Integer'; 104='Complex Long'; 106='Complex Double' }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("Table: $($t.Name)")
    # RecordCount on an ODBC-linked table forces DAO to run a live query against the
    # remote server — potentially slow/hung on a bad connection, independent of any COM
    # cleanup or timeout tuning. It's purely informational here, so skip it for links.
    $isOdbc = [bool]($t.Connect -and $t.Connect -like 'ODBC;*')
    $recordCountText = if ($isOdbc) { '(linked table — count not queried)' } else { [string]$t.RecordCount }
    [void]$sb.AppendLine("Records: $recordCountText")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Fields:')
    $fields = $t.Fields
    foreach ($f in $fields) {
        $tn = if ($daoTypes.ContainsKey([int]$f.Type)) { $daoTypes[[int]$f.Type] } else { "Type $($f.Type)" }
        $req = if ($f.Required) { ' REQUIRED' } else { '' }
        [void]$sb.AppendLine(("  {0,-32} {1}({2}){3}" -f $f.Name, $tn, $f.Size, $req))
        Clear-ComObject $f
    }
    Clear-ComObject $fields
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Indexes:')
    $indexes = $t.Indexes
    foreach ($ix in $indexes) {
        $flags = @()
        if ($ix.Primary) { $flags += 'PRIMARY' }
        if ($ix.Unique)  { $flags += 'UNIQUE' }
        $cols = @()
        $ixFields = $ix.Fields
        foreach ($f in $ixFields) { $cols += $f.Name; Clear-ComObject $f }
        Clear-ComObject $ixFields
        [void]$sb.AppendLine(("  {0,-32} ({1}) {2}" -f $ix.Name, ($cols -join ', '), ($flags -join ' ')))
        Clear-ComObject $ix
    }
    Clear-ComObject $indexes
    Clear-ComObject $t
    Clear-ComObject $db
    [System.IO.File]::WriteAllText($args_.file, $sb.ToString(), $Utf8NoBom)
    return @{ }
}

function Op-GetFormDef([hashtable]$args_)   { Assert-Open; Show-DesignObject 'form' $args_.name; $r = Export-ObjectText $acForm   $args_.name $args_.file; return @{ enc = $r.enc } }
function Op-GetReportDef([hashtable]$args_) { Assert-Open; Show-DesignObject 'report' $args_.name; $r = Export-ObjectText $acReport $args_.name $args_.file; return @{ enc = $r.enc } }

function Op-SaveFormDef([hashtable]$args_) {
    Assert-Open
    $enc = if ($args_.ContainsKey('enc') -and $args_.enc) { [string]$args_.enc } else { 'ansi' }
    Import-ObjectText $acForm $args_.name $args_.file $enc
    Show-DesignObject 'form' $args_.name
    return @{ saved = $true }
}

function Op-SaveReportDef([hashtable]$args_) {
    Assert-Open
    $enc = if ($args_.ContainsKey('enc') -and $args_.enc) { [string]$args_.enc } else { 'ansi' }
    Import-ObjectText $acReport $args_.name $args_.file $enc
    Show-DesignObject 'report' $args_.name
    return @{ saved = $true }
}

# Shows Access + the VBA editor so the user can type/paste the project password into the real,
# native Access dialog, then polls .Protection until it clears (or times out). There is no public
# COM API to supply the password programmatically, so this deliberately leaves the actual entry to
# the human — far more reliable across Access versions/locales than simulating keystrokes.
function Op-UnlockVba([hashtable]$args_) {
    Assert-Open
    $timeoutSeconds = if ($args_.ContainsKey('timeoutSeconds')) { [int]$args_.timeoutSeconds } else { 120 }
    try {
        $script:app.Visible = $true
        if ($script:accessHwnd -ne [System.IntPtr]::Zero) {
            [Bridge.Win32]::SetForegroundWindow($script:accessHwnd) | Out-Null
        }
        try { $script:app.VBE.MainWindow.Visible = $true } catch { }

        # .Protection only reports whether the project is CONFIGURED as locked (a static
        # attribute) — it never flips after the user types the password. The real signal is
        # whether VBComponents is actually reachable yet, the same call Get-VbComponent needs.
        $deadline = (Get-Date).AddSeconds($timeoutSeconds)
        while ((Get-Date) -lt $deadline) {
            try {
                $null = $script:app.VBE.ActiveVBProject.VBComponents.Count
                return @{ unlocked = $true }
            } catch { }
            Start-Sleep -Milliseconds 500
        }
        return @{ unlocked = $false }
    } finally {
        # Restore to the session's persistent visibility preference, not unconditionally
        # hidden — otherwise this would clobber an active "visible operations" setting.
        try { $script:app.Visible = $script:visibleOperations } catch { }
    }
}

function Op-Compile([hashtable]$args_) {
    Assert-Open
    # On a password-locked VBA project (or without VBA object-model trust) the
    # CompileAndSaveAllModules command silently no-ops and reports success even when a
    # module has a syntax error — false confidence. Detect via the same VBComponents
    # reachability probe unlockVba polls, and say the check was skipped instead.
    $vbeReachable = $true
    try { $null = $script:app.VBE.ActiveVBProject.VBComponents.Count } catch { $vbeReachable = $false }
    if (-not $vbeReachable) {
        return @{ compiled = $true; skipped = $true; message = 'VBA project is password-locked or the VBA object model is not trusted; compile check unavailable.' }
    }
    try {
        $script:app.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
        return @{ compiled = $true }
    } catch {
        $info = Get-ErrorInfo $_
        if ($info.number -eq 2046) {
            # "Command not available" — needs an open module context; open one and retry.
            try {
                $script:app.DoCmd.OpenModule($args_.name)
                try {
                    $script:app.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
                    return @{ compiled = $true }
                } catch {
                    $info2 = Get-ErrorInfo $_
                    return @{ compiled = $false; message = $info2.message; number = $info2.number }
                } finally {
                    try { $script:app.DoCmd.Close($acModule, $args_.name, 2) } catch { }
                }
            } catch {
                return @{ compiled = $false; message = $info.message; number = $info.number }
            }
        }
        return @{ compiled = $false; message = $info.message; number = $info.number }
    }
}

function Op-Backup([hashtable]$args_) {
    Assert-Open
    $target = [string]$args_.target
    $dir = Split-Path -Parent $target
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    # The DB is open with read sharing, so copying the live file is allowed.
    Copy-Item -LiteralPath $script:dbPath -Destination $target -Force
    return @{ target = $target }
}

# Compacts and repairs the current database in place: closes it, runs CompactRepair into a temp
# file next to it, then swaps the temp file over the original and reopens it. CompactRepair needs
# the source file free of any exclusive/shared hold from THIS automation instance (hence the
# CloseCurrentDatabase first) and can fail outright if another user/process has it open elsewhere
# (DB_LOCKED) — any failure here always leaves the original file untouched and reopens it, never
# destructive.
function Op-Compact([hashtable]$args_) {
    Assert-Open
    $sourcePath = $script:dbPath
    $dir = Split-Path -Parent $sourcePath
    $ext = [System.IO.Path]::GetExtension($sourcePath)
    $tempPath = Join-Path $dir ([System.IO.Path]::GetFileNameWithoutExtension($sourcePath) + '_compact_' + [System.Guid]::NewGuid().ToString('N').Substring(0, 8) + $ext)

    $script:app.CloseCurrentDatabase()

    # Every earlier op that called Get-Db (CurrentDb()) leaves behind a live RCW referencing
    # the DAO Database object — none of those call sites release it, so across a long-running
    # session dozens can accumulate. .NET never collects them on its own schedule, so the
    # underlying Jet/ACE engine sees the database as still "open" via those dangling references
    # even after CloseCurrentDatabase() — CompactRepair then fails with "already opened by user
    # ... on machine ..." even though no real client holds it. Force collection before compacting.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    # CloseCurrentDatabase() returns before the OS-level file lock is always fully released —
    # calling CompactRepair immediately can race and fail with "already opened by user ...
    # on machine ..." (this same instance's own, not-yet-cleared lock; the .laccdb itself can
    # already be gone when this happens, e.g. under OneDrive sync briefly holding a read
    # handle after heavy write activity — the busier the preceding session, the longer the
    # window — so waiting on the lock FILE isn't enough). Retry CompactRepair itself with
    # backoff (capped total ~60s, well under this op's 120s+ caller-side timeout budget).
    $info = $null
    $maxAttempts = 8
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            # [void] is required — an unassigned COM method call result otherwise leaks onto
            # PowerShell's output stream and becomes part of this function's return value,
            # silently corrupting the JSON response (this was a preexisting, latent bug: it
            # never surfaced before because CompactRepair had never actually succeeded here).
            [void]$script:app.CompactRepair($sourcePath, $tempPath, $false)
            $info = $null
            break
        } catch {
            $info = Get-ErrorInfo $_
            if ($info.code -ne 'DB_LOCKED' -or $attempt -eq $maxAttempts) { break }
            if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Milliseconds ([Math]::Min(2000 * $attempt, 10000))
        }
    }
    if ($null -ne $info) {
        # Access's CompactRepair insists on (re)compiling the VBA project and fails with
        # "Cannot Compile Project." when the project is password-locked or uncompilable.
        # The DAO engine compacts the same file without touching VBA — try that before
        # giving up (verified: recovers a bloated file with a locked project intact).
        $daoCompacted = $false
        try {
            if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
            $dao = New-Object -ComObject DAO.DBEngine.120
            try {
                [void]$dao.CompactDatabase($sourcePath, $tempPath)
                $daoCompacted = $true
            } finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dao)
            }
        } catch {
            [Console]::Error.WriteLine("bridge warning: DAO CompactDatabase fallback failed: $($_.Exception.Message)")
        }
        if (-not $daoCompacted) {
            try { [void](Open-DbShared $sourcePath) } catch { }
            if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
            return @{ _error = @{ code = $info.code; number = $info.number; hresult = $info.hresult; message = "Compact failed, database reopened unchanged: $($info.message)" } }
        }
    }

    # Move-Item -Force is a single rename/replace, not delete-then-copy: if it fails, the
    # original file at $sourcePath is still intact and the compacted result stays at $tempPath
    # (reported in the error) rather than being lost.
    try {
        Move-Item -LiteralPath $tempPath -Destination $sourcePath -Force
    } catch {
        $info = Get-ErrorInfo $_
        try { [void](Open-DbShared $sourcePath) } catch { }
        return @{ _error = @{ code = 'COM_ERROR'; number = 0; hresult = 0; message = "Compact succeeded (result left at $tempPath) but replacing the original file failed: $($info.message)" } }
    }

    $lockExt = if ($ext -ieq '.mdb') { 'ldb' } else { 'laccdb' }
    $lockPath = [System.IO.Path]::ChangeExtension($sourcePath, $lockExt)
    if (Test-Path -LiteralPath $lockPath) { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue }

    [void](Open-DbShared $sourcePath)
    $script:dbPath = $sourcePath
    return @{ compacted = $true; listing = (Op-List) }
}

function Close-App {
    if ($null -ne $script:app) {
        try { $script:app.CloseCurrentDatabase() } catch { }
        try { $script:app.Quit($acQuitSaveNone) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:app) } catch { }
        $script:app = $null
        $script:dbPath = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ---------- main loop ----------
try {
    while ($true) {
        $line = $stdin.ReadLine()
        if ($null -eq $line) { break }          # stdin EOF: parent died -> cleanup in finally
        $line = $line.Trim()
        if ($line -eq '') { continue }

        $reqId = $null
        try {
            $req = $line | ConvertFrom-Json
            $reqId = $req.id
            $op = [string]$req.op
            # PS 5.1 ConvertFrom-Json yields PSCustomObject; normalize args to hashtable.
            $opArgs = @{}
            if ($req.PSObject.Properties['args'] -and $null -ne $req.args) {
                foreach ($p in $req.args.PSObject.Properties) { $opArgs[$p.Name] = $p.Value }
            }

            if ($op -eq 'quit') {
                Write-Response @{ id = $reqId; ok = $true; data = @{ bye = $true } }
                break
            }

            $data = switch ($op) {
                'ping'         { @{ pong = $true; pid = $PID } }
                'open'         { Op-Open $opArgs }
                'list'         { Op-List }
                'getModule'    { Op-GetModule $opArgs }
                'saveModule'   { Op-SaveModule $opArgs }
                'getMacro'     { Op-GetMacro $opArgs }
                'saveMacro'    { Op-SaveMacro $opArgs }
                'getQuerySql'  { Op-GetQuerySql $opArgs }
                'saveQuerySql' { Op-SaveQuerySql $opArgs }
                'getTableDef'  { Op-GetTableDef $opArgs }
                'getFormDef'   { Op-GetFormDef $opArgs }
                'getReportDef' { Op-GetReportDef $opArgs }
                'saveFormDef'  { Op-SaveFormDef $opArgs }
                'saveReportDef' { Op-SaveReportDef $opArgs }
                'unlockVba'    { Op-UnlockVba $opArgs }
                'compile'      { Op-Compile $opArgs }
                'backup'       { Op-Backup $opArgs }
                'compact'      { Op-Compact $opArgs }
                'relinkCredentials' { Op-RelinkCredentials $opArgs }
                default        { throw "Unknown op: $op" }
            }

            if ($data -is [hashtable] -and $data.ContainsKey('_error')) {
                Write-Response @{ id = $reqId; ok = $false; error = $data._error }
            } else {
                Write-Response @{ id = $reqId; ok = $true; data = $data }
            }
        } catch {
            $info = Get-ErrorInfo $_
            [Console]::Error.WriteLine("bridge error: $($info.message)")
            Write-Response @{ id = $reqId; ok = $false; error = $info }
        }
    }
} finally {
    Close-App
}
