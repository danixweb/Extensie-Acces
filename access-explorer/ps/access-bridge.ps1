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

Add-Type -Name Win32 -Namespace Bridge -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
'@

$script:app = $null
$script:accessPid = 0
$script:dbPath = $null

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
    $hr = 0
    try { $hr = $ex.HResult } catch { }
    $num = 0
    # DAO/Access automation errors use FACILITY_CONTROL (0xA); low word = error number.
    if (($hr -band 0xFFFF0000) -eq 0x800A0000) { $num = $hr -band 0xFFFF }
    $code = 'COM_ERROR'
    if ($hr -eq 0x80040154 -or $hr -eq -2147221164) { $code = 'ACCESS_NOT_INSTALLED' }
    elseif ($num -in 3045, 3050, 3704, 3734, 3009, 3211, 3260) { $code = 'DB_LOCKED' }
    elseif ($num -in 3051, 3033) { $code = 'NO_PERMISSION' }
    elseif ($num -in 3024, 3044, 3055) { $code = 'DB_NOT_FOUND' }
    elseif ($num -in 2544, 3265, 7874, 2103, 29068) { $code = 'OBJECT_NOT_FOUND' }
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

function Get-Db { return $script:app.CurrentDb() }

function Op-Open([hashtable]$args_) {
    $path = [string]$args_.path
    if (-not (Test-Path -LiteralPath $path)) {
        return @{ _error = @{ code = 'DB_NOT_FOUND'; number = 0; hresult = 0; message = "File not found: $path" } }
    }
    $script:app = New-Object -ComObject Access.Application
    $script:app.Visible = $false
    try { $script:app.UserControl = $false } catch { }
    # Resolve the real MSACCESS.EXE PID so the extension can kill it if the bridge hangs.
    # hWndAccessApp surfaces as a method through PowerShell COM late binding.
    try {
        $hwnd = [System.IntPtr][int]$script:app.hWndAccessApp()
        $procId = [uint32]0
        [Bridge.Win32]::GetWindowThreadProcessId($hwnd, [ref]$procId) | Out-Null
        $script:accessPid = $procId
    } catch {
        [Console]::Error.WriteLine("bridge warning: could not resolve Access PID: $($_.Exception.Message)")
        $script:accessPid = 0
    }
    # $false = shared mode: coexists with the Access UI having the file open normally.
    $script:app.OpenCurrentDatabase($path, $false)
    $script:dbPath = $path
    return @{ accessPid = $script:accessPid; path = $path }
}

function Op-List {
    Assert-Open
    $db = Get-Db
    $tables = @()
    $db.TableDefs.Refresh()
    foreach ($t in $db.TableDefs) {
        $n = $t.Name
        if ($n -like 'MSys*' -or $n -like '~*' -or $n -like 'f_*ADO*') { continue }
        $tables += $n
    }
    $queries = @()
    $db.QueryDefs.Refresh()
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -like '~*') { continue }
        $queries += $q.Name
    }
    $forms = @();   foreach ($o in $script:app.CurrentProject.AllForms)   { $forms   += $o.Name }
    $reports = @(); foreach ($o in $script:app.CurrentProject.AllReports) { $reports += $o.Name }
    $macros = @();  foreach ($o in $script:app.CurrentProject.AllMacros)  { $macros  += $o.Name }
    $modules = @(); foreach ($o in $script:app.CurrentProject.AllModules) { $modules += $o.Name }
    return @{
        tables  = @($tables  | Sort-Object)
        queries = @($queries | Sort-Object)
        forms   = @($forms   | Sort-Object)
        reports = @($reports | Sort-Object)
        macros  = @($macros  | Sort-Object)
        modules = @($modules | Sort-Object)
    }
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

function Op-GetModule([hashtable]$args_) {
    Assert-Open
    $name = [string]$args_.name
    $r = Export-ObjectText $acModule $name $args_.file
    if ($r.text.Length -gt 0) {
        # Newer Access versions include the VERSION/Attribute header; detect class-ness
        # from it, refining via VBE when available.
        $isClass = $r.text -match '(?m)^\s*VERSION \d+\.\d+ CLASS\s*$'
        try { $isClass = ((Get-VbComponent $name).Type -eq 2) } catch { }
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
        return @{ isClass = ($c.Type -eq 2); viaVbe = $true }
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
        return @{ saved = $true }
    }
    # Modules read through the VBE must be written back through it too —
    # LoadFromText would replace the class module with a broken standard module.
    $text = [System.IO.File]::ReadAllText($args_.file, $Utf8NoBom)
    $cm = (Get-VbComponent $name).CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    if ($text.Length -gt 0) { $cm.AddFromString($text) }
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
    $db.QueryDefs.Refresh()
    $sql = $db.QueryDefs($args_.name).SQL
    [System.IO.File]::WriteAllText($args_.file, $sql, $Utf8NoBom)
    return @{ }
}

function Op-SaveQuerySql([hashtable]$args_) {
    Assert-Open
    $sql = [System.IO.File]::ReadAllText($args_.file, $Utf8NoBom)
    $db = Get-Db
    $db.QueryDefs.Refresh()
    # DAO validates the SQL at assignment and throws (e.g. 3129) on bad syntax.
    $db.QueryDefs($args_.name).SQL = $sql
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
    [void]$sb.AppendLine("Records: $($t.RecordCount)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Fields:')
    foreach ($f in $t.Fields) {
        $tn = if ($daoTypes.ContainsKey([int]$f.Type)) { $daoTypes[[int]$f.Type] } else { "Type $($f.Type)" }
        $req = if ($f.Required) { ' REQUIRED' } else { '' }
        [void]$sb.AppendLine(("  {0,-32} {1}({2}){3}" -f $f.Name, $tn, $f.Size, $req))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Indexes:')
    foreach ($ix in $t.Indexes) {
        $flags = @()
        if ($ix.Primary) { $flags += 'PRIMARY' }
        if ($ix.Unique)  { $flags += 'UNIQUE' }
        $cols = @(); foreach ($f in $ix.Fields) { $cols += $f.Name }
        [void]$sb.AppendLine(("  {0,-32} ({1}) {2}" -f $ix.Name, ($cols -join ', '), ($flags -join ' ')))
    }
    [System.IO.File]::WriteAllText($args_.file, $sb.ToString(), $Utf8NoBom)
    return @{ }
}

function Op-GetFormDef([hashtable]$args_)   { Assert-Open; Export-ObjectText $acForm   $args_.name $args_.file | Out-Null; return @{ } }
function Op-GetReportDef([hashtable]$args_) { Assert-Open; Export-ObjectText $acReport $args_.name $args_.file | Out-Null; return @{ } }

function Op-Compile([hashtable]$args_) {
    Assert-Open
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
                'compile'      { Op-Compile $opArgs }
                'backup'       { Op-Backup $opArgs }
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
