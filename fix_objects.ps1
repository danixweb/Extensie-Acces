$file = 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\.accdb-ai\Forms\10 COMENZI SIMPLA.form.txt'
$lines = Get-Content $file
$outLines = @()

$currentRoutine = ""
$recordsets = @{}
$routineLines = @()

foreach ($line in $lines) {
    if ($line -match "^\s*(?:(?:Private|Public|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([a-zA-Z0-9_]+)") {
        if ($currentRoutine) {
            $outLines += $routineLines
        }
        $currentRoutine = $matches[2]
        $recordsets = @{}
        $routineLines = @($line)
        continue
    }
    
    if ($currentRoutine) {
        $routineLines += $line
        
        if ($line -match "^\s*(?:\d+\s+)?Set\s+([a-zA-Z0-9_]+)\s*=\s*(.+)") {
            $varName = $matches[1]
            $rhs = $matches[2]
            if ($rhs -notmatch "Nothing") {
                if ($rhs -match "OpenRecordset|ADODB\.Recordset|\.Open") {
                    $recordsets[$varName] = $true
                }
            }
        }
        
        if ($line -match "^\s*(?:\d+\s+)?([a-zA-Z0-9_]+)\.Close") {
            $varName = $matches[1]
            $recordsets.Remove($varName)
        }
        
        if ($line -match "^\s*(?:\d+\s+)?End (Sub|Function|Property)") {
            if ($recordsets.Count -gt 0) {
                $closeBlock = @()
                $closeBlock += "    '===Start Generat AI==="
                $closeBlock += "    On Error Resume Next"
                foreach ($key in $recordsets.Keys) {
                    $closeBlock += "    $key.Close"
                }
                $closeBlock += "    '===Final Generat AI==="
                
                $insertIdx = -1
                for ($i = $routineLines.Count - 1; $i -ge 0; $i--) {
                    if ($routineLines[$i] -match "TRATARE_ERORI_iesire:") {
                        $insertIdx = $i + 1
                        break
                    }
                }
                
                if ($insertIdx -ne -1) {
                    $newLines = @()
                    for ($i = 0; $i -lt $routineLines.Count; $i++) {
                        $newLines += $routineLines[$i]
                        if ($i -eq ($insertIdx - 1)) {
                            $newLines += $closeBlock
                        }
                    }
                    $routineLines = $newLines
                } else {
                    $errLabelIdx = -1
                    for ($i = $routineLines.Count - 1; $i -ge 0; $i--) {
                        if ($routineLines[$i] -match "TRATARE_ERORI:") {
                            $errLabelIdx = $i
                            break
                        }
                    }
                    
                    if ($errLabelIdx -ne -1) {
                        $exitIdx = -1
                        for ($i = $errLabelIdx - 1; $i -ge 0; $i--) {
                            if ($routineLines[$i] -match "Exit (Sub|Function|Property)") {
                                $exitIdx = $i
                                break
                            }
                        }
                        if ($exitIdx -ne -1) {
                            $newLines = @()
                            for ($i = 0; $i -lt $routineLines.Count; $i++) {
                                if ($i -eq $exitIdx) {
                                    $newLines += $closeBlock
                                }
                                $newLines += $routineLines[$i]
                            }
                            $routineLines = $newLines
                        } else {
                            $newLines = @()
                            for ($i = 0; $i -lt $routineLines.Count; $i++) {
                                if ($i -eq $errLabelIdx) {
                                    $newLines += $closeBlock
                                }
                                $newLines += $routineLines[$i]
                            }
                            $routineLines = $newLines
                        }
                    } else {
                        $newLines = @()
                        for ($i = 0; $i -lt $routineLines.Count; $i++) {
                            if ($i -eq ($routineLines.Count - 1)) {
                                $newLines += $closeBlock
                            }
                            $newLines += $routineLines[$i]
                        }
                        $routineLines = $newLines
                    }
                }
            }
            
            $outLines += $routineLines
            $currentRoutine = ""
        }
    } else {
        $outLines += $line
    }
}

Set-Content -Path $file -Value $outLines -Encoding Unicode
Write-Host "Modified $file successfully."
