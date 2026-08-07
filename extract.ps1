$lines = Get-Content -Path 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\.accdb-ai\Forms\10 COMENZI SIMPLA.form.txt' -Encoding Unicode
$startIndex = -1
$endIndex = -1

for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match 'Sub Command565_Click\(') {
        $startIndex = $i
    }
    if ($startIndex -ge 0 -and $i -gt $startIndex -and $lines[$i] -match 'End Sub') {
        $endIndex = $i
        break
    }
}

if ($startIndex -ge 0 -and $endIndex -ge 0) {
    $lines[$startIndex..$endIndex] | Out-File -Encoding utf8 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\Command565_Click.txt'
    Write-Host "Extracted from line $($startIndex + 1) to $($endIndex + 1)"
} else {
    Write-Host "Not found"
}
