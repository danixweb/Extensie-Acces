$file = 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\.accdb-ai\Forms\10 COMENZI SIMPLA.form.txt'
$content = Get-Content $file -Encoding Unicode -Raw

$find = @"
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    Set db = Nothing
    'DBEngine.Rollback
1800 Exit Sub
"@

$replace = @"
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
'===Start Generat AI===
    If Not recStergeInutile Is Nothing Then: recStergeInutile.Close: Set recStergeInutile = Nothing
    If Not recSarjeMici Is Nothing Then: recSarjeMici.Close: Set recSarjeMici = Nothing
    If Not recOSinguraSarja Is Nothing Then: recOSinguraSarja.Close: Set recOSinguraSarja = Nothing
    If Not LISTA Is Nothing Then: LISTA.Close: Set LISTA = Nothing
'===Final Generat AI===
    Set db = Nothing
    'DBEngine.Rollback
1800 Exit Sub
"@

if ($content.Contains($find)) {
    $content = $content.Replace($find, $replace)
    Set-Content -Path $file -Value $content -Encoding Unicode
    Write-Host "Replacement successful."
} else {
    Write-Host "Target block not found."
    # Let's print the area where it should be
    $idx = $content.IndexOf("1790 Forms![10 COMENZI SIMPLA]!Child494.SourceObject")
    if ($idx -ge 0) {
        Write-Host $content.Substring($idx, 500)
    }
}
