$content = Get-Content 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\.accdb-ai\Forms\10 COMENZI SIMPLA.form.txt' -Encoding utf8 -Raw
Set-Content -Path 'c:\Users\coros\OneDrive\Desktop\Extensie Acces\.accdb-ai\Forms\10 COMENZI SIMPLA.form.txt' -Value $content -Encoding Unicode
