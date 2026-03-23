# 1. Google (GAS)
Write-Host "--- Sincronizando com Google Sheets ---" -ForegroundColor Yellow
clasp push -f

# 2. Criar o Index para o GitHub (Unificando os ficheiros)
Write-Host "--- Gerando index.html para o flowly.pt ---" -ForegroundColor Yellow
if (Test-Path "Template.html") {
    $content = Get-Content Template.html -Raw
    $matches = [regex]::Matches($content, "<\?!= include\('(.+?)'\); \?>")
    foreach ($m in $matches) {
        $f = $m.Groups[1].Value + ".html"
        if (Test-Path $f) {
            $c = Get-Content $f -Raw
            $content = $content.Replace($m.Value, $c)
        }
    }
    # Este é o ficheiro que o GitHub vai ler
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. GitHub
Write-Host "--- Enviando para o GitHub ---" -ForegroundColor Yellow
git add .
git commit -m "Correção de CSP e unificação de template"
git push origin main
Write-Host "🚀 TUDO PRONTO! Verifica o flowly.pt agora." -ForegroundColor Green