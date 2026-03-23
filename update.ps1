# 1. Enviar para o Google (GAS)
Write-Host "--- Sincronizando com Google Sheets ---" -ForegroundColor Yellow
clasp push -f

# 2. Gerar o ficheiro CERTO para o GitHub
Write-Host "--- Gerando index.html (Obrigatório para GitHub) ---" -ForegroundColor Yellow
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
    # GUARDAR COMO index.html (Sem o _gh)
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. Enviar para o GitHub
Write-Host "--- Enviando para o GitHub (flowly.pt) ---" -ForegroundColor Yellow
git add .
git commit -m "Fix: Nome do index e fontes CDN"
git push origin main
Write-Host "🚀 TUDO PRONTO! Faz Refresh no flowly.pt em 1 minuto." -ForegroundColor Green