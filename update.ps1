# 1. Google Apps Script
Write-Host "--- A enviar para o Google (GAS) ---" -ForegroundColor Yellow
clasp push -f

# 2. Preparar versão unificada para GitHub
if (Test-Path "Index.html") {
    Write-Host "--- A processar Index.html para o GitHub ---" -ForegroundColor Yellow
    $content = Get-Content Index.html -Raw
    # Procura por <?!= include('NOME'); ?> e substitui pelo conteúdo do ficheiro
    $matches = [regex]::Matches($content, "<\?!= include\('(.+?)'\); \?>")
    foreach ($m in $matches) {
        $f = $m.Groups[1].Value + ".html"
        if (Test-Path $f) {
            $c = Get-Content $f -Raw
            $content = $content.Replace($m.Value, $c)
        }
    }
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. GitHub
Write-Host "--- A enviar para o GitHub (flowly.pt) ---" -ForegroundColor Yellow
git add .
git commit -m "Sincronização modular completa"
git push origin main
Write-Host "✅ SUCESSO: App e Site atualizados!" -ForegroundColor Green