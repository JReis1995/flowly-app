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

# 3. Enviar para o GitHub (Com Rebase Automático)
Write-Host "--- A sincronizar com GitHub (flowly.pt) ---" -ForegroundColor Yellow
git config core.autocrlf true
git add .
git commit -m "Sincronização automática: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

# Tenta puxar as alterações antes de enviar, para evitar o erro 'rejected'
Write-Host "--- A verificar alterações remotas ---" -ForegroundColor Cyan
git pull --rebase origin main

# Envia para o GitHub
git push origin main

Write-Host "🚀 TUDO PRONTO! App e Site sincronizados." -ForegroundColor Green