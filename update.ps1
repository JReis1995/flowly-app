# 1. Sincronização com o Google (GAS)
Write-Host "--- Sincronizando com Google Sheets ---" -ForegroundColor Yellow
clasp push -f

# 2. Geração do ficheiro para GitHub (FORÇANDO MINÚSCULAS)
Write-Host "--- Gerando index.html (minúsculo) ---" -ForegroundColor Yellow
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

    # REMOVER o Index.html (Maiúsculo) se ele existir para não confundir o Git
    if (Test-Path "Index.html") { Remove-Item "Index.html" -Force }

    # Guardar como index.html (Minúsculo)
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. Sincronização com GitHub (flowly.pt)
Write-Host "--- A sincronizar com GitHub ---" -ForegroundColor Yellow
git config core.autocrlf true
git add .
git commit -m "Fix: Forçando index.html em minúsculas para deploy"

# Rebase para evitar o erro de 'rejected'
Write-Host "--- A verificar alterações remotas ---" -ForegroundColor Cyan
git pull --rebase origin main
git push origin main

Write-Host "🚀 SUCESSO! O GitHub agora deve detetar o ficheiro index.html." -ForegroundColor Green