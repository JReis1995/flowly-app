# 1. Sincronização com o Google Apps Script
Write-Host "--- 1/3: Sincronizando com Google Sheets (Apps Script) ---" -ForegroundColor Yellow
clasp push -f

# 2. Geração do ficheiro index.html (Obrigatório para o GitHub Pages)
Write-Host "--- 2/3: Gerando index.html a partir do Template ---" -ForegroundColor Yellow
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

    # FORÇAR o Git a esquecer o "Index" maiúsculo e criar o "index" minúsculo
    if (Test-Path "Index.html") { 
        git rm Index.html --force -q 2>$null
        Remove-Item "Index.html" -Force -ErrorAction SilentlyContinue
    }

    # Guardar o novo ficheiro
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. Sincronização com o GitHub (flowly.pt)
Write-Host "--- 3/3: A enviar para o GitHub ---" -ForegroundColor Yellow
git config core.autocrlf true
git add index.html
git add .
git commit -m "Fix: Deploy automático index.html $(Get-Date -Format 'HH:mm')"

# Puxar alterações remotas (rebase) para evitar rejeição
Write-Host "--- Sincronizando com o servidor ---" -ForegroundColor Cyan
git pull --rebase origin main
git push origin main

Write-Host "`n🚀 SUCESSO ABSOLUTO! Verifica o separador 'Actions' no GitHub em 30 segundos." -ForegroundColor Green