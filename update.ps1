# 1. Sincronização com o Google Apps Script
Write-Host "--- 1/3: Sincronizando com Google Sheets (Apps Script) ---" -ForegroundColor Yellow
clasp push -f

# 2. Geração do ficheiro index.html (Obrigatório para o GitHub Pages)
Write-Host "--- 2/3: Gerando index.html a partir do Template ---" -ForegroundColor Yellow
if (Test-Path "Template.html") {
    $content = Get-Content Template.html -Raw
    
    # PASSO A: Substituir includes <?!= ... ?> pelos conteúdos reais
    $matches = [regex]::Matches($content, "<\?!= include\('(.+?)'\); \?>")
    foreach ($m in $matches) {
        $f = $m.Groups[1].Value + ".html"
        if (Test-Path $f) {
            $c = Get-Content $f -Raw
            $content = $content.Replace($m.Value, $c)
            Write-Host "  - Incluído: $f" -ForegroundColor Green
        } else {
            Write-Host "  - Aviso: Ficheiro não encontrado: $f" -ForegroundColor Yellow
            $content = $content.Replace($m.Value, "")
        }
    }
    
    # PASSO B: Injetar o JS_Mock.html (Simulador de Ambiente Google)
    if (Test-Path "JS_Mock.html") {
        Write-Host "  - Injetando Mock de ambiente: JS_Mock.html" -ForegroundColor Cyan
        $mockContent = Get-Content "JS_Mock.html" -Raw
        # Inserir logo após o <head> para garantir prioridade
        if ($content -match "<head>") {
            $content = $content.Replace("<head>", "<head>`n$mockContent")
        }
    }
    
    # PASSO C: Limpar tags <?= ... ?> (Evita Syntax Errors no Browser)
    $content = [regex]::Replace($content, "<\?=\s*(.+?)\s*\?>", {
        param($match)
        return "''" # Substitui qualquer variável de servidor por string vazia
    })
    
    # PASSO D: Remover quaisquer tags restantes <?!= ... ?>
    $content = [regex]::Replace($content, "<\?!=\s*(.+?)\s*\?>", '')

    # FORÇAR o Git a esquecer o "Index" maiúsculo
    if (Test-Path "Index.html") { 
        git rm Index.html --force -q 2>$null
        Remove-Item "Index.html" -Force -ErrorAction SilentlyContinue
    }

    # Guardar o novo ficheiro final
    $content | Out-File -FilePath "index.html" -Encoding utf8 -Force
}

# 3. Sincronização com o GitHub (flowly.pt)
Write-Host "--- 3/3: A enviar para o GitHub ---" -ForegroundColor Yellow
git config core.autocrlf true
git add index.html
git add .
git commit -m "Fix: Deploy com Mock de Ambiente $(Get-Date -Format 'HH:mm')"

# Rebase e Push
Write-Host "--- Sincronizando com o servidor ---" -ForegroundColor Cyan
git pull --rebase origin main
git push origin main

Write-Host "`n🚀 SUCESSO! O sistema de Mocks já está a caminho do GitHub." -ForegroundColor Green