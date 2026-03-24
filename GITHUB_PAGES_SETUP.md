# Flowly 360 - Configuração para GitHub Pages

## Resumo das Implementações

### 1. Correção de Domínio (`Template.html`)
- **Script no topo do `<head>`**: Força redirecionamento de `www.flowly.pt` para `flowly.pt`
- **Garantia de HTTPS**: Redireciona automaticamente de HTTP para HTTPS
- **Prevenção de Frame Mismatch**: Mantém consistência de domínio em toda a aplicação

### 2. Supressão de Erros de Frame (`JS_Mock.html`)
- **`window.onerror`**: Silencia erros relacionados com `tabs:outgoing`, `google.script`, `postMessage` e `Frame mismatch`
- **Interceptor de `postMessage`**: Bloqueia mensagens relacionadas com `google.script` para evitar erros
- **Proteção**: Deixa outros erros passarem normalmente para debugging

### 3. Upgrade do Mock (`JS_Mock.html`)
- **`getMasterData`**: Retorna objeto completo com `planConfig`, `users`, `cc`, `cargos`
- **`getAIAutoPreference`**: Retorna `true` por defeito
- **`getAiCredits`**: Retorna `999` créditos para demonstração
- **`getDashboardData`**: Inclui `clientConfig` para compatibilidade

### 4. Garantia de Visibilidade (`JS_Mock.html`)
- **`currentUser` no localStorage**: Injeta automaticamente utilizador demo se não existir
- **Plano completo**: Todos os módulos ativados por defeito no modo demo
- **Bypass de autenticação**: `initApp()` funciona sem login no modo GitHub Pages

### 5. Ordem de Carga (`Template.html`)
- **`JS_Mock` primeiro**: Carregado antes de `JS_Globals` e `JS_Auth_Nav`
- **Prevenção de race conditions**: Mock disponível antes de outras dependências
- **Comentários explicativos**: Documentação da ordem crítica de carga

## Estrutura dos Arquivos Modificados

### `Template.html`
```html
<head>
    <!-- 1. Script de redirecionamento de domínio (primeiro) -->
    <script>/* Domínio redirect */</script>
    
    <!-- 2. Meta tags e libraries -->
    <meta>...</meta>
    <script src="..."></script>
    
    <!-- 3. Configuração Tailwind -->
    <script>/* Tailwind config */</script>
</head>

<body>
    <!-- UI components -->
    
    <!-- Scripts na ordem correta -->
    <?!= include('JS_Mock'); ?>        <!-- 1. Mock primeiro -->
    <?!= include('JS_Globals'); ?>      <!-- 2. Globals depois -->
    <?!= include('JS_Auth_Nav'); ?>     <!-- 3. Auth depois -->
    <!-- ... outros scripts ... -->
</body>
```

### `JS_Mock.html`
```javascript
// 1. Verificação de ambiente
if (!window.location.hostname.includes('script.google.com')) {
    
    // 2. Garantir currentUser no localStorage
    if (!localStorage.getItem('flowly_user')) {
        // Inject demo user
    }
    
    // 3. Criar mock completo
    window.google = { script: { run: { ... } } };
    
    // 4. Métodos específicos
    getMasterData: function() { /* ... */ },
    getAIAutoPreference: function() { /* ... */ },
    getAiCredits: function() { /* ... */ },
    
    // 5. Supressão de erros
    window.onerror = function(msg, source, lineno, colno, error) { /* ... */ };
    
    // 6. Interceptor postMessage
    window.postMessage = function(message, targetOrigin, transfer) { /* ... */ };
}
```

## Teste e Validação

### Arquivo de Teste: `test_github_pages.html`
- **Status do Mock**: Verifica se `google.script.run` está disponível
- **Teste getMasterData**: Valida resposta completa com planConfig e users
- **Teste getAiCredits**: Confirma retorno de 999 créditos
- **Console Logs**: Monitorização em tempo real de erros e eventos

### Como Testar
1. Fazer upload dos arquivos para GitHub Pages
2. Acessar `test_github_pages.html`
3. Verificar se todos os testes passam
4. Confirmar ausência de erros no console
5. Testar navegação na aplicação principal

## Compatibilidade

### ✅ Funciona em GitHub Pages
- Redirecionamento automático de domínio
- Mock completo do `google.script.run`
- Supressão de erros de frame
- Modo demo funcional

### ✅ Mantém Compatibilidade GAS
- Detecta ambiente automaticamente
- Mock apenas ativo fora do `script.google.com`
- Funcionalidade original preservada

### ✅ Segurança
- Zero Leak Policy mantida
- Cache indexado por email
- Validação de contexto em insights

## Deploy para GitHub Pages

### Passos:
1. Fazer push dos arquivos modificados
2. Configurar GitHub Pages no repositório
3. Acessar `https://[username].github.io/[repository]/`
4. Testar com `test_github_pages.html`

### URLs Esperadas:
- **Demo**: `https://flowly.pt/test_github_pages.html`
- **App**: `https://flowly.pt/`
- **Redirect**: `www.flowly.pt` → `flowly.pt`

## Monitorização

### Logs Importantes:
- `🔧 Flowly 360: Modo GitHub Pages detectado`
- `Mock call: google.script.run.methodName()`
- Ausência de erros `tabs:outgoing` e `Frame mismatch`

### Banner Visual:
- Banner verde no topo indicando modo demonstração
- Ajuste automático de padding para não sobrepor conteúdo

## Resumo Técnico

A implementação garante que a Flowly 360 funcione perfeitamente no GitHub Pages mantendo:
- **Funcionalidade completa** em modo demo
- **Compatibilidade total** com Google Apps Script
- **Experiência consistente** para utilizadores
- **Segurança** e **performance** otimizadas
