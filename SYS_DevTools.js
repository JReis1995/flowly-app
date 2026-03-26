// A nossa cascata de modelos (Do melhor/mais recente para o mais leve/rápido)
const GEMINI_MODELS = [
  "models/gemini-3.1-pro-preview",   // O "Cérebro" mais potente para a tua auditoria
  "models/gemini-2.5-pro",           // A versão estável de alta capacidade
  "models/gemini-2.5-flash"          // O plano B mais rápido e fiável
];

/**
 * 1. ATUALIZA O CACHE (MEMÓRIA) COM FALLBACK AUTOMÁTICO
 */
function updateGeminiCache() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const context = buildProjectContext();
  const ui = SpreadsheetApp.getUi();

  // Vamos tentar um modelo de cada vez
  for (let i = 0; i < GEMINI_MODELS.length; i++) {
    const currentModel = GEMINI_MODELS[i];
    const url = `https://generativelanguage.googleapis.com/v1beta/cachedContents?key=${apiKey}`;
    
    const payload = {
      "model": currentModel,
      "contents": [{ "role": "user", "parts": [{ "text": context }] }],
      "ttl": "3600s"
    };

    const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());

    // Se a Google aceitar, guardamos o ID do Cache e o Modelo que ganhou!
    if (result.name) {
      props.setProperty('LAST_GEMINI_CACHE_ID', result.name);
      props.setProperty('ACTIVE_GEMINI_MODEL', currentModel); // Memoriza o vencedor
      ui.alert(`✅ Memória atualizada com sucesso!\n\n🤖 Modelo ativo: ${currentModel}`);
      return; // Sai da função com sucesso
    } else {
      // Se for o último modelo da lista e mesmo assim falhar...
      if (i === GEMINI_MODELS.length - 1) {
        ui.alert("❌ Todos os modelos estão indisponíveis! Erro final:\n\n" + response.getContentText());
      }
      // Se não for o último, o loop continua e tenta o próximo modelo silenciosamente
    }
  }
}

/**
 * 2. GERA INSTRUÇÕES (Lê a linha onde tu estás)
 */
function gerarInstrucoesAntigravity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AI_Console");
  const ui = SpreadsheetApp.getUi();

  const activeCell = sheet.getActiveCell();
  const currentRow = activeCell.getRow();

  if (currentRow < 2) {
    ui.alert("Por favor, clica primeiro na linha da pergunta (Linha 2 ou inferior).");
    return;
  }

  const perguntaOriginal = sheet.getRange(currentRow, 1).getValue();
  if (!perguntaOriginal) {
    ui.alert("A célula da pergunta (Coluna A) está vazia nesta linha!");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const cacheId = props.getProperty('LAST_GEMINI_CACHE_ID');
  
  // LÊ O MODELO QUE GANHOU NA FASE DO CACHE
  const activeModel = props.getProperty('ACTIVE_GEMINI_MODEL'); 

  if (!cacheId || !activeModel) {
    ui.alert("Erro: Faz primeiro o 'Atualizar Memória'.");
    return;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/${activeModel}:generateContent?key=${apiKey}`;

  const systemInstruction = `És o Arquiteto e Engenheiro SWE (Software Engineering) Sénior do Flowly 360. 
    A tua missão é analisar o pedido e decidir a estratégia de entrega.

    REGRA DE TRIAGEM:
    1. Se o pedido for SIMPLES (ex: mudar uma cor, um texto, ou uma função isolada), devolve o [PROMPT] geral.
    2. Se o pedido for COMPLEXO ou envolver RISCO de quebrar a arquitetura (ex: auth, routing, integração GitHub/GAS), IGNORA o [PROMPT] geral e devolve APENAS os [SPRINTS].

    REGRAS PARA CADA SPRINT (Modo 1 Chat por Sprint):
    Como cada sprint será executado num chat NOVO e LIMPO do Windsurf, cada sprint DEVE ser auto-contido.
    Cada [SPRINT] deve obrigatoriamente incluir:
    - CONTEXTO TÉCNICO: Explica à IA do Windsurf o que está por trás (ex: "O site corre no GitHub mas valida no GAS").
    - FICHEIROS AFETADOS: Lista exata.
    - INSTRUÇÃO DE CÓDIGO: O comando técnico direto.
    - REGRAS DE OURO: Coisas que ela NÃO pode mudar (ex: "Não alteres o Regex do update.ps1").
    - CRITÉRIOS DE ACEITAÇÃO: Como validar se o sprint terminou com sucesso.

    ESTRUTURA DE RESPOSTA RIGOROSA:
    [ANALISE] (Breve avaliação da complexidade e risco)
    [PROMPT] (Apenas se for simples, senão escreve "N/A - Requer Sprints")
    [SPRINTS] (Lista de blocos autónomos para copiar/colar)
    `;


  const payload = {
    "cachedContent": cacheId,
    "contents": [{
      "role": "user",
      "parts": [{ "text": systemInstruction + "\n\nPERGUNTA: " + perguntaOriginal }]
    }],
    "generationConfig": {
      "maxOutputTokens": 8192,
      "temperature": 0.2,
      "topP": 0.8,
      "topK": 40
    }
  };

  const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const res = JSON.parse(response.getContentText());

  if (!res.candidates) {
    ui.alert(`❌ O modelo ${activeModel} falhou a gerar a resposta! Erro real:\n\n` + response.getContentText());
    return;
  }

  const txt = res.candidates[0].content.parts[0].text;
  const getTag = (tag) => {
    const regex = new RegExp(`\\[${tag}\\]([\\s\\S]*?)(?=\\[|$)`, 'i');
    const match = txt.match(regex);
    return match ? match[1].trim() : "";
  };

  sheet.getRange(currentRow, 2).setValue(getTag("MODELO"));
  sheet.getRange(currentRow, 3).setValue(getTag("ANALISE"));
  sheet.getRange(currentRow, 4).setValue(getTag("PROMPT"));
  sheet.getRange(currentRow, 5).setValue(getTag("SPRINTS"));
}

function buildProjectContext() {
  const scriptId = ScriptApp.getScriptId();
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const options = { "method": "get", "headers": { "Authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const response = UrlFetchApp.fetch(url, options);
  const files = JSON.parse(response.getContentText()).files;
  let code = "PROJETO FLOWLY 360:\n\n";
  files.forEach(f => { if (!['appsscript', 'SYS_DevTools'].includes(f.name)) code += `// FILE: ${f.name}\n${f.source}\n\n`; });
  return code;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 Flowly AI')
    .addItem('🔄 Atualizar Memória (Cache)', 'updateGeminiCache')
    .addSeparator()
    .addItem('💬 Abrir Chat DevTool', 'abrirChatModal') // <-- NOVO BOTÃO AQUI
    .addToUi();
}

/**
 * 3. ABRE A JANELA DE CHAT MODAL
 */
function abrirChatModal() {
  // Cria a interface a partir do ficheiro HTML que vamos criar a seguir
  const html = HtmlService.createHtmlOutputFromFile('ChatModal')
      .setWidth(800)  // Janela larga para leres bem o código
      .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Flowly 360 - AI DevTool');
}

/**
 * 4. PROCESSA A MENSAGEM VINDA DO CHAT HTML
 */
function enviarParaGeminiChat(mensagemUtilizador) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const cacheId = props.getProperty('LAST_GEMINI_CACHE_ID');
  const activeModel = props.getProperty('ACTIVE_GEMINI_MODEL') || "models/gemini-2.5-flash";

  if (!cacheId) {
    return "❌ Erro: Memória vazia. Por favor, corre o 'Atualizar Memória (Cache)' no menu primeiro.";
  }

  // O teu Prompt Implacável
  const systemInstruction = `És o Arquiteto e Engenheiro SWE (Software Engineering) Sénior do Flowly 360.
  A tua missão é analisar o pedido e devolver Sprints técnicos rigorosos.
  Garante que nunca cortas código a meio. Formata os blocos de código com clareza.`;

  const url = `https://generativelanguage.googleapis.com/v1beta/${activeModel}:generateContent?key=${apiKey}`;
  const payload = {
    "cachedContent": cacheId,
    "contents": [{
      "role": "user",
      "parts": [{ "text": systemInstruction + "\n\nPEDIDO DO DEV: " + mensagemUtilizador }]
    }],
    "generationConfig": {
      "maxOutputTokens": 8192,
      "temperature": 0.2
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, { 
      "method": "post", 
      "contentType": "application/json", 
      "payload": JSON.stringify(payload), 
      "muteHttpExceptions": true 
    });
    
    const res = JSON.parse(response.getContentText());

    if (res.candidates && res.candidates.length > 0) {
      return res.candidates[0].content.parts[0].text;
    } else {
      return "❌ Falha na IA: " + response.getContentText();
    }
  } catch (e) {
    return "❌ Erro de Ligação: " + e.message;
  }
}