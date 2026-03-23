const GEMINI_MODEL = "models/gemini-flash-latest";

/**
 * 1. ATUALIZA O CACHE (MEMÓRIA)
 */
function updateGeminiCache() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const context = buildProjectContext();
  const url = `https://generativelanguage.googleapis.com/v1beta/cachedContents?key=${apiKey}`;

  const payload = {
    "model": GEMINI_MODEL,
    "contents": [{ "role": "user", "parts": [{ "text": context }] }],
    "ttl": "3600s"
  };

  const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const result = JSON.parse(response.getContentText());

  if (result.name) {
    PropertiesService.getScriptProperties().setProperty('LAST_GEMINI_CACHE_ID', result.name);
    SpreadsheetApp.getUi().alert("✅ Memória atualizada com sucesso!");
  } else {
    SpreadsheetApp.getUi().alert("❌ Erro ao atualizar: " + response.getContentText());
  }
}

/**
 * 2. GERA INSTRUÇÕES (Lê a linha onde tu estás)
 */
function gerarInstrucoesAntigravity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AI_Console");

  // Agora ele lê EXATAMENTE a célula onde tens o retângulo azul
  const activeCell = sheet.getActiveCell();
  const currentRow = activeCell.getRow();

  if (currentRow < 2) {
    SpreadsheetApp.getUi().alert("Por favor, clica primeiro na linha da pergunta (Linha 2 ou inferior).");
    return;
  }

  const perguntaOriginal = sheet.getRange(currentRow, 1).getValue();
  if (!perguntaOriginal) {
    SpreadsheetApp.getUi().alert("A célula da pergunta (Coluna A) está vazia nesta linha!");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const cacheId = props.getProperty('LAST_GEMINI_CACHE_ID');
  const url = `https://generativelanguage.googleapis.com/v1beta/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const systemInstruction = `És o Arquiteto Sénior do Flowly 360. 
    Analisa o pedido e responde RIGOROSAMENTE com estas 4 tags:
    [MODELO] Nome da IA sugerida.
    [ANALISE] Breve análise técnica.
    [PROMPT] Prompt para o Antigravity.
    [SPRINTS] Lista de sprints.`;

  const payload = {
    "cachedContent": cacheId,
    "contents": [{
      "role": "user",
      "parts": [{ "text": systemInstruction + "\n\nPERGUNTA: " + perguntaOriginal }]
    }],
    "generationConfig": {
      "maxOutputTokens": 8192, // Aumentámos para o máximo sugerido para lógica
      "temperature": 0.2,      // Menos "criatividade", mais precisão técnica
      "topP": 0.8,
      "topK": 40
    }
  };

  const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const res = JSON.parse(response.getContentText());

  if (!res.candidates) {
    SpreadsheetApp.getUi().alert("Erro: Faz primeiro o 'Atualizar Memória'.");
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

// ATENÇÃO: COMENTEI O OPOPEN PARA NÃO CONFLITUAR COM O TEU ERP

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 Flowly AI')
    .addItem('🔄 Atualizar Memória (Cache)', 'updateGeminiCache')
    .addSeparator()
    .addItem('🧠 Gerar Prompt Antigravity', 'gerarInstrucoesAntigravity')
    .addToUi();
}
