// Ficheiro: MOD_AI.js
/// ==========================================
// 🤖 MÓDULO INTELIGÊNCIA ARTIFICIAL & OCR
// ==========================================

const MSG_IA_PENDENTE = "Configuração de IA pendente. Contacte o suporte.";

const promptInstrucoesOCR = `INSTRUÇÕES ESTRITAS DE EXTRAÇÃO FLOWLY 360:
1. EXTRAÇÃO EXAUSTIVA OBRIGATÓRIA: É estritamente proibido omitir, resumir ou agrupar artigos. Deves extrair 100% das linhas faturadas, uma a uma.
2. FUSÃO MULTI-LINHA OBRIGATÓRIA: Se o documento for um talão de supermercado/retalho e os dados de um artigo estiverem divididos (ex: Linha 1: Nome, Linha 2: Qtd x Preço), deves OBRIGATORIAMENTE fundir estas linhas num único objeto. NUNCA extraias a linha de valores sem o nome correspondente.
3. TRATAMENTO DE DESCONTOS: Linhas de desconto (ex: 'Desconto 30%', 'Poupe Já') devem ser extraídas como um artigo chamado 'DESCONTO', com quantidade 1 e preco_custo negativo.
4. PRECISÃO MATEMÁTICA: CÁLCULO DE PREÇO UNITÁRIO: Extrai ou calcula o custo unitário base (sem IVA) com exatamente 4 casas decimais. Exemplo: Se o preço final com IVA for 1.98€ e o IVA for 6%, o preco_custo no JSON deve ser 1.8679.`;

function getGeminiKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  } catch (e) { return ""; }
}

function setGeminiKey(key) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) {
    return { success: false, error: "Apenas o Super Admin pode configurar a chave." };
  }
  try {
    const k = (key && typeof key === "string") ? key.trim() : "";
    if (!k) return { success: false, error: "Chave inválida ou vazia." };
    PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", k);
    return { success: true, message: "Chave Gemini guardada com sucesso." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getAiCreditsFromMaster(email) {
  if (!email || typeof email !== 'string') return 0;
  try {
    ensureMasterDBColumns();
    var ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    var data = ss.getDataRange().getValues();
    var emailNorm = String(email).trim().toLowerCase();
    var row = data.find(function (r) { return String(r[MASTER_COL_EMAIL - 1] || '').trim().toLowerCase() === emailNorm; });
    if (!row) return 0;
    var val = row[MASTER_COL_AI_CREDITS - 1];
    var n = parseInt(val, 10);
    return isNaN(n) ? 0 : Math.max(0, n);
  } catch (e) { return 0; }
}

/** 1. Lê os créditos da conta principal (Patrão), mesmo que seja o Colaborador a pedir */
function getAiCredits(targetEmail) {
  try {
    var ctx = getClientContext(targetEmail || null);
    // ctx.clientEmail devolve sempre o email do Gestor dono da empresa
    var managerEmail = (ctx.clientEmail || "").trim().toLowerCase();
    if (!managerEmail) return 0;

    var props = PropertiesService.getScriptProperties();
    var cachedBalance = props.getProperty("FLOWLY_AI_CREDITS_" + managerEmail);
    if (cachedBalance !== null && cachedBalance !== "") {
      return parseInt(cachedBalance, 10) || 0;
    }

    ensureMasterDBColumns();
    var ss = SpreadsheetApp.openById(MASTER_DB_ID);
    var data = ss.getSheets()[0].getDataRange().getValues();
    var row = data.find(function (r) { return String(r[MASTER_COL_EMAIL - 1] || "").trim().toLowerCase() === managerEmail; });

    if (row) {
      var balance = parseInt(row[MASTER_COL_AI_CREDITS - 1], 10) || 0;
      props.setProperty("FLOWLY_AI_CREDITS_" + managerEmail, String(balance));
      return balance;
    }
    return 0;
  } catch (e) { return 0; }
}

/** 2. Consome localmente (fallback) da conta do Patrão */
function consumeAiCredit(impersonateTarget) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return getAiCredits(impersonateTarget);
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var managerEmail = (ctx.clientEmail || "").trim().toLowerCase();
    if (!managerEmail) return 0;

    var n = getAiCredits(impersonateTarget);
    if (n <= 0) return 0;
    var novo = n - 1;

    PropertiesService.getScriptProperties().setProperty("FLOWLY_AI_CREDITS_" + managerEmail, String(novo));
    return novo;
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

/** 3. Deduz permanentemente na Base de Dados Oficial, afetando o Patrão */
function deductCredits(email, amount) {
  if (!email || typeof amount !== "number" || isNaN(amount) || amount < 0) return -1;
  try {
    var ctx = getClientContext(email);
    var managerEmail = (ctx.clientEmail || "").trim().toLowerCase();
    if (!managerEmail) return -1;

    ensureMasterDBColumns();
    var ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    var data = ss.getDataRange().getValues();

    var idx = data.findIndex(function (r) { return String(r[MASTER_COL_EMAIL - 1] || "").trim().toLowerCase() === managerEmail; });
    if (idx < 0) return -1;

    var rowNum = idx + 1;
    var current = parseInt(ss.getRange(rowNum, MASTER_COL_AI_CREDITS).getValue(), 10) || 0;
    var novo = Math.max(0, current - amount);

    ss.getRange(rowNum, MASTER_COL_AI_CREDITS).setValue(novo);
    PropertiesService.getScriptProperties().setProperty("FLOWLY_AI_CREDITS_" + managerEmail, String(novo));

    return novo;
  } catch (e) { return -1; }
}

function addAiCredits(amount) {
  var n = getAiCredits();
  var added = (typeof amount === "number" && !isNaN(amount)) ? amount : 0;
  var novo = Math.max(0, n + added);
  PropertiesService.getDocumentProperties().setProperty("FLOWLY_AI_CREDITS", String(novo));
  return novo;
}

function addAiCreditsForClient(clientEmail, amount) {
  if (!clientEmail || typeof amount !== "number" || isNaN(amount) || amount <= 0) return -1;
  ensureMasterDBColumns();
  var ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
  var data = ss.getDataRange().getValues();
  var emailNorm = String(clientEmail).trim().toLowerCase();
  var idx = data.findIndex(function (r) { return String(r[MASTER_COL_EMAIL - 1] || "").trim().toLowerCase() === emailNorm; });
  if (idx < 0) return -1;
  var rowNum = idx + 1;
  var current = parseInt(ss.getRange(rowNum, MASTER_COL_AI_CREDITS).getValue(), 10) || 0;
  var novo = Math.max(0, current + amount);
  ss.getRange(rowNum, MASTER_COL_AI_CREDITS).setValue(novo);
  var clientKey = emailNorm;
  PropertiesService.getScriptProperties().setProperty("FLOWLY_AI_CREDITS_" + clientKey, String(novo));
  return novo;
}

function addCreditsToMasterDB(email, creditsToAdd) {
  var newTotal = addAiCreditsForClient(email, creditsToAdd);
  if (newTotal < 0) throw new Error("Utilizador não encontrado na Master DB para atribuir créditos.");
  return newTotal;
}

function cleanResponse(str) {
  if (str == null || typeof str !== "string") return "";
  var s = str.replace(/\uFEFF/g, "").trim();
  s = s.replace(/^```[\w]*\s*/i, "").replace(/\s*```$/i, "").trim();
  s = s.replace(/```[\w]*\s*[\s\S]*?```/gi, "").trim();
  return s;
}

function analyzeHeadersWithAI(headersArray) {
  try {
    if (!Array.isArray(headersArray) || headersArray.length === 0) return { success: false, error: "Lista de cabeçalhos vazia." };
    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada." };

    const systemPrompt = "És um assistente de engenharia de dados do ERP Flowly 360 (mercado PT). A tua tarefa é mapear os cabeçalhos de um ficheiro CSV para os campos internos da base de dados.\n" +
      "Retorna APENAS um objeto JSON válido (chave = cabeçalho original, valor = campo Flowly).\n" +
      "Campos Flowly disponíveis: 'data', 'artigo', 'categoria', 'quantidade', 'preco_custo', 'preco_venda', 'taxa_iva', 'valor_iva', 'fornecedor', 'observacoes', 'tipo', 'metodo', 'dedutivel', 'validado', 'status', 'data_pag', 'valor_pago', 'conta_stock'.\n" +
      "Regras de mapeamento (sinónimos):\n" +
      "- 'data': Data, Date, Data Documento, Data Movimento\n" +
      "- 'artigo': Designação, Nome, Produto, Descrição, Item\n" +
      "- 'quantidade': Qtd, Stock, Unidades\n" +
      "- 'preco_custo': Preço de Compra, Custo, Preço Unitário Custo, Preco\n" +
      "- 'preco_venda': PVP, Preço de Venda, Valor Venda, Venda\n" +
      "- 'taxa_iva': IVA %, Taxa IVA, Imposto, TaxaIva\n" +
      "- 'valor_iva': Montante IVA, Valor Imposto, ValorIva\n" +
      "- 'fornecedor': Entidade, Marca\n" +
      "- 'observacoes': Notas, Detalhes, Obs\n" +
      "- 'tipo': Tipo, Type, Movimento, Categoria de Movimento\n" +
      "- 'metodo': Metodo, Método, Forma Pagamento\n" +
      "- 'dedutivel': Dedutivel, Dedutível, IVA Dedutível\n" +
      "- 'validado': Validado, Estado Validação\n" +
      "- 'status': Status, Estado\n" +
      "- 'data_pag': DataPag, Data Pagamento, Data de Pagamento\n" +
      "- 'valor_pago': ValorPago, Valor Pago, Montante Pago\n" +
      "- 'conta_stock': ContaStock, Conta Stock\n" +
      "Se um cabeçalho não for relevante, mapeia para null.";

    const userPrompt = "Cabeçalhos do ficheiro: " + JSON.stringify(headersArray);

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const payload = { contents: [{ parts: [{ text: systemPrompt + "\n\n" + userPrompt }] }] };

    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
    const code = res.getResponseCode();
    const respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) {
      try {
        const errBody = JSON.parse(respText);
        const msg = (errBody.error && errBody.error.message) ? errBody.error.message : "API error " + code;
        return { success: false, error: msg };
      } catch (_) { return { success: false, error: "API error " + code }; }
    }

    const content = JSON.parse(respText);
    let rawText = (content.candidates && content.candidates[0] && content.candidates[0].content && content.candidates[0].content.parts && content.candidates[0].content.parts[0])
      ? content.candidates[0].content.parts[0].text : "";
    rawText = rawText.replace(/```json/gi, "").replace(/```/g, "").trim();
    const mapping = JSON.parse(rawText);
    return { success: true, mapping: mapping };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: errMsg.length > 80 ? "Erro na IA." : errMsg };
  }
}

function extractDataWithAI(payloadData) {
  try {
    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada." };
    let images = Array.isArray(payloadData) ? payloadData : [payloadData];
    const parts = [{
      "text": `${promptInstrucoesOCR}\nRetorna JSON estrito:[{"artigo":"Nome Item", "quantidade":1.0, "preco_custo":10.0, "taxa_iva":23, "data_doc":"DD/MM/AAAA", "fornecedor":"Nome Fornecedor"}]` }];

    images.forEach(b64 => {
      const clean = b64.includes(',') ? b64.split(',')[1] : b64;
      parts.push({ "inline_data": { "mime_type": "image/jpeg", "data": clean } });
    });

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: parts }] }), muteHttpExceptions: true });
    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) return { success: false, error: "API error " + code };

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: (content.error.message || "Quota exceeded") };

    const cand = content.candidates && content.candidates[0];
    if (!cand || !cand.content || !cand.content.parts || !cand.content.parts[0]) return { success: false, error: "Resposta inválida da API." };

    let rawText = (cand.content.parts[0].text || "").replace(/```json/g, '').replace(/```/g, '').trim();
    var data = JSON.parse(rawText);

    if (Array.isArray(data)) {
      data.forEach(function (item) { if (item && item.fornecedor) item.fornecedor = canonicalizeEntityForDisplay(item.fornecedor) || item.fornecedor; });
    } else if (data && data.fornecedor) {
      data.fornecedor = canonicalizeEntityForDisplay(data.fornecedor) || data.fornecedor;
    }

    return { success: true, data: data };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: errMsg.length > 60 ? "Erro no OCR." : errMsg };
  }
}

function extractInvoiceHeaders(base64Image) {
  try {
    const clean = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    const promptText = `EXTRATOR DE FATURAS PORTUGUESAS — MODO ERP FLOWLY 360.
    Output strictly JSON. Use dot (.) for decimals in JSON values. 
    Ensure NIF is validated against common OCR misreads (e.g., 6 vs 8, 0 vs O).

    ${promptInstrucoesOCR}
    
    PRIORIDADE 1 — NIF: Procura o NIF do FORNECEDOR/EMITENTE exclusivamente no cabeçalho/topo da fatura (nunca no rodapé). O NIF tem exactly 9 dígitos.
    PRIORIDADE 2 — TOTAIS: Usa a linha fiscal "Total c/ IVA", "Total a Pagar" ou "TOTAL".
    
    Extrai:
    1) "nif" - NIF fornecedor (string 9 dígitos)
    2) "fornecedor" - nome comercial
    3) "data" - formato (DD/MM/AAAA)
    4) "valorTotal" - total float (ponto para decimal)
    5) "valorIva" - total iva float (ponto para decimal)
    6) "litros" - se combustível (float), senão null
    
    Retorna APENAS JSON (sem markdown):
    { "cabecalho": { "nif": "123456789", "fornecedor": "Nome da Empresa", "data": "DD/MM/AAAA" }, "valorTotal": 0.00, "valorIva": 0.00, "litros": null, "linhas":[ { "artigo": "Descrição", "quantidade": 1.0, "preco_custo": 0.00, "taxa_iva": 23, "valor_iva": 0.00 } ] }
    Inclui TODAS as linhas. Se só houver total, cria linha artigo="Resumo", quantidade=1, preco_custo=valorTotal/(1+taxa_iva/100).`;

    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada." };
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const payload = { contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: "image/jpeg", data: clean } }] }] };
    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });

    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) return { success: false, error: "API error " + code };

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: (content.error.message || "Quota exceeded") };

    const cand0 = content.candidates && content.candidates[0];
    if (!cand0 || !cand0.content || !cand0.content.parts || !cand0.content.parts[0]) return { success: false, error: "Resposta inválida." };

    let rawText = cand0.content.parts[0].text.replace(/```json/gi, '').replace(/```/g, '').trim();
    const parsed = JSON.parse(rawText);

    // Intercetação matemática do NIF para robustez OCR (8 vs 6, etc)
    if (parsed.cabecalho && parsed.cabecalho.nif) {
      let nifStr = String(parsed.cabecalho.nif).replace(/\D/g, '');
      if (nifStr.length === 9) {
        let nifBase = nifStr.substring(0, 8);
        parsed.cabecalho.nif = calcularCheckDigitNIF(nifBase);
      }
    }

    if (parsed.cabecalho && parsed.cabecalho.fornecedor) parsed.cabecalho.fornecedor = canonicalizeEntityForDisplay(parsed.cabecalho.fornecedor) || parsed.cabecalho.fornecedor;

    if (!parsed.cabecalho && (parsed.fornecedor != null || parsed.nif != null)) {
      return {
        success: true,
        data: {
          cabecalho: { nif: parsed.nif || "", fornecedor: canonicalizeEntityForDisplay(parsed.fornecedor || "") || parsed.fornecedor || "", data: parsed.data || "" },
          valorTotal: parseFloat(parsed.valorTotal || parsed.valor_total) || 0,
          valorIva: parseFloat(parsed.valorIva || parsed.valor_iva) || 0,
          litros: parsed.litros != null ? parseFloat(parsed.litros) : null,
          linhas: [{ artigo: parsed.artigo || "Resumo", quantidade: 1, preco_custo: parseFloat(parsed.valor_base) || parseFloat(parsed.valor_total) || 0, taxa_iva: parseInt(parsed.taxa_iva) || 23, valor_iva: parseFloat(parsed.valor_iva) || 0 }]
        }
      };
    }

    if (!parsed.linhas || !Array.isArray(parsed.linhas)) parsed.linhas = [];
    if (parsed.valorIva == null && parsed.linhas.length > 0) {
      var sumIva = 0;
      for (var i = 0; i < parsed.linhas.length; i++) sumIva += parseFloat(parsed.linhas[i].valor_iva) || 0;
      parsed.valorIva = sumIva > 0 ? sumIva : null;
    }

    return { success: true, data: parsed };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: errMsg.length > 60 ? "Erro OCR Fatura." : errMsg };
  }
}

function classifyDocumentType(base64Image) {
  try {
    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada.", tipoDocumento: "FaturaCompra" };
    const clean = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    const promptText = `Classifica este documento numa destas 3 categorias. Responde APENAS com uma destas palavras, sem mais texto:
    FaturaCompra - fatura de compra a fornecedor, recibo de compras, documento de entrada de mercadoria
    FechoCaixa - fecho de caixa, resumo de vendas do dia, totalizador de caixa, report de faturação do dia
    RelatorioSaidas - lista de saídas de artigos, inventário de saídas, relatório de artigos que saíram
    Resposta (uma palavra apenas):`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: "image/jpeg", data: clean } }] }] }), muteHttpExceptions: true });

    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded", tipoDocumento: "FaturaCompra" };
    if (code >= 400) return { success: false, error: "API error " + code, tipoDocumento: "FaturaCompra" };

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: content.error.message || "Quota exceeded", tipoDocumento: "FaturaCompra" };

    const cand = content.candidates && content.candidates[0];
    if (!cand || !cand.content || !cand.content.parts || !cand.content.parts[0]) return { success: true, tipoDocumento: "FaturaCompra" };

    let raw = (cand.content.parts[0].text || "").trim().toLowerCase();
    if (raw.indexOf("fechocaixa") >= 0 || raw.indexOf("fecho") >= 0) return { success: true, tipoDocumento: "FechoCaixa" };
    if (raw.indexOf("relatoriosaidas") >= 0 || raw.indexOf("saidas") >= 0) return { success: true, tipoDocumento: "RelatorioSaidas" };

    return { success: true, tipoDocumento: "FaturaCompra" };
  } catch (e) {
    return { success: false, error: (e && e.message) ? e.message : String(e), tipoDocumento: "FaturaCompra" };
  }
}

function extractFechoCaixa(base64Image) {
  try {
    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada." };
    const clean = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    const promptText = `Analisa este documento de FECHO DE CAIXA / resumo de vendas. Retorna APENAS um JSON válido (sem markdown):
    { "cabecalho": { "data": "DD/MM/AAAA" }, "linhas":[ { "artigo": "Nome do item vendido", "quantidade": número, "preco_venda": preço de venda unitário (float), "preco_custo": opcional } ] }
    Inclui TODAS as linhas de artigos vendidos. Se só houver totais, cria uma linha com artigo "Vendas dia" e quantidade 1.`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: "image/jpeg", data: clean } }] }] }), muteHttpExceptions: true });

    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) return { success: false, error: "API error " + code };

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: content.error.message || "Quota exceeded" };

    const cand = content.candidates && content.candidates[0];
    if (!cand || !cand.content || !cand.content.parts || !cand.content.parts[0]) return { success: false, error: "Resposta inválida da API." };

    let rawText = (cand.content.parts[0].text || "").replace(/```json/gi, '').replace(/```/g, '').trim();
    const parsed = JSON.parse(rawText);

    if (!parsed.cabecalho) parsed.cabecalho = { data: "" };
    if (!parsed.linhas || !Array.isArray(parsed.linhas)) parsed.linhas = [];

    parsed.linhas = parsed.linhas.map(l => ({ artigo: l.artigo || "", quantidade: parseFloat(l.quantidade) || 1, preco_custo: parseFloat(l.preco_custo) || 0, preco_venda: parseFloat(l.preco_venda) || 0, taxa_iva: 23, valor_iva: 0 }));

    return { success: true, data: parsed };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: errMsg.length > 60 ? "Erro OCR Fecho Caixa." : errMsg };
  }
}

function extractRelatorioSaidas(base64Image) {
  try {
    const key = getGeminiKey();
    if (!key) return { success: false, error: "Chave Gemini não configurada." };
    const clean = base64Image.includes(',') ? base64Image.split(',')[1] : base64Image;
    const promptText = `Analisa este RELATÓRIO DE SAÍDAS de artigos (itens que saíram). Retorna APENAS um JSON válido (sem markdown):
    { "cabecalho": { "data": "DD/MM/AAAA" }, "linhas":[ { "artigo": "Nome do artigo", "quantidade": número, "preco_custo": custo unitário (float), "preco_venda": opcional } ] }
    Inclui todas as linhas de artigos que saíram.`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + key;
    const res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: "image/jpeg", data: clean } }] }] }), muteHttpExceptions: true });

    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) return { success: false, error: "API error " + code };

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: content.error.message || "Quota exceeded" };

    const cand = content.candidates && content.candidates[0];
    if (!cand || !cand.content || !cand.content.parts || !cand.content.parts[0]) return { success: false, error: "Resposta inválida da API." };

    let rawText = (cand.content.parts[0].text || "").replace(/```json/gi, '').replace(/```/g, '').trim();
    const parsed = JSON.parse(rawText);

    if (!parsed.cabecalho) parsed.cabecalho = { data: "" };
    if (!parsed.linhas || !Array.isArray(parsed.linhas)) parsed.linhas = [];

    parsed.linhas = parsed.linhas.map(l => ({ artigo: l.artigo || "", quantidade: parseFloat(l.quantidade) || 1, preco_custo: parseFloat(l.preco_custo) || 0, preco_venda: parseFloat(l.preco_venda) || 0, taxa_iva: 23, valor_iva: 0 }));

    return { success: true, data: parsed };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: errMsg.length > 60 ? "Erro OCR Relatório Saídas." : errMsg };
  }
}

function extractDocument(base64Image) {
  const classification = classifyDocumentType(base64Image);
  if (!classification.success) return { success: false, error: classification.error, tipoDocumento: "FaturaCompra" };
  const tipo = classification.tipoDocumento || "FaturaCompra";

  if (tipo === "FechoCaixa") {
    const res = extractFechoCaixa(base64Image);
    if (!res.success) return res;
    return { success: true, tipoDocumento: "FechoCaixa", data: res.data };
  }

  if (tipo === "RelatorioSaidas") {
    const res = extractRelatorioSaidas(base64Image);
    if (!res.success) return res;
    return { success: true, tipoDocumento: "RelatorioSaidas", data: res.data };
  }

  const res = extractInvoiceHeaders(base64Image);
  if (!res.success) return res;
  return { success: true, tipoDocumento: "FaturaCompra", data: res.data };
}

function getOrCreateAIHistorySheet(ss) {
  var sheet = ss.getSheetByName(AI_HISTORY_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(AI_HISTORY_TAB);
    sheet.appendRow(["Timestamp", "Email", "Category", "Analysis"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  return sheet;
}

function saveAIHistoryEntry(sheetId, userEmail, category, analysisObject) {
  if (!sheetId || !userEmail) return;
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = getOrCreateAIHistorySheet(ss);
    var jsonStr = typeof analysisObject === "string" ? analysisObject : JSON.stringify(analysisObject || {});
    sheet.appendRow([new Date(), userEmail, category || "dashboard", jsonStr]);
  } catch (e) { }
}

function getAIHistory(email, category) {
  if (!email) return [];
  try {
    var ctx = getClientContext(email);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(AI_HISTORY_TAB);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 4).getValues();
    var emailNorm = (email || "").toString().trim().toLowerCase();

    var rows = data.filter(function (row) { return String(row[1] || "").trim().toLowerCase() === emailNorm; });

    if (category && category.trim()) {
      var catNorm = category.trim().toLowerCase();
      rows = rows.filter(function (row) {
        var rowCat = String(row[2] || "").trim().toLowerCase();
        if (catNorm === "financeiro") return rowCat === "dashboard" || rowCat === "financeiro";
        return rowCat === catNorm;
      });
    }

    rows.sort(function (a, b) {
      var tA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
      var tB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
      return tB - tA;
    });

    var out = rows.slice(0, 5).map(function (row) {
      var raw = row[3];
      var obj = null;
      try {
        if (typeof raw === "string" && raw.trim()) obj = JSON.parse(raw);
        else if (raw && typeof raw === "object") obj = raw;
      } catch (e) { }
      return { timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0]), userEmail: row[1], category: String(row[2] || "dashboard").trim(), analysisObject: obj || (typeof raw === "string" ? { _raw: raw } : {}) };
    });
    return out;
  } catch (e) { return []; }
}

function _nuclearStripLabels(s) {
  if (!s || typeof s !== "string") return "";
  var t = String(s).trim();
  t = t.replace(/^[,;\s]+/, "").replace(/[,;\s]+$/, "");
  t = t.replace(/(?:^|[,;\s]+)(?:financeiro|stocks|rh|summary|todos|category)\s*[:\-]\s*/gi, " ");
  t = t.replace(/\s+/g, " ").trim();
  return t;
}

function getLastAIHistoryEntry(email) {
  if (!email) return null;
  try {
    var ctx = getClientContext(email);
    if (!ctx || !ctx.sheetId) return null;
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(AI_HISTORY_TAB);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 4).getValues();
    var emailNorm = (email || "").toString().trim().toLowerCase();
    var rows = data.filter(function (row) { return String(row[1] || "").trim().toLowerCase() === emailNorm; });

    rows.sort(function (a, b) {
      var tA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
      var tB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
      return tB - tA;
    });

    var entries = rows.slice(0, 15);
    var ensureBlock = function (b) {
      var sum = "";
      var arr = [];
      if (b && typeof b === "object") { sum = (b.summary || b.resumo || b.text || "").toString().trim(); arr = Array.isArray(b.todos) ? b.todos : (Array.isArray(b.actions) ? b.actions : []); }
      else if (b && typeof b === "string") { sum = String(b).trim(); }
      sum = _nuclearStripLabels(sum);
      arr = arr.map(function (t) { return _nuclearStripLabels(String(t || "")); }).filter(Boolean);
      if (!sum && arr.length === 0) return null;
      return { summary: sum || "Erro na análise.", todos: arr };
    };

    var merged = { financeiro: null, stocks: null, rh: null, fleet: null };
    for (var i = 0; i < entries.length; i++) {
      var row = entries[i];
      var raw = row[3];
      var obj = null;
      try { if (typeof raw === "string" && raw.trim()) obj = JSON.parse(raw); else if (raw && typeof raw === "object") obj = raw; } catch (e) { }
      if (!obj || typeof obj !== "object") continue;
      if (obj._raw && typeof obj._raw === "string") { try { var p = JSON.parse(obj._raw); if (p) obj = p; } catch (e) { } }
      if (!merged.financeiro && obj.financeiro) merged.financeiro = ensureBlock(obj.financeiro);
      if (!merged.stocks && obj.stocks) merged.stocks = ensureBlock(obj.stocks);
      if (!merged.rh && obj.rh) merged.rh = ensureBlock(obj.rh);
      if (!merged.fleet && obj.fleet) merged.fleet = ensureBlock(obj.fleet);
      if (merged.financeiro && merged.stocks && merged.rh && merged.fleet) break;
    }

    var hasAny = merged.financeiro || merged.stocks || merged.rh || merged.fleet;
    if (!hasAny) return null;

    return { success: true, financeiro: merged.financeiro || { summary: "", todos: [] }, stocks: merged.stocks || { summary: "", todos: [] }, rh: merged.rh || { summary: "", todos: [] }, fleet: merged.fleet || { summary: "", todos: [] }, currentCredits: getAiCredits(email) };
  } catch (e) { return null; }
}

function getAIAutoPreference(email) {
  if (!email) return true;
  try { return PropertiesService.getScriptProperties().getProperty("aiAuto_" + String(email).trim().toLowerCase()) !== "0"; } catch (e) { return true; }
}

function setAIAutoPreference(email, value) {
  if (!email) return;
  try { PropertiesService.getScriptProperties().setProperty("aiAuto_" + String(email).trim().toLowerCase(), value ? "1" : "0"); } catch (e) { }
}

function getFlowlyAIInsight(res, impersonateTarget, category) {
  var validCategories = ["financeiro", "stocks", "rh", "fleet"];
  if (category && validCategories.indexOf(category) !== -1) {
    return getFlowlyAIInsightCategory(impersonateTarget, category, res);
  }
  const fallback = { financeiro: { summary: "Chave Gemini não configurada. Configure no Perfil.", todos: [] }, stocks: { summary: "Insight indisponível.", todos: [] }, rh: { summary: "Insight indisponível.", todos: [] } };
  try {
    if (getAiCredits(impersonateTarget) <= 0) return { success: false, error: "Atingiu o limite de análises IA do seu plano. Faça upgrade ou adquira um Pack Booster.", currentCredits: getAiCredits(impersonateTarget) };
    const key = getGeminiKey();
    if (!key || !key.trim()) return { success: false, error: "Chave Gemini não configurada. Configure no Perfil.", currentCredits: getAiCredits(impersonateTarget) };
    if (!res || (res.fat != null && res.fat === 0)) {
      return { success: false, error: "Aguarda as primeiras vendas para uma análise estratégica completa.", currentCredits: getAiCredits(impersonateTarget) };
    }

    var ctx = {};
    try { ctx = getClientContext(impersonateTarget || null); } catch (e) { }
    var cacheKey = "aiInsight_" + (ctx.sheetId || "def") + "_" + (res.fat || 0) + "_" + (res.stock && res.stock.valor || 0) + "_" + (res.rh && res.rh.ativos || 0);
    var cache = CacheService.getDocumentCache();
    var props = PropertiesService.getDocumentProperties();
    var cooldownUntil = parseInt(props.getProperty("aiInsight_429_until") || "0", 10);

    if (Date.now() < cooldownUntil) {
      var cached = cache.get(cacheKey);
      if (cached) { try { var parsed = JSON.parse(cached); if (parsed && parsed.success === true) { parsed.currentCredits = getAiCredits(impersonateTarget); return parsed; } } catch (e) { } }
      return { success: false, error: "Limite de velocidade da API atingido. Por favor, aguarde alguns minutos e tente novamente. Obrigado pela sua compreensão.", currentCredits: getAiCredits(impersonateTarget) };
    }

    var cached = cache.get(cacheKey);
    if (cached) { try { var parsed = JSON.parse(cached); if (parsed && parsed.success === true) { parsed.currentCredits = getAiCredits(impersonateTarget); return parsed; } } catch (e) { } }

    const fin = res.financeiro || {};
    const be = res.breakEven || {};
    const taxa = res.taxaAbsorcao || {};
    const st = res.stock || {};
    const rhData = res.rh || {};
    const dataAtual = new Date().toLocaleDateString("pt-PT");
    const deadStockVal = res.deadStockValue != null ? String(res.deadStockValue) : "0";
    const erosionCnt = res.erosionCount != null ? String(res.erosionCount) : "0";
    const deadList = (res.deadStockList || []).slice(0, 10);
    const erosionList = (res.erosionItems || []).slice(0, 10);

    const margemVal = (fin.margem_perc != null && fin.margem_perc !== "") ? String(fin.margem_perc) : "0";
    const fatVal = (fin.faturacao != null && fin.faturacao !== "") ? String(fin.faturacao) : "0";
    const stockVal = (st.valor != null) ? String(st.valor) : "0";
    const diasCobertura = (st.diasCobertura != null) ? String(st.diasCobertura) : "N/A";
    const rotatividade = (st.rotatividade != null) ? String(st.rotatividade) : "N/A";
    const ativos = rhData.ativos != null ? String(rhData.ativos) : "0";
    const ideal = rhData.ideal != null ? String(rhData.ideal) : "N/A";
    const custoMedio = (rhData.custo_medio != null) ? String(rhData.custo_medio) : "N/A";
    const custoMensal = (rhData.custo_mensal != null) ? String(rhData.custo_mensal) : "0";

    const caixaLivreVal = (fin.caixaLivre != null) ? String(fin.caixaLivre) : "0";
    const fatAPagVal = (fin.faturasAPagamento != null) ? String(fin.faturasAPagamento) : "0";
    const saldoBancVal = (fin.saldoBancario != null) ? String(fin.saldoBancario) : "0";
    const compRHVal = (fin.compromissosRH != null) ? String(fin.compromissosRH) : "0";
    const financeSummaryText = "REGRA: Saídas=Vendas (entrada de €). Entradas=Stock (saída de €). Saldo Bancário deduz custos RH e faturas pagas. Caixa Livre (" + caixaLivreVal + " €) = Entradas PAGAS - Saídas PAGAS (faturas a pagar " + fatAPagVal + " € são apenas para visualização, NÃO subtraídas). Se Caixa Livre > 0: elogia. Se < 0: alerta fundo de maneio.";

    const systemPrompt = "Flowly 360. Responde APENAS com um OBJETO JSON VÁLIDO, SEM MARKDOWN E SEM CRASES. STRICT JSON OBJECT ONLY, NO MARKDOWN. Estrutura exata: { \"financeiro\": { \"summary\": \"...\", \"todos\": [\"...\"] }, \"stocks\": { \"summary\": \"...\", \"todos\": [\"...\"] }, \"rh\": { \"summary\": \"...\", \"todos\": [\"...\"] } }. Summary: máx 3 frases. Todos: 2-4 ações. FINANCEIRO: CFO, Break-even, Margem, Taxa Absorção, Caixa Livre (ver regra abaixo). STOCKS: Stock morto, Dias cobertura, Rotatividade. RH: Equipa ideal vs ativos, custo_mensal total. Usa só os números fornecidos. Custo RH = custo_mensal.";

    const userPrompt = "Data: " + dataAtual + "\n\n" +
      "DADOS FINANCEIROS: Receita " + fatVal + " €, Margem " + margemVal + "%, Saldo Bancário " + saldoBancVal + " €, Faturas a Pagar " + fatAPagVal + " €, Compromissos RH " + compRHVal + " €, Caixa Livre (Liquidez Imediata) " + caixaLivreVal + " €. Taxa Absorção " + (taxa.rate != null ? taxa.rate + "%" : "N/A") + (taxa.label ? " (" + taxa.label + ")" : "") + ". " +
      financeSummaryText + "\n" +
      "Break-even: target " + (be.target != null ? be.target.toFixed(2) : "0") + " €, projeção " + (be.projectedRev != null ? be.projectedRev.toFixed(2) : "0") + " €, gap " + (be.gap != null ? be.gap.toFixed(2) : "0") + " €, atingido: " + (be.reached ? "sim" : "não") + ".\n" +
      "DADOS STOCKS: Valor Stock " + stockVal + " €, Rotatividade " + rotatividade + ", Dias Cobertura " + diasCobertura + ", Dead Stock " + deadStockVal + " €, Erosão " + erosionCnt + ". " +
      (deadList.length ? "Stock Morto: " + (deadList.map(function (d) { return d.artigo || d.art; }).join(", ")) + ". " : "") +
      (erosionList.length ? "Erosão: " + (erosionList.map(function (e) { return e.artigo; }).join(", ")) + ". " : "") + "\n" +
      "DADOS RH: Ativos " + ativos + ", Equipa Ideal " + ideal + ", Custo Médio " + custoMedio + " €, Custo Mensal Total " + custoMensal + " €.\n\n" +
      "Responde APENAS com o objeto JSON válido. STRICT JSON ONLY, NO MARKDOWN, NO BACKTICKS:";

    const models = ["gemini-3.1-flash", "gemini-3.0-flash", "gemini-3.0-flash-preview", "gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"];
    let content = null;
    var lastEx = "";

    for (var i = 0; i < models.length; i++) {
      try {
        const url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent?key=" + key;
        const payload = { contents: [{ parts: [{ text: systemPrompt + "\n\n" + userPrompt }] }], generationConfig: { maxOutputTokens: 1024, temperature: 0.3 } };
        var resp = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
        var respCode = resp.getResponseCode();
        var respText = resp.getContentText();

        if (respCode === 429) {
          props.setProperty("aiInsight_429_until", String(Date.now() + 300000));
          var c2 = cache.get(cacheKey);
          if (c2) { try { var p = JSON.parse(c2); if (p && p.success === true) return p; } catch (e) { } }
          return { success: false, error: "Limite de velocidade da API atingido. Por favor, aguarde alguns minutos e tente novamente.", currentCredits: getAiCredits(impersonateTarget) };
        }

        if (respCode >= 400) {
          lastEx = "API " + respCode;
          try { var errBody = JSON.parse(respText); if (errBody.error && errBody.error.message) lastEx = errBody.error.message; } catch (_) { }
          Utilities.sleep(2000);
          continue;
        }

        content = JSON.parse(respText);
        if (!content.error) break;

        lastEx = content.error && content.error.message ? content.error.message : respCode + " " + (respText || "").substring(0, 100);
        Utilities.sleep(2000);
      } catch (err) { lastEx = (err && err.message ? err.message : String(err)).substring(0, 120); }
    }

    if (!content || content.error) {
      var errMsg = (content && content.error) ? (content.error.message || content.error.status || JSON.stringify(content.error)) : (lastEx || "Sem resposta da API");
      var shortErr = (errMsg && errMsg.length < 80) ? errMsg : (lastEx ? lastEx : "Erro na API");
      return { success: false, error: "Não foi possível gerar o insight. " + (shortErr || "Verifica a chave Gemini."), currentCredits: getAiCredits(impersonateTarget) };
    }

    const cand = content.candidates && content.candidates[0];
    if (!cand) return { success: false, error: "Resposta da API sem conteúdo.", currentCredits: getAiCredits(impersonateTarget) };

    let text = cand.content && cand.content.parts ? (cand.content.parts || []).map(function (p) { return p.text || ""; }).join("").trim() : "";
    if (!text) return { success: false, error: "Resposta da API vazia.", currentCredits: getAiCredits(impersonateTarget) };

    var userEmail = (impersonateTarget != null && impersonateTarget !== "") ? (function () { try { return getClientContext(impersonateTarget).clientEmail || impersonateTarget; } catch (e) { return impersonateTarget; } })() : (Session.getActiveUser().getEmail() || "");
    var newCredits = deductCredits(userEmail, 10);
    if (newCredits < 0) consumeAiCredit(impersonateTarget);

    text = cleanResponse(text) || text;
    let cleaned = text.trim();
    const firstBrace = cleaned.indexOf("{");
    if (firstBrace >= 0) {
      let depth = 0, end = -1;
      for (var j = firstBrace; j < cleaned.length; j++) {
        if (cleaned.charAt(j) === "{") depth++;
        else if (cleaned.charAt(j) === "}") { depth--; if (depth === 0) { end = j; break; } }
      }
      if (end >= 0) cleaned = cleaned.substring(firstBrace, end + 1);
    }

    try {
      let parsed;
      try { parsed = JSON.parse(cleaned); } catch (e1) {
        var fixed = cleaned.replace(/,(\s*[}\]])/g, "$1").replace(/[\u201C\u201D]/g, '"').replace(/[\u2018\u2019]/g, "'");
        try { parsed = JSON.parse(fixed); } catch (e2) { fixed = fixed.replace(/\r?\n/g, " ").replace(/\t/g, " "); parsed = JSON.parse(fixed); }
      }

      const ensureBlock = function (b) {
        if (!b || typeof b !== "object") return { summary: "", todos: [] };
        const sum = (b.summary || b.resumo || b.text || "").toString().trim();
        const arr = Array.isArray(b.todos) ? b.todos : (Array.isArray(b.actions) ? b.actions : []);
        return { summary: sum || "Erro na análise.", todos: arr.map(function (t) { return String(t || "").trim(); }).filter(Boolean) };
      };

      const finBlock = parsed.financeiro || parsed.financial;
      const stkBlock = parsed.stocks || parsed.stock;
      const rhBlock = parsed.rh || parsed.hr;

      var result = { success: true, financeiro: ensureBlock(finBlock), stocks: ensureBlock(stkBlock), rh: ensureBlock(rhBlock), currentCredits: getAiCredits(impersonateTarget) };

      try { cache.put(cacheKey, JSON.stringify(result), 7200); } catch (e) { }

      if (ctx && ctx.sheetId && userEmail) saveAIHistoryEntry(ctx.sheetId, userEmail, "dashboard", result);

      try {
        ensureMasterDBColumns();
        var ssMaster = SpreadsheetApp.openById(MASTER_DB_ID);
        var sheetMaster = ssMaster.getSheets()[0];
        var dataMaster = sheetMaster.getDataRange().getValues();
        var emailNorm = (userEmail || "").toString().trim().toLowerCase();
        for (var i = 1; i < dataMaster.length; i++) {
          if (String(dataMaster[i][MASTER_COL_EMAIL - 1] || "").trim().toLowerCase() === emailNorm) {
            sheetMaster.getRange(i + 1, MASTER_COL_LAST_AI_INSIGHT).setValue(JSON.stringify(result));
            break;
          }
        }
      } catch (saveDbErr) { }

      return result;
    } catch (parseErr) {
      var rawSummary = (cleaned || text || "").substring(0, 2000);
      var fallbackResult = { success: true, financeiro: { summary: rawSummary || "Resposta recebida mas em formato inesperado.", todos: [] }, stocks: { summary: "Insight indisponível.", todos: [] }, rh: { summary: "Insight indisponível.", todos: [] }, currentCredits: getAiCredits(impersonateTarget) };
      if (ctx && ctx.sheetId && userEmail) saveAIHistoryEntry(ctx.sheetId, userEmail, "dashboard", fallbackResult);
      return fallbackResult;
    }
  } catch (e) { return { success: false, error: (e && e.message) ? e.message : "Erro inesperado na análise IA.", currentCredits: getAiCredits(impersonateTarget) }; }
}

function getFlowlyAIInsightCategory(impersonateTarget, category, dashRes) {
  var valid = ["financeiro", "stocks", "rh", "fleet"];
  if (!category || valid.indexOf(category) === -1) return { success: false, error: "Categoria inválida.", currentCredits: getAiCredits(impersonateTarget) };
  if (getAiCredits(impersonateTarget) < 4) return { success: false, error: "Créditos insuficientes (necessário 4).", currentCredits: getAiCredits(impersonateTarget) };
  var key = getGeminiKey();
  if (!key || !key.trim()) return { success: false, error: "Chave Gemini não configurada.", currentCredits: getAiCredits(impersonateTarget) };

  var res = dashRes;
  if (!res || !res.success) {
    try { res = getDashboardData(impersonateTarget, null, null, null, null, false); } catch (e) { return { success: false, error: "Dados do dashboard indisponíveis.", currentCredits: getAiCredits(impersonateTarget) }; }
  }
  if (!res || !res.success) return { success: false, error: "Dados do dashboard indisponíveis.", currentCredits: getAiCredits(impersonateTarget) };

  var userEmail = (impersonateTarget != null && impersonateTarget !== "") ? (function () { try { return getClientContext(impersonateTarget).clientEmail || impersonateTarget; } catch (e) { return impersonateTarget; } })() : (Session.getActiveUser().getEmail() || "");

  var fin = res.financeiro || {};
  var be = res.breakEven || {};
  var taxa = res.taxaAbsorcao || {};
  var st = res.stock || {};
  var rhData = res.rh || {};
  var dataAtual = new Date().toLocaleDateString("pt-PT");
  var systemPrompt = "";
  var userPrompt = "";

  if (category === "financeiro") {
    var caixaLivreVal = (fin.caixaLivre != null) ? String(fin.caixaLivre) : "0";
    var fatAPagVal = (fin.faturasAPagamento != null) ? String(fin.faturasAPagamento) : "0";
    var financeSummaryText = "REGRA: Caixa Livre (" + caixaLivreVal + " €) = Entradas PAGAS - Saídas PAGAS. Faturas a pagar são apenas para visualização. Se Caixa Livre > 0: elogia. Se < 0: alerta fundo de maneio.";
    var finComplete = Object.assign({}, fin, { breakEven: be, taxaAbsorcao: taxa, _financeSummaryText: financeSummaryText });
    systemPrompt = "Flowly 360. Responde APENAS com um OBJETO JSON com uma única chave. STRICT JSON: { \"financeiro\": { \"summary\": \"...\", \"todos\":[\"...\"] } }. Summary: máx 3 frases. Todos: 2-4 ações. Analisa os dados financeiros fornecidos. SEM MARKDOWN.";
    userPrompt = "Data: " + dataAtual + "\n\n" + financeSummaryText + "\n\nDADOS FINANCEIROS COMPLETOS:\n" + JSON.stringify(finComplete, null, 2);
  } else if (category === "stocks") {
    var stComplete = Object.assign({}, st, { deadStockValue: res.deadStockValue, deadStockList: res.deadStockList || [], erosionCount: res.erosionCount, erosionItems: res.erosionItems || [] });
    systemPrompt = "Flowly 360. Responde APENAS com um OBJETO JSON com uma única chave. STRICT JSON: { \"stocks\": { \"summary\": \"...\", \"todos\": [\"...\"] } }. Summary: máx 3 frases. Todos: 2-4 ações. Analisa os dados de stocks fornecidos, focando em stock morto, dias de cobertura e rotatividade. SEM MARKDOWN.";
    userPrompt = "Data: " + dataAtual + "\n\nDADOS DE STOCKS COMPLETOS:\n" + JSON.stringify(stComplete, null, 2);
  } else if (category === "fleet") {
    var frotaData = (res.frota || getDashboardFrotaData(impersonateTarget));
    systemPrompt = "Flowly 360. Responde APENAS com um OBJETO JSON com uma única chave. STRICT JSON: { \"fleet\": { \"summary\": \"...\", \"todos\": [\"...\"] } }. Summary: máx 3 frases. Todos: 2-4 ações. Analisa os dados da frota fornecidos. Sugere melhorias na condução, rotas ou manutenção preventiva com base nos dados. SEM MARKDOWN.";
    userPrompt = "Data: " + dataAtual + "\n\nDADOS DA FROTA:\n" + JSON.stringify(frotaData, null, 2);
  } else {
    systemPrompt = "Flowly 360. Responde APENAS com um OBJETO JSON com uma única chave. STRICT JSON: { \"rh\": { \"summary\": \"...\", \"todos\": [\"...\"] } }. Summary: máx 3 frases. Todos: 2-4 ações. Analisa os dados de RH fornecidos, focando em equipa ideal vs ativos e custo_mensal total. SEM MARKDOWN.";
    userPrompt = "Data: " + dataAtual + "\n\nDADOS DE RH COMPLETOS:\n" + JSON.stringify(rhData, null, 2);
  }

  var models = ["gemini-3.1-flash", "gemini-3.0-flash", "gemini-3.0-flash-preview", "gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"];
  var content = null;
  var lastEx = "";

  for (var i = 0; i < models.length; i++) {
    try {
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent?key=" + key;
      var payload = { contents: [{ parts: [{ text: systemPrompt + "\n\n" + userPrompt }] }], generationConfig: { maxOutputTokens: 2048, temperature: 0.3, responseMimeType: "application/json" } };
      var resp = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
      var respCode = resp.getResponseCode();
      var respText = resp.getContentText();

      if (respCode === 429) return { success: false, error: "Limite de velocidade da API atingido.", currentCredits: getAiCredits(impersonateTarget) };
      if (respCode >= 400) { lastEx = "API " + respCode; Utilities.sleep(2000); continue; }

      content = JSON.parse(respText);
      if (!content.error) break;

      lastEx = content.error && content.error.message ? content.error.message : respCode + "";
      Utilities.sleep(2000);
    } catch (err) { lastEx = (err && err.message) ? err.message : String(err); }
  }

  if (!content || content.error) return { success: false, error: "Não foi possível gerar. " + (lastEx || "Verifica a chave Gemini."), currentCredits: getAiCredits(impersonateTarget) };

  var cand = content.candidates && content.candidates[0];
  if (!cand) return { success: false, error: "Resposta da API sem conteúdo.", currentCredits: getAiCredits(impersonateTarget) };

  var text = (cand.content && cand.content.parts ? (cand.content.parts || []).map(function (p) { return p.text || ""; }).join("") : "").trim();
  if (!text) return { success: false, error: "Resposta da API vazia.", currentCredits: getAiCredits(impersonateTarget) };

  text = cleanResponse(text) || text;
  text = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  var firstBrace = text.indexOf("{");
  if (firstBrace >= 0) {
    var depth = 0, end = -1;
    for (var j = firstBrace; j < text.length; j++) {
      if (text.charAt(j) === "{") depth++;
      else if (text.charAt(j) === "}") { depth--; if (depth === 0) { end = j; break; } }
    }
    if (end >= 0) text = text.substring(firstBrace, end + 1);
  }

  var ensureBlock = function (b) {
    if (!b || typeof b !== "object") return { summary: "", todos: [] };
    var sum = (b.summary || b.resumo || b.text || "").toString().trim();
    var arr = Array.isArray(b.todos) ? b.todos : (Array.isArray(b.actions) ? b.actions : []);
    return { summary: sum || "Erro na análise.", todos: arr.map(function (t) { return String(t || "").trim(); }).filter(Boolean) };
  };

  var parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e1) {
    // Fallback de segurança (muito útil na Frota devido às sugestões descritivas gerarem quebras de linha ou aspas)
    var fixed = text.replace(/,(\s*[}\]])/g, "$1").replace(/[\u201C\u201D]/g, '"').replace(/[\u2018\u2019]/g, "'");
    try {
      parsed = JSON.parse(fixed);
    } catch (e2) {
      fixed = fixed.replace(/\r?\n/g, " ").replace(/\t/g, " ");
      try {
        parsed = JSON.parse(fixed);
      } catch (e3) {
        return { success: false, error: "Resposta da IA mal formatada. Tente novamente.", currentCredits: getAiCredits(impersonateTarget) };
      }
    }
  }

  var newCredits = deductCredits(userEmail, 4);
  if (newCredits < 0) return { success: false, error: "Créditos insuficientes.", currentCredits: getAiCredits(impersonateTarget) };

  var rootBlock = (parsed && typeof parsed === "object" && parsed.summary != null) ? parsed : null;
  var newBlock = rootBlock ? ensureBlock(rootBlock) : (category === "financeiro" ? ensureBlock(parsed.financeiro || parsed.financial) : (category === "stocks" ? ensureBlock(parsed.stocks || parsed.stock) : (category === "fleet" ? ensureBlock(parsed.fleet || parsed.frota) : ensureBlock(parsed.rh || parsed.hr))));
  var merged = { success: true, financeiro: { summary: "Insight indisponível.", todos: [] }, stocks: { summary: "Insight indisponível.", todos: [] }, rh: { summary: "Insight indisponível.", todos: [] }, fleet: { summary: "Insight indisponível.", todos: [] } };

  try {
    ensureMasterDBColumns();
    var ssMaster = SpreadsheetApp.openById(MASTER_DB_ID);
    var sheetMaster = ssMaster.getSheets()[0];
    var dataMaster = sheetMaster.getDataRange().getValues();
    var emailNorm = (userEmail || "").toString().trim().toLowerCase();
    for (var i = 1; i < dataMaster.length; i++) {
      if (String(dataMaster[i][MASTER_COL_EMAIL - 1] || "").trim().toLowerCase() === emailNorm) {
        var stored = sheetMaster.getRange(i + 1, MASTER_COL_LAST_AI_INSIGHT).getValue();
        if (stored && String(stored).trim()) {
          try {
            var existing = JSON.parse(String(stored));
            if (existing && typeof existing === "object") {
              if (existing.financeiro) merged.financeiro = ensureBlock(existing.financeiro);
              if (existing.stocks) merged.stocks = ensureBlock(existing.stocks);
              if (existing.rh) merged.rh = ensureBlock(existing.rh);
              if (existing.fleet) merged.fleet = ensureBlock(existing.fleet);
            }
          } catch (e) { }
        }
        merged[category] = newBlock;
        merged.currentCredits = getAiCredits(impersonateTarget);
        sheetMaster.getRange(i + 1, MASTER_COL_LAST_AI_INSIGHT).setValue(JSON.stringify(merged));
        var ctx = {};
        try { ctx = getClientContext(impersonateTarget || null); } catch (e) { }
        if (ctx && ctx.sheetId && userEmail) saveAIHistoryEntry(ctx.sheetId, userEmail, category, merged);
        return merged;
      }
    }
  } catch (e) { }

  merged[category] = newBlock;
  merged.currentCredits = getAiCredits(impersonateTarget);
  var ctx = {};
  try { ctx = getClientContext(impersonateTarget || null); } catch (e) { }
  if (ctx && ctx.sheetId && userEmail) saveAIHistoryEntry(ctx.sheetId, userEmail, category, merged);
  return merged;
}

function getFinancialSummaryForAI(impersonateTarget) {
  try {
    const dash = getDashboardData(impersonateTarget, null, null, null, null);
    if (!dash || !dash.success) return null;

    const fin = dash.financeiro || {};
    const st = dash.stock || {};
    const rh = dash.rh || {};
    const ivaT = dash.alertaIvaTrimestral || {};
    const master = getMasterData(impersonateTarget);
    const ccList = (master && master.cc) ? master.cc : [];
    const staffList = (master && master.staff) ? master.staff : [];
    const usersList = (master && master.users) ? master.users : [];

    const topGastos = ccList.filter(x => x.tipo === "entrada").sort((a, b) => parseFloat(b.total) - parseFloat(a.total)).slice(0, 15);
    const topVendas = ccList.filter(x => x.tipo === "saida" || x.tipo === "fechocaixa").sort((a, b) => parseFloat(b.total) - parseFloat(a.total)).slice(0, 15);
    const faturasAbertas = ccList.filter(x => x.estado === "Aberto");
    const faturasPagas = ccList.filter(x => x.estado === "Pago");

    const alerts = getAIAlerts(impersonateTarget, dash);

    const staffResumo = staffList.map(s => ({
      nome: s.nome, cargo: s.cargo, estado: s.status, email: s.email,
      vencimento: s.vencimento, subAlim: s.subAlim, seguro: s.seguro,
      tsu: s.tsuPct,
      custoMensalReal: s.custoMensalReal,
      custoBase: s.custoBase || s.custoMensalReal,
      fatorProporcional: s.fatorProporcional != null ? s.fatorProporcional : 100,
      dataSaida: s.dataSaida || "",
      provFerias: s.provFerias, provNatal: s.provNatal,
      provRescisao: s.provRescisao, provFormacao: s.provFormacao,
      premios: s.premios, admissao: s.admissao, diasContrato: s.diasContrato
    }));

    const caixaLivreVal = (fin.caixaLivre != null) ? String(fin.caixaLivre) : "0";
    const fatAPagVal = (fin.faturasAPagamento != null) ? String(fin.faturasAPagamento) : "0";
    const financeSummaryText = "REGRA: Caixa Livre (" + caixaLivreVal + " €) = Entradas PAGAS - Saídas PAGAS. Faturas a pagar (" + fatAPagVal + " €) são apenas para visualização. Se Caixa Livre > 0: contas cobertas. Se < 0: alerta fundo de maneio.";

    return {
      kpis: {
        faturacao: fin.faturacao, lucro: fin.lucro_liq, margem_perc: fin.margem_perc,
        despesas: fin.despesas, iva_pagar: fin.iva_pagar, irc_est: fin.irc_est,
        ticket_medio: fin.ticket_medio, transacoes: fin.transacoes,
        saldoBancario: fin.saldoBancario, caixaLivre: fin.caixaLivre, faturasAPagamento: fin.faturasAPagamento,
        despesasAbertoMes: fin.despesasAbertoMes, compromissosRH: fin.compromissosRH
      },
      _financeSummaryText: financeSummaryText,
      stock: { valor: st.valor, rotatividade: st.rotatividade, compras: st.compras },
      iva_trimestral: {
        T1: ivaT.T1, T2: ivaT.T2, T3: ivaT.T3, T4: ivaT.T4,
        trimestre_atual: ivaT.trimestreAtual, valor_atual: ivaT.valorTrimestreAtual
      },
      rh: {
        custo_mensal: (rh.custo_mensal != null) ? String(rh.custo_mensal) : null,
        custo_total: rh.custo_total, ativos: rh.ativos,
        custo_medio: rh.custo_medio, colaboradores: staffResumo
      },
      tesouraria: {
        total_faturas_abertas: faturasAbertas.length,
        valor_total_aberto: faturasAbertas.reduce((s, x) => s + (parseFloat(x.totalPendente) || 0), 0).toFixed(2),
        total_faturas_pagas: faturasPagas.length,
        top_gastos: topGastos.map(g => ({ fornecedor: g.fornecedor, total: g.total, data: g.data })),
        top_vendas: topVendas.map(v => ({ total: v.total, fornecedor: v.fornecedor, data: v.data }))
      },
      alertas_ativos: (alerts || []).map(a => a.message),
      utilizadores_com_acesso: usersList.map(u => ({ email: u.email, nome: u.name, perfil: u.role, estado: u.status }))
    };
  } catch (e) { return null; }
}

function chatWithCFO(userMsg, context, impersonateTarget) {
  try {
    var target = impersonateTarget != null && impersonateTarget !== "" ? impersonateTarget : (context && context.impersonateEmail) ? context.impersonateEmail : Session.getActiveUser().getEmail();
    if (getAiCredits(target) <= 0) return { success: false, resposta: "⚠️ Atingiu o limite de análises do seu plano. Faça upgrade para continuar a nossa conversa." };
    const key = getGeminiKey();
    if (!key || !key.trim()) return { success: false, resposta: "Chave Gemini não configurada. Configure no Perfil." };

    var summary = getFinancialSummaryForAI(target);
    if (!summary) {
      var ctx = context;
      if (!ctx || typeof ctx !== "object") {
        ctx = getDashboardData(target, null, null, null, null);
        if (!ctx || !ctx.success) return { success: false, resposta: "Não foi possível obter os dados do dashboard. Carregue o dashboard primeiro." };
        summary = getFinancialSummaryForAI(target);
      }
      if (!summary) summary = (ctx && typeof ctx === "object") ? { kpis: ctx.financeiro, stock: ctx.stock, rh: ctx.rh } : null;
    }
    if (!summary) return { success: false, resposta: "Dados indisponíveis. Carregue o dashboard primeiro." };

    var contextStr = JSON.stringify(summary, null, 0);
    if (contextStr.length > 12000) contextStr = contextStr.substring(0, 12000) + '..." (resumo truncado)';

    const systemPrompt = "És o consultor financeiro do Flowly 360. Usa os dados fornecidos no contexto para responder. Sê breve, estratégico e profissional. Responde em português de Portugal.";
    const fullPrompt = systemPrompt + "\n\nCONTEXTO (resumo):\n" + contextStr + "\n\nPergunta: " + (userMsg || "");
    const models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-2.5-flash", "gemini-1.5-pro"];

    let content = null;
    var lastErr = "";
    for (var i = 0; i < models.length; i++) {
      try {
        const url = "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent?key=" + key;
        const resp = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            contents: [{ parts: [{ text: fullPrompt }] }],
            generationConfig: { maxOutputTokens: 2048, temperature: 0.4 }
          }),
          muteHttpExceptions: true
        });
        var code = resp.getResponseCode();
        var respText = resp.getContentText();
        if (code === 429) return { success: false, resposta: "Quota exceeded. Tente mais tarde." };
        if (code >= 400) {
          try { var eb = JSON.parse(respText); if (eb.error && eb.error.message) lastErr = eb.error.message; } catch (_) { lastErr = "API " + code; }
          continue;
        }
        content = JSON.parse(respText);
        if (!content.error) break;
        lastErr = content.error.message || "API error";
      } catch (err) { lastErr = (err && err.message) ? err.message : String(err); }
    }

    if (!content || content.error) return { success: false, resposta: lastErr ? ("Erro: " + lastErr) : "Não foi possível obter resposta. Verifica a chave Gemini." };
    const text = content.candidates && content.candidates[0] && content.candidates[0].content
      ? (content.candidates[0].content.parts || []).map(function (p) { return p.text || ""; }).join("").trim()
      : "";
    consumeAiCredit(target);
    return { success: true, resposta: text || "Sem resposta." };
  } catch (e) {
    return { success: false, resposta: "Erro: " + (e.toString() || "Erro desconhecido.") };
  }
}

function getAIChatResponse(pergunta, impersonateTarget) {
  try {
    if (!getGeminiKey()) return { success: false, error: MSG_IA_PENDENTE };
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião para selecionar um cliente." };

    const summary = getFinancialSummaryForAI(impersonateTarget);
    if (!summary) return { success: false, error: "Não foi possível obter os dados financeiros." };

    const contextText = JSON.stringify(summary, null, 0);
    const prompt = `És o assistente financeiro e de gestão da Flowly 360, uma plataforma de gestão empresarial.
Tens acesso COMPLETO aos dados reais da empresa do utilizador. Analisa TUDO antes de responder.

DADOS COMPLETOS DA EMPRESA (JSON):
${contextText}

INSTRUÇÕES:
- Responde SEMPRE em português de Portugal.
- REGRA ABSOLUTA: Para custos de RH, usa apenas o valor "rh.custo_mensal" (Total Monthly Cost).
- Se existirem ALERTAS ativos, MENCIONA-OS proativamente na resposta.
- Usa dados de "rh.colaboradores" quando aplicável.
- Se a pergunta envolver tesouraria: usa faturas abertas, pendentes. Caixa Livre = Entradas PAGAS - Saídas PAGAS.
- Apresenta valores em € com 2 casas decimais.
- Se identificares situações críticas (margens negativas, faturas vencidas, contratos a expirar), alerta o utilizador.
- Sê específico: usa nomes de colaboradores, fornecedores, artigos e valores concretos.

Pergunta do utilizador: "${pergunta}"

Resposta (detalhada e completa, com dados concretos, tabelas quando útil, sem limitar informação):`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + getGeminiKey();
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 4096, temperature: 0.3 }
      }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var respText = res.getContentText();

    if (code === 429) return { success: false, error: "Quota exceeded" };
    if (code >= 400) {
      try { var errBody = JSON.parse(respText); if (errBody.error && errBody.error.message) return { success: false, error: errBody.error.message }; } catch (_) { }
      return { success: false, error: "API error " + code };
    }

    const content = JSON.parse(respText);
    if (content.error) return { success: false, error: (content.error.message || MSG_IA_PENDENTE) };
    const text = content.candidates && content.candidates[0] && content.candidates[0].content
      ? (content.candidates[0].content.parts || []).map(p => p.text || "").join("").trim()
      : "";

    return { success: true, resposta: text || "Não foi possível gerar uma resposta." };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    if (errMsg.indexOf("quota") >= 0 || errMsg.indexOf("429") >= 0) return { success: false, error: "Quota exceeded" };
    return { success: false, error: MSG_IA_PENDENTE };
  }
}

function getAIAlerts(impersonateTarget, dashRes) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return [];

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const dash = dashRes && dashRes.success ? dashRes : getDashboardData(impersonateTarget, null, null, null, null);
    const alerts = [];
    const now = new Date();
    const currentQuarter = now.getMonth() <= 2 ? "T1" : (now.getMonth() <= 5 ? "T2" : (now.getMonth() <= 8 ? "T3" : "T4"));
    const yearKey = now.getFullYear() + "-" + currentQuarter;

    const sheetLog = ss.getSheetByName(SHEET_TAB_NAME);
    let ivaTrimestre = 0;
    const margensBaixas = [];
    const margensNegativas = [];
    const setAberto = new Set();
    const setVencidas = new Set();
    let totalVencido = 0;
    let totalAberto = 0;
    let maiorDevedor = { nome: "", valor: 0 };
    const devedores = {};
    let comprasMes = 0, vendasMes = 0;

    if (sheetLog && sheetLog.getLastRow() > 1) {
      const dataLog = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 21).getValues();
      dataLog.forEach(function (r, i) {
        const docId = (r[20] && String(r[20]).trim() !== "") ? String(r[20]).trim() : (String(r[0]) + "_" + String(r[3] || ""));
        const tipo = normalizeTipo(r[1]);
        const contaStock = (r[19] != null && String(r[19]).trim() !== "") ? String(r[19]).trim() : "Sim";
        if (contaStock === "Não") return;

        const qty = parseFloat(r[5]) || 0;
        const cost = parseFloat(r[6]) || 0;
        const sell = parseFloat(r[7]) || 0;
        const artigo = (r[4] || "").toString().trim();
        const fornecedor = (r[3] || "").toString().trim();
        const taxaRaw = parseFloat(String(r[8] || "").replace("%", "").trim());
        const taxa = isNaN(taxaRaw) ? 0.23 : (taxaRaw > 1 ? taxaRaw / 100 : taxaRaw);
        const dedutivel = String(r[10] || "").trim().toLowerCase() === "sim";
        const valorIvaSheet = parsePTFloat(r[9]);
        const estado = (r[16] != null && r[16] !== "") ? String(r[16]).trim() : "Aberto";
        const valorPago = parseFloat(r[18]) || 0;

        let d = r[0];
        if (typeof d === "string") { const p = d.split("/"); if (p.length === 3) d = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
        if (!(d instanceof Date) || isNaN(d)) return;

        const total = (tipo === "entrada") ? (qty * cost) : (qty * sell);
        const pendente = Math.max(0, total - valorPago);

        if (estado === "Aberto" && pendente > 0) {
          setAberto.add(docId);
          totalAberto += pendente;
          const diffDays = Math.ceil((now - d) / (1000 * 60 * 60 * 24));
          if (diffDays > 30) { setVencidas.add(docId); totalVencido += pendente; }
          if (tipo === "saida" || tipo === "fechocaixa") { const key = fornecedor || "Desconhecido"; devedores[key] = (devedores[key] || 0) + pendente; }
        }

        const month = d.getMonth();
        const q = month <= 2 ? "T1" : (month <= 5 ? "T2" : (month <= 8 ? "T3" : "T4"));
        const rYearKey = d.getFullYear() + "-" + q;

        if (rYearKey === yearKey) {
          if (tipo === "saida" || tipo === "fechocaixa") ivaTrimestre += (valorIvaSheet > 0 ? valorIvaSheet : (qty * sell) * taxa);
          if (tipo === "entrada" && dedutivel) ivaTrimestre -= (valorIvaSheet > 0 ? valorIvaSheet : (qty * cost) * taxa);
          if ((tipo === "despesas" || tipo === "consumo" || tipo === "despesa") && dedutivel) ivaTrimestre -= (valorIvaSheet > 0 ? valorIvaSheet : (qty * (cost || sell)) * taxa);
        }

        if ((tipo === "saida" || tipo === "fechocaixa") && cost > 0 && sell > 0 && artigo) {
          const margem = ((sell - cost) / sell) * 100;
          if (margem < 0) margensNegativas.push({ artigo, margem: round2(margem), custo: cost, venda: sell });
          else if (margem < 15) margensBaixas.push({ artigo, margem: round2(margem), custo: cost, venda: sell });
        }

        if (d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()) {
          if (tipo === "entrada") comprasMes += total;
          else vendasMes += total;
        }
      });
    }

    const faturasVencidas = setVencidas.size;
    const countAberto = setAberto.size;

    if (faturasVencidas > 0) alerts.push({ type: "vencidas", message: "⚠️ " + faturasVencidas + " fatura(s) vencida(s) há mais de 30 dias! Total em dívida: " + round2(totalVencido).toFixed(2) + "€.", severity: "critical" });
    if (countAberto > 3) alerts.push({ type: "aberto", message: "📋 Tem " + countAberto + " documentos em aberto (total: " + round2(totalAberto).toFixed(2) + "€).", severity: "warning" });

    Object.keys(devedores).forEach(k => { if (devedores[k] > maiorDevedor.valor) maiorDevedor = { nome: k, valor: devedores[k] }; });
    if (maiorDevedor.valor > 500) alerts.push({ type: "devedor", message: "💰 Maior valor a receber: \"" + maiorDevedor.nome + "\" deve " + round2(maiorDevedor.valor).toFixed(2) + "€.", severity: "warning" });

    margensNegativas.slice(0, 3).forEach(m => { alerts.push({ type: "margem_negativa", message: "🚨 PERDA em \"" + m.artigo + "\": margem " + m.margem.toFixed(1) + "% (custo " + m.custo.toFixed(2) + "€ > venda " + m.venda.toFixed(2) + "€).", severity: "critical" }); });
    margensBaixas.slice(0, 3).forEach(m => { alerts.push({ type: "margem", message: "📉 Margem baixa em \"" + m.artigo + "\": " + m.margem.toFixed(1) + "%.", severity: "warning" }); });

    if (ivaTrimestre > 0) alerts.push({ type: "iva", message: "🧾 IVA estimado " + currentQuarter + ": " + round2(ivaTrimestre).toFixed(2) + "€.", severity: "info" });
    if (vendasMes > 0 && comprasMes > vendasMes * 1.2) alerts.push({ type: "racio", message: "📊 Compras do mês (" + round2(comprasMes).toFixed(2) + "€) excedem vendas em " + round2(((comprasMes / vendasMes) - 1) * 100).toFixed(0) + "%.", severity: "warning" });

    const sheetRH = getStaffSheet(ss);
    const custoMensalRH = (dash && dash.rh && dash.rh.custo_mensal != null) ? parseFloat(String(dash.rh.custo_mensal).replace(",", ".")) || 0 : 0;
    const staffAtivosFromDash = (dash && dash.rh && dash.rh.ativos != null) ? parseInt(dash.rh.ativos, 10) || 0 : 0;

    if (sheetRH && sheetRH.getLastRow() > 1) {
      const rhAlertLastCol = Math.max(sheetRH.getLastColumn(), 20);
      const staffData = sheetRH.getRange(2, 1, sheetRH.getLastRow() - 1, rhAlertLastCol).getValues();
      let staffAtivos = 0, semProvisoes = 0, contratosExpirar = 0;
      const mesAtual = now.getMonth() + 1;
      const anoAtual = now.getFullYear();

      staffData.forEach(r => {
        const status = String(r[8] || "");
        const custoBase = parsePTFloat(r[9]);
        const fator = getStaffCostFactor(status, r[14], r[19] || null, mesAtual, anoAtual);
        if (status === "Ativo") {
          staffAtivos++;
          const provF = parsePTFloat(r[10]), provN = parsePTFloat(r[11]);
          if (provF === 0 && provN === 0) semProvisoes++;
        }
        if (status === "Ativo" || (fator > 0 && custoBase > 0)) {
          const diasContrato = parseInt(r[15], 10) || 0;
          if (diasContrato > 0) {
            let admissao = r[14];
            if (typeof admissao === "string") { const p = admissao.split("/"); if (p.length === 3) admissao = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
            if (admissao instanceof Date && !isNaN(admissao)) {
              const fim = new Date(admissao.getTime() + diasContrato * 86400000);
              const diasRestantes = Math.ceil((fim - now) / 86400000);
              if (diasRestantes > 0 && diasRestantes <= 60) contratosExpirar++;
            }
          }
        }
      });

      if (semProvisoes > 0) alerts.push({ type: "provisoes", message: "📝 " + semProvisoes + " colaborador(es) sem provisões definidas (férias/natal).", severity: "warning" });
      if (contratosExpirar > 0) alerts.push({ type: "contratos", message: "📅 " + contratosExpirar + " contrato(s) a expirar nos próximos 60 dias.", severity: "critical" });

      const ativosParaProd = staffAtivos > 0 ? staffAtivos : staffAtivosFromDash;
      if (ativosParaProd > 0 && vendasMes > 0 && custoMensalRH > 0) {
        const custoPerCapita = custoMensalRH / ativosParaProd;
        const receitaPerCapita = vendasMes / ativosParaProd;
        if (custoPerCapita > receitaPerCapita * 0.8) alerts.push({ type: "produtividade", message: "⚡ Custo médio por colaborador representa mais de 80% da receita per capita.", severity: "warning" });
      }
      if (custoMensalRH > 0) alerts.push({ type: "custoRH", message: "👥 Custo mensal total de RH: " + round2(custoMensalRH).toFixed(2) + "€.", severity: "info" });
    }

    try {
      const ssMaster = SpreadsheetApp.openById(MASTER_DB_ID);
      ensureMasterDBColumns();
      const dataMaster = ssMaster.getSheets()[0].getDataRange().getValues();
      if (ctx.sheetId) {
        const clientRow = dataMaster.find(r => r[2] === ctx.sheetId);
        if (clientRow) {
          const clientName = clientRow[0];
          const mensalidadesOferta = parseInt(clientRow[MASTER_COL_MENSALIDADES_OFERTA - 1]) || 0;
          if (mensalidadesOferta === 1) alerts.push({ type: "oferta_fim", message: "⚠️ Atenção: O período de oferta do cliente " + (clientName || "este") + " termina no próximo mês.", severity: "warning" });
          else if (mensalidadesOferta === 2) alerts.push({ type: "oferta_fim", message: "ℹ️ O cliente " + (clientName || "este") + " tem 2 mensalidades de oferta restantes.", severity: "info" });
        }
      }
    } catch (eOferta) { }

    const sevOrder = { critical: 0, warning: 1, info: 2, error: 3 };
    alerts.sort((a, b) => (sevOrder[a.severity] || 9) - (sevOrder[b.severity] || 9));

    return alerts;
  } catch (e) { return [{ type: "erro", message: MSG_IA_PENDENTE, severity: "error" }]; }
}

/** Extrai último insight por categoria (fleet, finance, stock, rh) da AI_History em memória. */
function _getLastAIInsightsFromBatch(dataAI, email) {
  if (!dataAI || dataAI.length === 0 || !email) return null;
  var emailNorm = (email || "").toString().trim().toLowerCase();
  var rows = dataAI.filter(function (row) { return String(row[1] || "").trim().toLowerCase() === emailNorm; });
  rows.sort(function (a, b) {
    var tA = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
    var tB = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
    return tB - tA;
  });

  var merged = { financeiro: null, stocks: null, rh: null, fleet: null };
  var ensureBlock = function (b) {
    if (!b || typeof b !== "object") return null;
    var sum = (b.summary || b.resumo || b.text || "").toString().trim();
    var arr = Array.isArray(b.todos) ? b.todos : (Array.isArray(b.actions) ? b.actions : []);
    sum = (sum || "").replace(/\{/g, "").replace(/\}/g, "").trim();
    if (!sum && arr.length === 0) return null;
    return { summary: sum || "Erro na análise.", todos: arr };
  };

  for (var i = 0; i < Math.min(rows.length, 20); i++) {
    var raw = rows[i][3];
    var obj = null;
    try { if (typeof raw === "string" && raw.trim()) obj = JSON.parse(raw); else if (raw && typeof raw === "object") obj = raw; } catch (e) { }
    if (!obj || typeof obj !== "object") continue;
    if (obj._raw && typeof obj._raw === "string") { try { var p = JSON.parse(obj._raw); if (p) obj = p; } catch (e) { } }

    if (!merged.financeiro && obj.financeiro) merged.financeiro = ensureBlock(obj.financeiro);
    if (!merged.stocks && obj.stocks) merged.stocks = ensureBlock(obj.stocks);
    if (!merged.rh && obj.rh) merged.rh = ensureBlock(obj.rh);
    if (!merged.fleet && obj.fleet) merged.fleet = ensureBlock(obj.fleet);

    var cat = String(rows[i][2] || "").trim().toLowerCase();
    if (cat === "financeiro" && !merged.financeiro && obj) merged.financeiro = ensureBlock(obj);
    if (cat === "stocks" && !merged.stocks && obj) merged.stocks = ensureBlock(obj);
    if (cat === "rh" && !merged.rh && obj) merged.rh = ensureBlock(obj);
    if (cat === "fleet" && !merged.fleet && obj) merged.fleet = ensureBlock(obj);

    if (merged.financeiro && merged.stocks && merged.rh && merged.fleet) break;
  }

  var hasAny = merged.financeiro || merged.stocks || merged.rh || merged.fleet;
  if (!hasAny) return null;
  return {
    success: true,
    financeiro: merged.financeiro || { summary: "", todos: [] },
    stocks: merged.stocks || { summary: "", todos: [] },
    rh: merged.rh || { summary: "", todos: [] },
    fleet: merged.fleet || { summary: "", todos: [] },
    currentCredits: getAiCredits(email)
  };
}