// Ficheiro: MOD_CRM_Export.js
/// ==========================================
// 🤝 MÓDULO DE CRM E EXPORTAÇÃO CONTABILÍSTICA
// ==========================================

function normalizeEntityName(nomeRaw) {
  if (!nomeRaw || typeof nomeRaw !== "string") return "";
  var s = String(nomeRaw).trim();
  if (!s) return "";
  s = s.replace(/[àáâãäå]/g, "a").replace(/[èéêë]/g, "e").replace(/[ìíîï]/g, "i")
    .replace(/[òóôõö]/g, "o").replace(/[ùúûü]/g, "u").replace(/ç/g, "c").replace(/ñ/g, "n");
  s = s.replace(/\s+/g, " ").trim();
  return s.toLowerCase();
}

function canonicalizeEntityForDisplay(nomeRaw) {
  if (!nomeRaw || typeof nomeRaw !== "string") return (nomeRaw || "").trim();
  var s = String(nomeRaw).trim().toUpperCase();
  if (!s) return "";
  if (/PETROGAL|GALP/.test(s)) return "GALP";
  if (/REPSOL/.test(s)) return "REPSOL";
  if (/PETROB|PRIO/.test(s)) return "PRIO";
  if (/VIA\s*VERDE|BRISA/.test(s)) return "VIA VERDE";
  return nomeRaw.trim();
}

function _aliasEntityForMatch(normalized) {
  if (!normalized) return normalized;
  if (/petrogal|galp/.test(normalized)) return "galp";
  if (/repsol/.test(normalized)) return "repsol";
  if (/petrob|prio/.test(normalized)) return "prio";
  if (/via\s*verde|viaverde|brisa/.test(normalized)) return "via verde";
  return normalized;
}

function _searchEntityInSheet(ss, sheetName, nomeNorm, aliasNorm, colId, colNome) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return "";
  var data = sheet.getRange(2, 1, sheet.getLastRow(), Math.max(colNome + 1, 2)).getValues();
  for (var i = 0; i < data.length; i++) {
    var idVal = String(data[i][colId] || "").trim();
    var nomeVal = String(data[i][colNome] || "").trim();
    if (!idVal || !nomeVal) continue;
    var nomeValNorm = normalizeEntityName(nomeVal);
    var match = (nomeValNorm === nomeNorm || nomeValNorm === aliasNorm || _aliasEntityForMatch(nomeValNorm) === nomeNorm);
    if (!match && nomeNorm.length > 2) {
      var nomeSemEsp = nomeNorm.replace(/\s/g, "");
      var nomeValSemEsp = nomeValNorm.replace(/\s/g, "");
      match = (nomeValSemEsp === nomeSemEsp || nomeValNorm.indexOf(nomeNorm) >= 0 || nomeNorm.indexOf(nomeValNorm) >= 0);
    }
    if (match) return idVal;
  }
  return "";
}

function getEntityIdByName(name, tipoLancamento, impersonateTarget) {
  try {
    var nomeNorm = normalizeEntityName(name || "");
    if (!nomeNorm) return "";
    var aliasNorm = _aliasEntityForMatch(nomeNorm);
    var isCliente = (String(tipoLancamento || "").toLowerCase().trim() === "saída" || String(tipoLancamento || "").toLowerCase().trim() === "saida");
    var colId = 0;
    var colNome = 1;
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return "";

    var ss;
    try { ss = SpreadsheetApp.openById(ctx.sheetId); } catch (openErr) { return ""; }

    var sheetPrimario = isCliente ? CLIENTES_DB_TAB : FORNECEDORES_DB_TAB;
    var sheetSecundario = isCliente ? FORNECEDORES_DB_TAB : CLIENTES_DB_TAB;
    var id = _searchEntityInSheet(ss, sheetPrimario, nomeNorm, aliasNorm, colId, colNome);
    if (id) return id;
    id = _searchEntityInSheet(ss, sheetSecundario, nomeNorm, aliasNorm, colId, colNome);
    return id || "";
  } catch (e) { return ""; }
}

function checkAndCreateDBSheets(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);

    var sheetClientes = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheetClientes) {
      sheetClientes = ss.insertSheet(CLIENTES_DB_TAB);
      sheetClientes.appendRow(CLIENTES_HEADERS);
      sheetClientes.getRange(1, 1, 1, CLIENTES_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
    }

    var sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    if (!sheetAudit) {
      sheetAudit = ss.insertSheet(AUDIT_DB_TAB);
      sheetAudit.appendRow(AUDIT_HEADERS);
      sheetAudit.getRange(1, 1, 1, AUDIT_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function addCliente(payload, impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return { success: false, error: init.error || "Erro ao criar abas." };

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheet) return { success: false, error: "Clientes_DB não encontrada." };

    var idCliente = Utilities.getUuid();
    var now = new Date();
    var dataConsent = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    var row = [
      idCliente,
      (payload.nomeEmpresa || "").toString().trim(),
      (payload.nif || "").toString().trim(),
      (payload.email || "").toString().trim(),
      (payload.telefone || "").toString().trim(),
      (payload.morada || "").toString().trim(),
      dataConsent,
      "Ativo"
    ];
    sheet.appendRow(row);
    return { success: true, idCliente: idCliente };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function editCliente(idCliente, payload, impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return { success: false, error: init.error || "Erro ao criar abas." };

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Cliente não encontrado." };

    var data = sheet.getRange(2, 1, sheet.getLastRow(), 8).getValues();
    var idx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idCliente || "").trim()) { idx = i; break; }
    }
    if (idx < 0) return { success: false, error: "Cliente não encontrado." };

    var row = idx + 2;
    if (payload.nomeEmpresa !== undefined) sheet.getRange(row, 2).setValue(String(payload.nomeEmpresa || "").trim());
    if (payload.nif !== undefined) sheet.getRange(row, 3).setValue(String(payload.nif || "").trim());
    if (payload.email !== undefined) sheet.getRange(row, 4).setValue(String(payload.email || "").trim());
    if (payload.telefone !== undefined) sheet.getRange(row, 5).setValue(String(payload.telefone || "").trim());
    if (payload.morada !== undefined) sheet.getRange(row, 6).setValue(String(payload.morada || "").trim());
    if (payload.status !== undefined) sheet.getRange(row, 8).setValue(String(payload.status || "Ativo").trim());

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function apagarCliente(idCliente, impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return { success: false, error: init.error || "Erro ao criar abas." };

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetClientes = ss.getSheetByName(CLIENTES_DB_TAB);
    var sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    if (!sheetClientes || sheetClientes.getLastRow() < 2) return { success: false, error: "Cliente não encontrado." };

    var data = sheetClientes.getRange(2, 1, sheetClientes.getLastRow(), 8).getValues();
    var idx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idCliente || "").trim()) { idx = i; break; }
    }
    if (idx < 0) return { success: false, error: "Cliente não encontrado." };

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    if (sheetAudit) {
      sheetAudit.appendRow([new Date(), userEmail, idCliente, "ELIMINAR_CLIENTE", "Eliminação a pedido - RGPD"]);
    }

    sheetClientes.deleteRow(idx + 2);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getClienteById(idCliente, impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return null;

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getRange(2, 1, sheet.getLastRow(), 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idCliente || "").trim()) {
        var r = data[i];
        return {
          idCliente: String(r[0] || "").trim(),
          nomeEmpresa: String(r[1] || "").trim(),
          nif: String(r[2] || "").trim(),
          email: String(r[3] || "").trim(),
          telefone: String(r[4] || "").trim(),
          morada: String(r[5] || "").trim(),
          status: String(r[7] || "Ativo").trim()
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

function getClientesCRM(impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return [];

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow(), 8).getValues();
    var out = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var idCliente = String(r[0] || "").trim();
      if (!idCliente) continue;
      var dataConsent = r[6];
      if (dataConsent instanceof Date) dataConsent = Utilities.formatDate(dataConsent, Session.getScriptTimeZone(), "dd/MM/yyyy");
      else dataConsent = String(dataConsent || "");
      out.push({
        idCliente: idCliente,
        nomeEmpresa: String(r[1] || "").trim(),
        emailMasked: maskEmail(r[3]),
        nifMasked: maskNIF(r[2]),
        telefoneMasked: maskTelefone(r[4]),
        morada: String(r[5] || "").trim(),
        dataConsentimento: dataConsent,
        status: String(r[7] || "Ativo").trim()
      });
    }
    return out;
  } catch (e) { return []; }
}

function desbloquearClienteComAuditoria(idCliente, motivoCompleto, impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return { success: false, error: init.error || "Erro ao criar abas." };

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    var sheetClientes = ss.getSheetByName(CLIENTES_DB_TAB);

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    sheetAudit.appendRow([new Date(), userEmail, idCliente, "DESMASCARAR_DADOS", motivoCompleto || ""]);

    if (!sheetClientes || sheetClientes.getLastRow() < 2) return { success: false, error: "Cliente não encontrado." };
    var data = sheetClientes.getRange(2, 1, sheetClientes.getLastRow(), 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idCliente || "").trim()) {
        var r = data[i];
        return {
          success: true,
          email: String(r[3] || "").trim(),
          nif: String(r[2] || "").trim(),
          telefone: String(r[4] || "").trim()
        };
      }
    }
    return { success: false, error: "Cliente não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getClientesForDatalist(impersonateTarget) {
  try {
    var init = checkAndCreateDBSheets(impersonateTarget);
    if (!init.success) return [];
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CLIENTES_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 2, sheet.getLastRow(), 2).getValues();
    var out = [];
    for (var i = 0; i < data.length; i++) {
      var v = String(data[i][0] || "").trim();
      if (v) out.push(v);
    }
    return out.length ? [...new Set(out)].sort() : [];
  } catch (e) { return []; }
}

function ensureFornecedoresSheet(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(FORNECEDORES_DB_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(FORNECEDORES_DB_TAB);
      sheet.appendRow(FORNECEDORES_HEADERS);
      sheet.getRange(1, 1, 1, FORNECEDORES_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
    }
    return { success: true, sheet: sheet };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function generateNextId(prefixo, nomeFolha, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return prefixo + "-001";
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(nomeFolha);
    if (!sheet || sheet.getLastRow() < 2) return prefixo + "-001";
    var colA = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
    var lastNum = 0;
    var re = new RegExp("^" + prefixo + "-?(\\d+)$", "i");
    for (var i = colA.length - 1; i >= 0; i--) {
      var val = String(colA[i][0] || "").trim();
      var m = val.match(re);
      if (m) {
        lastNum = parseInt(m[1], 10) || 0;
        break;
      }
    }
    return prefixo + "-" + ("" + (lastNum + 1)).padStart(3, "0");
  } catch (e) { return prefixo + "-001"; }
}

function getFornecedoresCRM(impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) return [];
    var sheet = init.sheet;
    if (sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 7).getValues();
    var out = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var idVal = String(r[0] || "").trim();
      if (!idVal) continue;
      out.push({
        idFornecedor: idVal,
        nomeEmpresa: String(r[1] || "").trim(),
        emailMasked: maskEmail(r[4]),
        nifMasked: maskNIF(r[2]),
        telefoneMasked: maskTelefone(r[5]),
        categoria: String(r[3] || "").trim(),
        condicoesPagamento: String(r[6] || "").trim()
      });
    }
    return out;
  } catch (e) { return []; }
}

function addFornecedor(payload, impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) throw new Error(init.error || "Erro ao criar aba Fornecedores.");
    var sheet = init.sheet;
    var nif = String(payload.nif || "").trim().replace(/\D/g, "");
    if (nif && sheet.getLastRow() >= 2) {
      var data = sheet.getRange(2, 1, sheet.getLastRow(), 3).getValues();
      for (var i = 0; i < data.length; i++) {
        var existingNif = String(data[i][2] || "").trim().replace(/\D/g, "");
        if (existingNif && existingNif === nif) throw new Error("Já existe um fornecedor com este NIF.");
      }
    }
    var novoId = generateNextId("FORN", FORNECEDORES_DB_TAB, impersonateTarget || null);
    var row = [
      novoId,
      (payload.nomeEmpresa || "").toString().trim(),
      (payload.nif || "").toString().trim(),
      (payload.categoria || "").toString().trim(),
      (payload.email || "").toString().trim(),
      (payload.telefone || "").toString().trim(),
      (payload.condicoesPagamento || "").toString().trim()
    ];
    sheet.appendRow(row);
    return { success: true, idFornecedor: novoId };
  } catch (e) { throw (e && e.message) ? e.message : String(e); }
}

function editFornecedor(idFornecedor, payload, impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) throw new Error(init.error || "Erro ao criar aba Fornecedores.");
    var sheet = init.sheet;
    if (sheet.getLastRow() < 2) throw new Error("Fornecedor não encontrado.");
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 7).getValues();
    var idx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idFornecedor || "").trim()) { idx = i; break; }
    }
    if (idx < 0) throw new Error("Fornecedor não encontrado.");
    var row = idx + 2;
    if (payload.nomeEmpresa !== undefined) sheet.getRange(row, 2).setValue(String(payload.nomeEmpresa || "").trim());
    if (payload.nif !== undefined) sheet.getRange(row, 3).setValue(String(payload.nif || "").trim());
    if (payload.categoria !== undefined) sheet.getRange(row, 4).setValue(String(payload.categoria || "").trim());
    if (payload.email !== undefined) sheet.getRange(row, 5).setValue(String(payload.email || "").trim());
    if (payload.telefone !== undefined) sheet.getRange(row, 6).setValue(String(payload.telefone || "").trim());
    if (payload.condicoesPagamento !== undefined) sheet.getRange(row, 7).setValue(String(payload.condicoesPagamento || "").trim());
    return { success: true };
  } catch (e) { throw (e && e.message) ? e.message : String(e); }
}

function desbloquearFornecedorComAuditoria(idFornecedor, motivoCompleto, impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao aceder à aba Fornecedores." };

    var ctx = getClientContext(impersonateTarget || null);
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    var sheetFornecedores = ss.getSheetByName(FORNECEDORES_DB_TAB);

    if (!sheetAudit) {
      var checkInit = checkAndCreateDBSheets(impersonateTarget);
      if (!checkInit.success) return { success: false, error: checkInit.error || "Erro ao criar abas." };
      sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    }

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";

    if (!sheetFornecedores || sheetFornecedores.getLastRow() < 2) return { success: false, error: "Fornecedor não encontrado." };
    var data = sheetFornecedores.getRange(2, 1, sheetFornecedores.getLastRow(), 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idFornecedor || "").trim()) {
        var r = data[i];
        var nomeFornecedor = String(r[1] || "").trim();
        sheetAudit.appendRow([new Date(), userEmail, idFornecedor + " (" + nomeFornecedor + ")", "DESMASCARAR_DADOS_FORNECEDOR", motivoCompleto || ""]);
        return {
          success: true,
          email: String(r[4] || "").trim(),
          nif: String(r[2] || "").trim(),
          telefone: String(r[5] || "").trim()
        };
      }
    }
    return { success: false, error: "Fornecedor não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function apagarFornecedor(idFornecedor, impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) throw new Error(init.error || "Erro ao criar aba Fornecedores.");
    var sheet = init.sheet;
    if (sheet.getLastRow() < 2) throw new Error("Fornecedor não encontrado.");
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 7).getValues();
    var idx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idFornecedor || "").trim()) { idx = i; break; }
    }
    if (idx < 0) throw new Error("Fornecedor não encontrado.");
    sheet.deleteRow(idx + 2);
    return { success: true };
  } catch (e) { throw (e && e.message) ? e.message : String(e); }
}

function getFornecedorById(idFornecedor, impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) return null;
    var sheet = init.sheet;
    if (sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow(), 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === String(idFornecedor || "").trim()) {
        var r = data[i];
        return {
          idFornecedor: String(r[0] || "").trim(),
          nomeEmpresa: String(r[1] || "").trim(),
          nif: String(r[2] || "").trim(),
          email: String(r[4] || "").trim(),
          telefone: String(r[5] || "").trim(),
          categoria: String(r[3] || "").trim(),
          condicoesPagamento: String(r[6] || "").trim()
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

function getFornecedoresForDatalist(impersonateTarget) {
  try {
    var init = ensureFornecedoresSheet(impersonateTarget || null);
    if (!init.success || !init.sheet) return [];
    var sheet = init.sheet;
    if (sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 2, sheet.getLastRow(), 2).getValues();
    var out = [];
    for (var i = 0; i < data.length; i++) {
      var v = String(data[i][0] || "").trim();
      if (v) out.push(v);
    }
    return out.length ? [...new Set(out)].sort() : [];
  } catch (e) { return []; }
}

function ensureConfigMapeamentoSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG_MAPEAMENTO_TAB);
  if (!sheet) sheet = ss.insertSheet(CONFIG_MAPEAMENTO_TAB);
  if (sheet.getLastRow() < 1) sheet.getRange(1, 1, 1, MAPEAMENTO_HEADERS.length).setValues([MAPEAMENTO_HEADERS]);
  return sheet;
}

function getUnmappedItems(impersonateTarget) {
  try {
    const data = getMasterData(impersonateTarget);
    const cc = data.cc || [];
    const artigosCC = [];
    cc.forEach(function (item) {
      const a = (item.artigo || "").toString().trim();
      if (a && artigosCC.indexOf(a) === -1) artigosCC.push(a);
    });
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { count: 0, items: [] };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(CONFIG_MAPEAMENTO_TAB);
    const mapped = {};
    if (sheet && sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
      rows.forEach(function (r) {
        const key = (r[0] || "").toString().trim();
        if (key) mapped[key] = true;
      });
    }
    const unmapped = artigosCC.filter(function (a) { return !mapped[a]; });
    return { count: unmapped.length, items: unmapped };
  } catch (e) { return { count: 0, items: [] }; }
}

function verifyMappingForExport(dataInicio, dataFim, impersonateTarget) {
  try {
    const data = getMasterData(impersonateTarget);
    const cc = data.cc || [];
    const dIni = parseDate(dataInicio);
    const dFim = parseDate(dataFim);
    if (!dIni || !dFim) return { ok: false, missing: [] };
    const artigosPeriodo = [];
    cc.forEach(function (item) {
      const d = parseDate(item.dataNorm || item.data);
      if (d && !isNaN(d.getTime()) && d >= dIni && d <= dFim) {
        const a = (item.artigo || "").toString().trim();
        if (a && artigosPeriodo.indexOf(a) === -1) artigosPeriodo.push(a);
      }
    });
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { ok: artigosPeriodo.length === 0, missing: artigosPeriodo };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(CONFIG_MAPEAMENTO_TAB);
    const mapped = {};
    if (sheet && sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
      rows.forEach(function (r) {
        const key = (r[0] || "").toString().trim();
        if (key) mapped[key] = true;
      });
    }
    const missing = artigosPeriodo.filter(function (a) { return !mapped[a]; });
    return { ok: missing.length === 0, missing: missing };
  } catch (e) { return { ok: false, missing: [] }; }
}

function saveMapping(mappingData, impersonateTarget) {
  try {
    if (!Array.isArray(mappingData) || mappingData.length === 0) return { success: false, error: "Dados inválidos." };
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureConfigMapeamentoSheet(ss);
    const existing = {};
    if (sheet.getLastRow() > 1) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
      rows.forEach(function (r, i) {
        const key = (r[0] || "").toString().trim();
        if (key) existing[key] = i + 2;
      });
    }
    mappingData.forEach(function (p) {
      const artigo = (p.artigoFlowly || "").toString().trim();
      const conta = (p.contaSoftware || "").toString().trim();
      if (!artigo || !conta) return;
      if (existing[artigo]) {
        sheet.getRange(existing[artigo], 2).setValue(conta);
      } else {
        sheet.appendRow([artigo, conta]);
        existing[artigo] = sheet.getLastRow();
      }
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function clearMappings(impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(CONFIG_MAPEAMENTO_TAB);
    if (!sheet) return { success: true };
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow, 2).clearContent();
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function generateAccountingFile(dataInicio, dataFim, template, impersonateTarget) {
  try {
    const data = getMasterData(impersonateTarget);
    const cc = data.cc || [];
    const dIni = parseDate(dataInicio);
    const dFim = parseDate(dataFim);
    if (!dIni || !dFim) return { success: false, error: "Datas inválidas." };
    const filtered = cc.filter(function (item) {
      const d = parseDate(item.dataNorm || item.data);
      return d && !isNaN(d.getTime()) && d >= dIni && d <= dFim;
    });
    const ctx = getClientContext(impersonateTarget);
    const mapping = {};
    if (ctx.sheetId) {
      const ss = SpreadsheetApp.openById(ctx.sheetId);
      const sheet = ss.getSheetByName(CONFIG_MAPEAMENTO_TAB);
      if (sheet && sheet.getLastRow() > 1) {
        const rows = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
        rows.forEach(function (r) {
          const k = (r[0] || "").toString().trim();
          const v = (r[1] || "").toString().trim();
          if (k) mapping[k] = v;
        });
      }
    }
    let matrix = [];
    let ccSheet = null;
    if (ctx.sheetId) {
      try {
        ccSheet = SpreadsheetApp.openById(ctx.sheetId).getSheetByName(SHEET_TAB_NAME);
      } catch (e) { }
    }
    const templateNorm = (template || "").toString().toLowerCase();
    if (templateNorm === "primavera") {
      matrix = [["TipoDoc", "Serie", "NumDoc", "Data", "Entidade", "NIF", "ContaBase", "ContaIVA", "ValorBase", "ValorIVA"]];
      filtered.forEach(function (item) {
        const contaBase = mapping[item.artigo] || item.artigo || "";
        const valorBase = parseFloat(item.total) || 0;
        const taxaIva = item.taxaIva != null ? (item.taxaIva > 1 ? item.taxaIva / 100 : item.taxaIva) : 0.23;
        const valorIVA = valorBase * taxaIva;
        let nif = "";
        if (ccSheet && item.rowIndex) {
          try {
            const obs = ccSheet.getRange(item.rowIndex, 13).getValue();
            const m = String(obs || "").match(/NIF:\s*([^\s]+)/i);
            if (m) nif = m[1].trim();
          } catch (e) { }
        }
        if (!nif && (item.fornecedor || "").indexOf("NIF:") >= 0) nif = String(item.fornecedor).replace(/.*NIF:\s*/i, "").trim();
        matrix.push([
          item.tipo === "entrada" ? "FR" : "FR",
          "",
          item.docId || "",
          item.dataNorm || item.data || "",
          item.fornecedor || "",
          nif,
          contaBase,
          "2431",
          valorBase.toFixed(2),
          valorIVA.toFixed(2)
        ]);
      });
    } else if (templateNorm === "sage50c") {
      matrix = [["Diario", "Documento", "Data", "Conta", "Contribuinte", "Descricao", "Debito", "Credito"]];
      filtered.forEach(function (item) {
        const conta = mapping[item.artigo] || item.artigo || "";
        const valor = parseFloat(item.total) || 0;
        const isEntrada = (item.tipo || "").toLowerCase() === "entrada";
        matrix.push([
          "Vendas",
          item.docId || "",
          item.dataNorm || item.data || "",
          conta,
          item.fornecedor || "",
          item.artigo || "",
          isEntrada ? valor.toFixed(2) : "",
          isEntrada ? "" : valor.toFixed(2)
        ]);
      });
    } else {
      return { success: false, error: "Template não suportado." };
    }
    const csv = matrix.map(function (row) {
      return row.map(function (cell) {
        const s = String(cell || "");
        return s.indexOf(";") >= 0 || s.indexOf('"') >= 0 ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(";");
    }).join("\r\n");
    const bom = "\uFEFF";
    const b64 = Utilities.base64Encode(Utilities.newBlob(bom + csv, "text/csv; charset=utf-8").getBytes());
    const today = new Date();
    const fn = "export_" + today.getFullYear() + String(today.getMonth() + 1).padStart(2, "0") + String(today.getDate()).padStart(2, "0") + ".csv";
    return { success: true, base64: b64, filename: fn, mimeType: "text/csv; charset=utf-8" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getEmailHtmlBody(gestorNome, count, itemsList) {
  const itemsHtml = itemsList.map(function (item) {
    return "<li style='margin-bottom:8px;'>" + (item || "").replace(/</g, "&lt;") + "</li>";
  }).join("");
  const appUrl = ScriptApp.getService().getUrl();
  var innerBody = "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Olá, <strong>" + (gestorNome || "Gestor").replace(/</g, "&lt;") + "</strong></p>" +
                  "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Detetámos que a sua conta tem <strong style='color:#EF4444;'>" + count + " itens pendentes</strong> de mapeamento contabilístico.</p>" +
                  "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Para garantir que a sua exportação para o Sage/Primavera seja precisa e sem erros, necessitamos que associe estes artigos ao seu plano de contas:</p>" +
                  "<ul style='background-color:#F1F5F9;padding:20px 40px;border-radius:8px;color:#64748B;margin-bottom:24px;'>" + itemsHtml + "</ul>" +
                  "<table width='100%' cellpadding='0' cellspacing='0'><tr><td align='center'><a href=\"" + appUrl + "\" style=\"display:inline-block;background:#10B981;color:#ffffff;font-size:13px;font-weight:800;text-decoration:none;padding:16px 40px;border-radius:50px;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 4px 16px rgba(16,185,129,0.4);\">Mapear Agora na Flowly</a></td></tr></table>";
  return _buildStandardEmailHTML("Aviso de Mapeamento", innerBody);
}

function cronSendUnmappedAlerts() {
  if (new Date().getDate() % 2 !== 0) return;
  try {
    const res = getSuperAdminDashboardData(SUPER_ADMIN_EMAIL);
    if (!res.success || !res.clients) return;
    const ativos = res.clients.filter(function (c) { return c.status === "Ativo" && c.modules && c.modules.cc !== false; });
    ativos.forEach(function (cli) {
      try {
        const r = getUnmappedItems(cli.email);
        if (r && r.count > 0 && r.items && r.items.length > 0) {
          const html = getEmailHtmlBody(cli.name || "Gestor", r.count, r.items);
          let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
          var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: html, inlineImages: { flowlyLogo: logoBlob } };
          GmailApp.sendEmail(cli.email, "⚠️ Flowly 360: Tem " + r.count + " artigos pendentes para exportação", "", options);
        }
      } catch (inner) { Logger.log("cronSendUnmappedAlerts: " + cli.email + " - " + inner); }
    });
  } catch (e) { Logger.log("cronSendUnmappedAlerts: " + e); }
}