// Ficheiro: MOD_Condominios.js
/// ==========================================
// 🏢 MÓDULO DE GESTÃO DE CONDOMÍNIOS
// ==========================================

function _assertCondominium(ctx) {
  if (!ctx || (ctx.businessVertical !== 'Condominium' && ctx.role !== 'SuperAdmin')) {
    return { success: false, error: "Esta funcionalidade está disponível apenas para clientes com vertical 'Condominium'." };
  }
  return { success: true };
}

function _getCondoCacheKey(ctx, type) {
  if (!ctx || !ctx.sheetId) return null;
  return "CONDO_" + type + "_" + ctx.sheetId;
}

function _invalidateCondoCache(ctx) {
  if (!ctx || !ctx.sheetId) return;
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll([
      "CONDO_BLDG_SAFE_" + ctx.sheetId,
      "CONDO_FRAC_" + ctx.sheetId
    ]);
  } catch (e) {}
}

function setupCondominiumTabs(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    var ss = SpreadsheetApp.openById(ctx.sheetId);

    var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheetB) {
      sheetB = ss.insertSheet(BUILDINGS_DB_TAB);
      sheetB.appendRow(BUILDINGS_HEADERS);
      sheetB.getRange(1, 1, 1, BUILDINGS_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
      var protection = sheetB.protect().setDescription("Proteção Buildings_DB — IBAN e Stripe_Key só alteráveis via protocolo 2FA");
      protection.setUnprotectedRanges([sheetB.getRange("A:D"), sheetB.getRange("F:H"), sheetB.getRange("J:K")]);
    }

    var sheetU = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheetU) {
      sheetU = ss.insertSheet(UNITS_DB_TAB);
      sheetU.appendRow(UNITS_HEADERS);
      sheetU.getRange(1, 1, 1, UNITS_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
    }

    return { success: true, message: "Buildings_DB e Units_DB criadas com sucesso." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getBuildingsSafe(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    
    var cacheKey = _getCondoCacheKey(ctx, "BLDG_SAFE");
    var cache = CacheService.getScriptCache();
    if (cacheKey) {
      var cached = cache.get(cacheKey);
      if (cached) { try { return JSON.parse(cached); } catch(e){} }
    }

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, buildings: [] };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, BUILDINGS_HEADERS.length).getValues();
    var buildings = data.map(function (row) {
      var obj = {};
      BUILDINGS_HEADERS.forEach(function (h, i) { obj[h] = row[i] !== undefined ? row[i] : ""; });
      obj["NIF_Condominio"] = _maskSensitiveField(String(obj["NIF_Condominio"] || ""), 2, 2);
      obj["IBAN"] = _maskSensitiveField(String(obj["IBAN"] || ""), 4, 4);
      obj["Stripe_Key"] = _maskSensitiveField(String(obj["Stripe_Key"] || ""), 7, 4);
      return obj;
    });
    var res = { success: true, buildings: buildings };
    if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(res), 1800); } catch(e){} }
    return res;
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getBuildings(impersonateTarget) {
  var res = getBuildingsSafe(impersonateTarget || null);
  if (!res.success) return [];
  return (res.buildings || []).map(function (b) {
    return { id: b['ID_Predio'] || '', nome: b['Nome_Predio'] || '', morada: b['Morada'] || '', nif: b['NIF_Condominio'] || '', iban: b['IBAN'] || '', status: b['Status'] || 'Ativo' };
  });
}

function getFracoes(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return [];

    var cacheKey = _getCondoCacheKey(ctx, "FRAC");
    var cache = CacheService.getScriptCache();
    if (cacheKey) {
      var cached = cache.get(cacheKey);
      if (cached) { try { return JSON.parse(cached); } catch(e){} }
    }

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, UNITS_HEADERS.length).getValues();
    var res = data.map(function (row) {
      return { id: row[0] || '', predioId: row[1] || '', identificador: row[2] || '', proprietario: row[3] || '', email: row[5] || '', telefone: row[6] || '', quota: row[7] || 0, tipo: row[8] !== undefined ? row[8] : '', permilagem: row[8] || 0, status: row[9] || 'Vazia' };
    }).filter(function (f) { return f.id; });
    if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(res), 1800); } catch(e){} }
    return res;
  } catch (e) { return []; }
}

function getBuildingById(predioId, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return null;
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, BUILDINGS_HEADERS.length).getValues();
    var row = data.find(function (r) { return r[0] === predioId; });
    if (!row) return null;
    return { id: row[0] || '', nome: row[1] || '', morada: row[2] || '', nif: '', iban: '', temElevador: row[11] === 'Sim', manutencaoData: row[12] || '' };
  } catch (e) { return null; }
}

function getFracaoById(fracaoId, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return null;
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, UNITS_HEADERS.length).getValues();
    var row = data.find(function (r) { return r[0] === fracaoId; });
    if (!row) return null;
    return { id: row[0] || '', predioId: row[1] || '', identificador: row[2] || '', proprietario: row[3] || '', email: row[5] || '', telefone: row[6] || '', quota: row[7] || 0, tipo: row[8] !== undefined ? String(row[8]) : '', permilagem: row[8] || 0, status: row[9] || 'Vazia' };
  } catch (e) { return null; }
}

function deleteBuilding(predioId, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet) return { success: false, error: 'Buildings_DB não encontrada.' };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === predioId) { sheet.deleteRow(i + 1); break; }
    }
    _invalidateCondoCache(ctx);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteFracao(fracaoId, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet) return { success: false, error: 'Units_DB não encontrada.' };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === fracaoId) { sheet.deleteRow(i + 1); break; }
    }
    _invalidateCondoCache(ctx);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _generateCondoId(sheet, prefix, colIndex) {
  if (!sheet || sheet.getLastRow() < 2) return prefix + "-001";
  var data = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  var max = 0;
  data.forEach(function (r) {
    var m = String(r[0] || "").match(/^[A-Z]+-(\d+)$/);
    if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
  });
  return prefix + "-" + String(max + 1).padStart(3, "0");
}

function _isValidNIF(nif) { return /^\d{9}$/.test(String(nif || "").trim()); }

function _ensureUnitsIdEntidadeCol(sheet) {
  if (!sheet) return;
  if (sheet.getLastColumn() < 12) sheet.getRange(1, 12).setValue("ID_Entidade");
}

function _incrementBuildingFracoes(sheetB, idPredio, delta) {
  if (!sheetB || sheetB.getLastRow() < 2) return;
  var data = sheetB.getRange(2, 1, sheetB.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][_BLDG_COL_ID] || "").trim() === String(idPredio).trim()) {
      var atual = parseInt(data[i][6], 10) || 0;
      sheetB.getRange(i + 2, 7).setValue(Math.max(0, atual + delta));
      return;
    }
  }
}

function createBuilding(payload, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;

    if (!payload || !payload.nomePredio || !payload.nifCondominio || !payload.adminEmail) return { success: false, error: "Campos obrigatórios em falta." };
    if (!_isValidNIF(payload.nifCondominio)) return { success: false, error: "NIF inválido." };
    if (String(payload.adminEmail).indexOf("@") === -1) return { success: false, error: "Email inválido." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet) {
      setupCondominiumTabs(impersonateTarget);
      sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
      if (!sheet) return { success: false, error: "Falha ao criar a aba Buildings_DB." };
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim().toLowerCase() === String(payload.nomePredio || "").trim().toLowerCase() ||
          String(data[i][3] || "").trim() === String(payload.nifCondominio || "").trim()) {
        return { success: false, error: "Este prédio já existe na sua conta (nome ou NIF). Edite o registo existente em vez de criar um novo." };
      }
    }

    var idPredio = _generateCondoId(sheet, "PRED", _BLDG_COL_ID);
    var now = new Date();
    var row = new Array(BUILDINGS_HEADERS.length).fill("");
    row[0] = idPredio; row[1] = String(payload.nomePredio || "").trim(); row[2] = String(payload.morada || "").trim();
    row[3] = String(payload.nifCondominio).trim(); row[4] = ""; row[5] = String(payload.adminEmail || "").trim();
    row[6] = 0; row[7] = "Ativo"; row[8] = ""; row[9] = String(payload.notas || "").trim(); row[10] = now;
    row[11] = payload.temElevador ? "Sim" : "Não"; row[12] = payload.manutencaoData || "";
    sheet.appendRow(row);

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Predio", idPredio, "CREATE_BUILDING", "", "", "Criação: " + row[1]);
    _invalidateCondoCache(ctx);
    return { success: true, idPredio: idPredio, message: "Prédio '" + row[1] + "' criado com sucesso." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateBuilding(idPredio, payload, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;
    if (!idPredio) return { success: false, error: "idPredio é obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Prédio não encontrado." };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_BLDG_COL_ID] || "").trim() === String(idPredio).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Prédio não encontrado: " + idPredio };

    var sr = rowIndex + 1;
    if (payload.nomePredio != null) sheet.getRange(sr, 2).setValue(String(payload.nomePredio).trim());
    if (payload.morada != null) sheet.getRange(sr, 3).setValue(String(payload.morada).trim());
    if (payload.adminEmail != null) sheet.getRange(sr, 6).setValue(String(payload.adminEmail).trim());
    if (payload.totalFracoes != null) sheet.getRange(sr, 7).setValue(parseInt(payload.totalFracoes, 10) || 0);
    if (payload.notas != null) sheet.getRange(sr, 10).setValue(String(payload.notas).trim());
    if (payload.temElevador !== undefined) sheet.getRange(sr, 12).setValue(payload.temElevador ? "Sim" : "Não");
    if (payload.manutencaoData !== undefined) sheet.getRange(sr, 13).setValue(String(payload.manutencaoData || ""));

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Predio", idPredio, "UPDATE_BUILDING", "", "", "Campos atualizados.");
    _invalidateCondoCache(ctx);
    return { success: true, message: "Prédio " + idPredio + " atualizado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function inactivateBuilding(idPredio, motivo, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;
    if (!idPredio) return { success: false, error: "idPredio é obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Prédio não encontrado." };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_BLDG_COL_ID] || "").trim() === String(idPredio).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Prédio não encontrado: " + idPredio };
    if (String(data[rowIndex][7] || "").trim() === "Inativo") return { success: false, error: "O prédio " + idPredio + " já se encontra inativo." };

    sheet.getRange(rowIndex + 1, 8).setValue("Inativo");
    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Predio", idPredio, "INACTIVATE_BUILDING", "Ativo", "Inativo", motivo || "Sem motivo.");
    _invalidateCondoCache(ctx);
    return { success: true, message: "Prédio " + idPredio + " desativado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function createUnit(payload, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;

    if (!payload || !payload.idPredio || !payload.designacao || !payload.proprietario) return { success: false, error: "Campos obrigatórios em falta." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
    var sheetU = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheetU || !sheetB) {
      setupCondominiumTabs(impersonateTarget);
      sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
      sheetU = ss.getSheetByName(UNITS_DB_TAB);
      if (!sheetU) return { success: false, error: "Falha ao criar a aba Units_DB." };
    }

    if (sheetB && sheetB.getLastRow() > 1) {
      var bData = sheetB.getRange(2, 1, sheetB.getLastRow() - 1, 1).getValues();
      var bExists = false;
      for (var bi = 0; bi < bData.length; bi++) { if (String(bData[bi][0] || "").trim() === String(payload.idPredio).trim()) { bExists = true; break; } }
      if (!bExists) return { success: false, error: "Prédio de referência não encontrado." };
    }

    _ensureUnitsIdEntidadeCol(sheetU);

    var idFracao = _generateCondoId(sheetU, "FRAC", _UNIT_COL_ID);
    var now = new Date();
    var row = new Array(12).fill("");
    row[0] = idFracao; row[1] = String(payload.idPredio || "").trim(); row[2] = String(payload.designacao || "").trim();
    row[3] = String(payload.proprietario || "").trim(); row[4] = String(payload.nifProprietario || "").trim();
    row[5] = String(payload.emailProprietario || "").trim(); row[6] = String(payload.telefone || "").trim();
    row[7] = parseFloat(payload.quotaMensal) || 0; row[8] = parseFloat(payload.quinhaoPerc) || 0;
    row[9] = "Ativo"; row[10] = now; row[11] = String(payload.idEntidade || "").trim();
    sheetU.appendRow(row);
    _incrementBuildingFracoes(sheetB, payload.idPredio, 1);

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Fracao", idFracao, "CREATE_UNIT", "", "", "Fração criada no prédio: " + payload.idPredio);
    _invalidateCondoCache(ctx);
    return { success: true, idFracao: idFracao, message: "Fração '" + row[2] + "' criada." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateUnit(idFracao, payload, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;
    if (!idFracao) return { success: false, error: "idFracao é obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Fração não encontrada." };
    _ensureUnitsIdEntidadeCol(sheet);

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_UNIT_COL_ID] || "").trim() === String(idFracao).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Fração não encontrada: " + idFracao };

    var sr = rowIndex + 1;
    if (payload.designacao != null) sheet.getRange(sr, _UNIT_COL_DESIGNACAO + 1).setValue(String(payload.designacao).trim());
    if (payload.proprietario != null) sheet.getRange(sr, _UNIT_COL_PROPRIETARIO + 1).setValue(String(payload.proprietario).trim());
    if (payload.emailProprietario != null) sheet.getRange(sr, _UNIT_COL_EMAIL_PROP + 1).setValue(String(payload.emailProprietario).trim());
    if (payload.telefone != null) sheet.getRange(sr, _UNIT_COL_TELEFONE + 1).setValue(String(payload.telefone).trim());
    if (payload.quotaMensal != null) sheet.getRange(sr, _UNIT_COL_QUOTA + 1).setValue(parseFloat(payload.quotaMensal) || 0);
    if (payload.quinhaoPerc != null) sheet.getRange(sr, _UNIT_COL_QUINHAO + 1).setValue(parseFloat(payload.quinhaoPerc) || 0);
    if (payload.idEntidade != null) sheet.getRange(sr, _UNIT_COL_ID_ENTIDADE + 1).setValue(String(payload.idEntidade).trim());

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Fracao", idFracao, "UPDATE_UNIT", "", "", "Campos atualizados");
    _invalidateCondoCache(ctx);
    return { success: true, message: "Fração " + idFracao + " atualizada." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function inactivateUnit(idFracao, motivo, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;
    if (!idFracao) return { success: false, error: "idFracao é obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Fração não encontrada." };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_UNIT_COL_ID] || "").trim() === String(idFracao).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Fração não encontrada: " + idFracao };
    if (String(data[rowIndex][_UNIT_COL_STATUS] || "").trim() === "Inativo") return { success: false, error: "Fração já inativa." };

    sheet.getRange(rowIndex + 1, _UNIT_COL_STATUS + 1).setValue("Inativo");

    var idPredio = String(data[rowIndex][_UNIT_COL_ID_PREDIO] || "").trim();
    if (idPredio) {
      var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
      if (sheetB) _incrementBuildingFracoes(sheetB, idPredio, -1);
    }

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Fracao", idFracao, "INACTIVATE_UNIT", "Ativo", "Inativo", motivo || "Sem motivo.");
    _invalidateCondoCache(ctx);
    return { success: true, message: "Fração " + idFracao + " desativada." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function createFracao(payload, impersonateTarget) {
  var mapped = { 
      idPredio: payload.predioId || '', 
      designacao: payload.identificador || '', 
      proprietario: payload.proprietario || '', 
      emailProprietario: payload.emailProprietario || '',
      telefone: payload.telefone || '',
      quotaMensal: parseFloat(payload.quota) || 0,
      quinhaoPerc: parseFloat(payload.permilagem) || 0 
  };
  return createUnit(mapped, impersonateTarget);
}

function solicitarTrocaIBAN(idPredio, novoIBAN, motivo, impersonateTarget) {
  try {
    if (!idPredio || !novoIBAN || !motivo) return { success: false, error: "Parâmetros obrigatórios em falta." };
    var ibanClean = String(novoIBAN).replace(/\s/g, "").toUpperCase();
    if (!/^[A-Z]{2}[0-9]{2}[A-Z0-9]{1,30}$/.test(ibanClean) || ibanClean.length < 15) return { success: false, error: "Formato de IBAN inválido." };
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Buildings_DB não encontrada." };

    var data = sheet.getDataRange().getValues();
    var buildingRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_BLDG_COL_ID] || "").trim() === String(idPredio).trim()) { buildingRow = i; break; }
    }
    if (buildingRow === -1) return { success: false, error: "Prédio não encontrado." };

    var token = String(Math.floor(100000 + Math.random() * 900000));
    var requestedBy = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    var adminEmail = String(data[buildingRow][_BLDG_COL_ADMIN_EMAIL] || ctx.clientEmail || "").trim();
    var expires = new Date().getTime() + (10 * 60 * 1000);

    var propKey = "IBAN_TOKEN_" + ctx.sheetId + "_" + String(idPredio).replace(/[^A-Z0-9]/gi, "_");
    var propValue = JSON.stringify({ token: token, novoIBAN: ibanClean, motivo: String(motivo).trim(), requestedBy: requestedBy, expires: expires });
    PropertiesService.getScriptProperties().setProperty(propKey, propValue);

    if (adminEmail && adminEmail.includes("@")) {
      var ibanMasked = _maskSensitiveField(ibanClean, 4, 4);
      var nomePredio = String(data[buildingRow][_BLDG_COL_NOME] || idPredio);
      var innerBody = "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Foi solicitada uma alteração de IBAN para o prédio <strong>" + nomePredio + "</strong>.</p>" +
                      "<ul style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 24px;padding-left:20px;'>" +
                      "<li><strong>IBAN Novo:</strong> " + ibanMasked + "</li>" +
                      "<li><strong>Motivo:</strong> " + String(motivo).trim() + "</li>" +
                      "</ul>" +
                      "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>O seu código de confirmação é:</p>" +
                      "<div style='background:#F1F5F9;padding:12px 24px;border-radius:8px;font-size:24px;font-weight:bold;letter-spacing:4px;text-align:center;color:#0F172A;margin-bottom:24px;'>" + token + "</div>" +
                      "<p style='color:#94A3B8;font-size:12px;margin:0;'>Este código expira em 10 minutos.</p>";
      var htmlBody = _buildStandardEmailHTML("Alteração de IBAN", innerBody);
      let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
      var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
      GmailApp.sendEmail(adminEmail, "[Flowly Condomínios] Pedido de Alteração de IBAN — " + nomePredio, "", options);
    }
    return { success: true, message: "Código de verificação enviado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function validarETrocarIBAN(idPredio, token, impersonateTarget) {
  try {
    if (!idPredio || !token) return { success: false, error: "idPredio e token são obrigatórios." };
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var propKey = "IBAN_TOKEN_" + ctx.sheetId + "_" + String(idPredio).replace(/[^A-Z0-9]/gi, "_");
    var raw = PropertiesService.getScriptProperties().getProperty(propKey);
    if (!raw) return { success: false, error: "Pedido expirado. Inicie um novo pedido." };
    var pending;
    try { pending = JSON.parse(raw); } catch (e) { PropertiesService.getScriptProperties().deleteProperty(propKey); return { success: false, error: "Token corrompido." }; }
    if (!pending.expires || new Date().getTime() > pending.expires) { PropertiesService.getScriptProperties().deleteProperty(propKey); return { success: false, error: "Token expirado." }; }
    if (String(token).trim() !== String(pending.token).trim()) return { success: false, error: "Código inválido." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheet) return { success: false, error: "Buildings_DB não encontrada." };

    var data = sheet.getDataRange().getValues();
    var buildingRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_BLDG_COL_ID] || "").trim() === String(idPredio).trim()) { buildingRow = i; break; }
    }
    if (buildingRow === -1) return { success: false, error: "Prédio não encontrado." };

    var ibanAnterior = String(data[buildingRow][_BLDG_COL_IBAN] || "").trim();
    sheet.getRange(buildingRow + 1, _BLDG_COL_IBAN + 1).setValue(pending.novoIBAN);

    var userEmail = Session.getActiveUser().getEmail() || pending.requestedBy || "desconhecido";
    _auditCondominioAction(ss, userEmail, "IBAN", idPredio, "TROCA_IBAN", _maskSensitiveField(ibanAnterior, 4, 4), _maskSensitiveField(pending.novoIBAN, 4, 4), pending.motivo);

    PropertiesService.getScriptProperties().deleteProperty(propKey);
    return { success: true, message: "IBAN atualizado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function confirmarTrocaIBAN(predioId, codigo, impersonateTarget) {
  return validarETrocarIBAN(predioId, codigo, impersonateTarget);
}

// ==========================================
// 💰 MOTOR FINANCEIRO DE CONDOMÍNIOS
// ==========================================

/**
 * Calcula e redistribui a permilagem automática para todas as frações de um prédio.
 * Total = 1000. Cada fração recebe 1000 / totalFrações.
 * Permite override manual posterior.
 */
function recalcularPermilagem(idPredio, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idPredio) return { success: false, error: "idPredio obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, message: "Sem frações.", total: 0 };

    var data = sheet.getDataRange().getValues();
    var indices = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_UNIT_COL_ID_PREDIO] || "").trim() === String(idPredio).trim() &&
          String(data[i][_UNIT_COL_STATUS] || "").trim() === "Ativo") {
        indices.push(i + 1); // linha da sheet (1-indexed)
      }
    }
    if (indices.length === 0) return { success: true, message: "Sem frações ativas.", total: 0 };

    var permilagem = Math.round((1000 / indices.length) * 100) / 100;
    var resto = Math.round((1000 - permilagem * indices.length) * 100) / 100;

    indices.forEach(function (rowNum, idx) {
      var valor = (idx === indices.length - 1) ? permilagem + resto : permilagem;
      sheet.getRange(rowNum, _UNIT_COL_QUINHAO + 1).setValue(valor);
    });

    _invalidateCondoCache(ctx);
    return { success: true, message: "Permilagem recalculada para " + indices.length + " frações (" + permilagem + "‰ cada).", total: indices.length };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Aba "Quotas_CC" — estrutura de cabeçalhos.
 */
var QUOTAS_CC_TAB = "Quotas_CC";
var QUOTAS_CC_HEADERS = [
  "ID_Quota", "ID_Predio", "ID_Fracao", "Designacao_Fracao", "Proprietario",
  "Email_Proprietario", "Telefone", "Mes_Referencia", "Ano_Referencia",
  "Valor", "Tipo", "Status", "Data_Vencimento", "Data_Pagamento",
  "Metodo_Pagamento", "Notas", "Timestamp_Criacao", "Criado_Por"
];
var _QUOTA_COL_ID = 0;
var _QUOTA_COL_ID_PREDIO = 1;
var _QUOTA_COL_ID_FRACAO = 2;
var _QUOTA_COL_DESIGNACAO = 3;
var _QUOTA_COL_PROP = 4;
var _QUOTA_COL_EMAIL = 5;
var _QUOTA_COL_TELEFONE = 6;
var _QUOTA_COL_MES = 7;
var _QUOTA_COL_ANO = 8;
var _QUOTA_COL_VALOR = 9;
var _QUOTA_COL_TIPO = 10;
var _QUOTA_COL_STATUS = 11;
var _QUOTA_COL_VENCIMENTO = 12;
var _QUOTA_COL_DATA_PAG = 13;
var _QUOTA_COL_METODO = 14;
var _QUOTA_COL_NOTAS = 15;
var _QUOTA_COL_TS = 16;
var _QUOTA_COL_USER = 17;

function _ensureQuotasCCTab(ss) {
  var sheet = ss.getSheetByName(QUOTAS_CC_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(QUOTAS_CC_TAB);
    sheet.appendRow(QUOTAS_CC_HEADERS);
    sheet.getRange(1, 1, 1, QUOTAS_CC_HEADERS.length).setFontWeight("bold").setBackground("#E2F8F0");
  }
  return sheet;
}

function _generateQuotaId(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return "QT-001";
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var max = 0;
  data.forEach(function (r) {
    var m = String(r[0] || "").match(/^QT-(\d+)$/);
    if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
  });
  return "QT-" + String(max + 1).padStart(3, "0");
}

/**
 * Gera quotas mensais para TODAS as frações ativas de um prédio.
 * Evita duplicação (verifica se já existe quota para o mesmo mês/ano/fracao).
 */
function gerarQuotasMensais(idsPredio, mes, ano, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;

    if (!idsPredio || (Array.isArray(idsPredio) && idsPredio.length === 0)) return { success: false, error: "idPredio obrigatório." };
    var arrPredios = Array.isArray(idsPredio) ? idsPredio : [idsPredio];

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetUnits = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheetUnits || sheetUnits.getLastRow() < 2) return { success: false, error: "Nenhuma fração configurada." };

    var unitsData = sheetUnits.getDataRange().getValues();
    var sheetQuotas = ss.getSheetByName(QUOTAS_CC_TAB);
    if (!sheetQuotas) {
      sheetQuotas = ss.insertSheet(QUOTAS_CC_TAB);
      sheetQuotas.appendRow(QUOTAS_CC_HEADERS);
      sheetQuotas.getRange(1, 1, 1, QUOTAS_CC_HEADERS.length).setFontWeight("bold").setBackground("#E2E8F0");
    }

    var existingQuotas = sheetQuotas.getDataRange().getValues();
    var mapKey = function (r) { return r[_QUOTA_COL_ID_PREDIO] + "|" + r[_QUOTA_COL_ID_FRACAO] + "|" + r[_QUOTA_COL_MES] + "|" + r[_QUOTA_COL_ANO]; };
    var setKeys = new Set(existingQuotas.map(mapKey));

    var now = new Date();
    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    var criadas = 0;
    arrPredios.forEach(function (idPredio) {
      for (var i = 1; i < unitsData.length; i++) {
        var row = unitsData[i];
        if (String(row[_UNIT_COL_ID_PREDIO] || "").trim() !== String(idPredio).trim()) continue;
        if (String(row[_UNIT_COL_STATUS] || "").trim() !== "Ativo") continue;

        var idFracao = row[_UNIT_COL_ID];
        var key = idPredio + "|" + idFracao + "|" + mes + "|" + ano;
        if (setKeys.has(key)) continue;

        var quota = parseFloat(row[_UNIT_COL_QUOTA]) || 0;
        if (quota <= 0) continue;

        var vencimento = new Date(ano, mes - 1, 10);
        var idQuota = "Q-" + Utilities.getUuid().substring(0, 8).toUpperCase();

        sheetQuotas.appendRow([
          idQuota, idPredio, idFracao, row[_UNIT_COL_DESIGNACAO], row[_UNIT_COL_PROPRIETARIO],
          row[_UNIT_COL_EMAIL_PROP], row[_UNIT_COL_TELEFONE], mes, ano, quota,
          "Mensal", "Pendente", vencimento, "", "", "", now, userEmail
        ]);
        criadas++;
      }
    });

    _auditCondominioAction(ss, userEmail, "Quotas", arrPredios.join(","), "GERAR_QUOTAS_MENSAL", "", "", "Mês: " + mes + " | Ano: " + ano + " | Criadas: " + criadas);
    
    return { success: true, message: "Sucesso! " + criadas + " quotas geradas." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Adiciona uma quota extra (obras, fundo, seguro, etc) a UMA ou TODAS as frações.
 */
function adicionarQuotaExtra(idsPredio, idFracao, descricao, valor, dataVencimento, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idsPredio) return { success: false, error: "Prédio(s) obrigatório(s)." };
    var arrPredios = Array.isArray(idsPredio) ? idsPredio : [idsPredio];

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetUnits = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheetUnits || sheetUnits.getLastRow() < 2) return { success: false, error: "Nenhuma fração configurada." };

    var unitsData = sheetUnits.getDataRange().getValues();
    var sheetQuotas = ss.getSheetByName(QUOTAS_CC_TAB);
    if (!sheetQuotas) {
      sheetQuotas = ss.insertSheet(QUOTAS_CC_TAB);
      sheetQuotas.appendRow(QUOTAS_CC_HEADERS);
    }

    var criadas = 0;
    var now = new Date();
    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    var mesRef = now.getMonth() + 1;
    var anoRef = now.getFullYear();
    var vencimento = dataVencimento ? new Date(dataVencimento) : new Date(now.getTime() + 15 * 24 * 3600 * 1000);

    arrPredios.forEach(function (idPredio) {
      for (var i = 1; i < unitsData.length; i++) {
        var row = unitsData[i];
        if (String(row[_UNIT_COL_ID_PREDIO] || "").trim() !== String(idPredio).trim()) continue;
        if (String(row[_UNIT_COL_STATUS] || "").trim() !== "Ativo") continue;

        // Se especificou idFracao, pular os outros. Se idFracao for null, cria para todos do prédio.
        if (idFracao && String(row[_UNIT_COL_ID] || "").trim() !== String(idFracao).trim()) continue;

        var idQuota = "QE-" + Utilities.getUuid().substring(0, 8).toUpperCase();
        sheetQuotas.appendRow([
          idQuota, idPredio, row[_UNIT_COL_ID], row[_UNIT_COL_DESIGNACAO], row[_UNIT_COL_PROPRIETARIO],
          row[_UNIT_COL_EMAIL_PROP], row[_UNIT_COL_TELEFONE], mesRef, anoRef, parseFloat(valor) || 0,
          "Extra: " + String(descricao).trim(), "Pendente", vencimento, "", "", String(descricao).trim(), now, userEmail
        ]);
        criadas++;
      }
    });

    _auditCondominioAction(ss, userEmail, "Quotas", arrPredios.join(","), "ADICIONAR_QUOTA_EXTRA", "", "", "Valor: " + valor + " | Criadas: " + criadas);
    return { success: true, message: "✅ " + criadas + " quota(s) extra(s) criada(s) com sucesso." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Obtém a conta corrente de quotas de um prédio (ou todas se idPredio=null).
 * Suporta filtro por status e mês.
 */
function getQuotasCC(idsPredio, filtros, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(QUOTAS_CC_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, quotas: [], totais: { pendente: 0, pago: 0, totalPendente: 0, totalPago: 0 } };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, QUOTAS_CC_HEADERS.length).getValues();
    filtros = filtros || {}; // Ensure filtros object exists
    var quotas = data.map(function (row, idx) {
      row.push(idx + 1); // Row index for internal use
      return row;
    }).filter(function (r) {
      if (idsPredio) {
        var arrPredios = Array.isArray(idsPredio) ? idsPredio.map(String) : [String(idsPredio)];
        if (arrPredios.indexOf(String(r[_QUOTA_COL_ID_PREDIO] || "")) === -1) return false;
      }
      if (filtros.status && String(r[_QUOTA_COL_STATUS] || "").trim() !== filtros.status) return false;
      
      // Filtros de Data Múltiplos
      if (filtros.meses && filtros.meses.length > 0) {
        if (filtros.meses.indexOf(String(r[_QUOTA_COL_MES])) === -1) return false;
      } else if (filtros.mes) {
        if (String(r[_QUOTA_COL_MES]) !== String(filtros.mes)) return false;
      }
      
      if (filtros.anos && filtros.anos.length > 0) {
        if (filtros.anos.indexOf(String(r[_QUOTA_COL_ANO])) === -1) return false;
      } else if (filtros.ano) {
        if (String(r[_QUOTA_COL_ANO]) !== String(filtros.ano)) return false;
      }

      if (filtros.dias && filtros.dias.length > 0) {
        var venc = r[_QUOTA_COL_VENCIMENTO];
        if (!venc) return false;
        var dia = new Date(venc).getDate();
        if (filtros.dias.indexOf(String(dia)) === -1) return false;
      }
      return true;
    }).map(function (r) {
      var venc = r[_QUOTA_COL_VENCIMENTO];
      var dataPag = r[_QUOTA_COL_DATA_PAG];
      var hoje = new Date();
      var status = String(r[_QUOTA_COL_STATUS] || "Pendente");
      var emAtraso = status === "Pendente" && venc && new Date(venc) < hoje;
      return {
        idQuota: r[_QUOTA_COL_ID], idPredio: r[_QUOTA_COL_ID_PREDIO], idFracao: r[_QUOTA_COL_ID_FRACAO],
        designacao: r[_QUOTA_COL_DESIGNACAO], proprietario: r[_QUOTA_COL_PROP],
        email: r[_QUOTA_COL_EMAIL], telefone: r[_QUOTA_COL_TELEFONE],
        mes: r[_QUOTA_COL_MES], ano: r[_QUOTA_COL_ANO],
        valor: parseFloat(r[_QUOTA_COL_VALOR]) || 0, tipo: r[_QUOTA_COL_TIPO],
        status: status, emAtraso: emAtraso,
        dataVencimento: venc ? Utilities.formatDate(new Date(venc), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        dataPagamento: dataPag ? Utilities.formatDate(new Date(dataPag), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        metodoPagamento: r[_QUOTA_COL_METODO], notas: r[_QUOTA_COL_NOTAS]
      };
    });

    var totais = { pendente: 0, pago: 0, totalPendente: 0, totalPago: 0, emAtraso: 0, valorAtraso: 0 };
    quotas.forEach(function (q) {
      if (q.status === "Pendente") { totais.pendente++; totais.totalPendente += q.valor; if (q.emAtraso) { totais.emAtraso++; totais.valorAtraso += q.valor; } }
      else if (q.status === "Pago") { totais.pago++; totais.totalPago += q.valor; }
    });
    totais.totalPendente = Math.round(totais.totalPendente * 100) / 100;
    totais.totalPago = Math.round(totais.totalPago * 100) / 100;
    totais.valorAtraso = Math.round(totais.valorAtraso * 100) / 100;

    return { success: true, quotas: quotas, totais: totais };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Registar pagamento de uma quota.
 */
function pagarQuota(idQuota, metodoPagamento, notas, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idQuota) return { success: false, error: "idQuota obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(QUOTAS_CC_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Quota não encontrada." };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_QUOTA_COL_ID] || "").trim() === String(idQuota).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Quota não encontrada: " + idQuota };
    if (String(data[rowIndex][_QUOTA_COL_STATUS] || "").trim() === "Pago") return { success: false, error: "Quota já está paga." };

    var sr = rowIndex + 1;
    sheet.getRange(sr, _QUOTA_COL_STATUS + 1).setValue("Pago");
    sheet.getRange(sr, _QUOTA_COL_DATA_PAG + 1).setValue(new Date());
    sheet.getRange(sr, _QUOTA_COL_METODO + 1).setValue(String(metodoPagamento || "Transferência").trim());
    if (notas) sheet.getRange(sr, _QUOTA_COL_NOTAS + 1).setValue(String(notas).trim());

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Quota", idQuota, "PAGAR_QUOTA", "Pendente", "Pago", metodoPagamento || "");
    return { success: true, message: "Quota " + idQuota + " marcada como paga." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Obtém lista de devedores (frações com quotas PENDENTES + em atraso) de um prédio.
 */
function getDevedoresCondominios(idsPredio, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;

    var res = getQuotasCC(idsPredio, { status: "Pendente" }, impersonateTarget);
    if (!res.success) return res;

    // Obter nomes dos prédios para o sumário
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
    var prediosNomes = {};
    if (sheetB && sheetB.getLastRow() > 1) {
      var bData = sheetB.getDataRange().getValues();
      bData.forEach(function(r) { prediosNomes[String(r[_BLDG_COL_ID])] = String(r[_BLDG_COL_NOME]); });
    }

    var devedores = {};
    res.quotas.filter(function (q) { return q.emAtraso; }).forEach(function (q) {
      var key = q.idFracao;
      if (!devedores[key]) {
        devedores[key] = { idFracao: q.idFracao, designacao: q.designacao, proprietario: q.proprietario, email: q.email, telefone: q.telefone, totalDivida: 0, quotasAtraso: 0, idPredio: q.idPredio, predioNome: prediosNomes[q.idPredio] || q.idPredio };
      }
      devedores[key].totalDivida += q.valor;
      devedores[key].quotasAtraso++;
    });

    var lista = Object.values(devedores).map(function (d) {
      d.totalDivida = Math.round(d.totalDivida * 100) / 100;
      return d;
    });
    lista.sort(function (a, b) { return b.totalDivida - a.totalDivida; });

    return { success: true, devedores: lista, totalGlobalDivida: Math.round(lista.reduce(function (s, d) { return s + d.totalDivida; }, 0) * 100) / 100 };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 📲 NOTIFICAÇÕES — SMS e EMAIL
// ==========================================

/**
 * Envia notificação de cobrança por Email a devedores de um prédio (manual).
 * idFracao: se fornecido, envia só a essa fração. Caso contrário, envia a todos os devedores.
 */
function enviarNotificacaoEmail(idsPredio, idFracao, mensagemCustom, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idsPredio) return { success: false, error: "Prédio(s) obrigatório(s)." };
    var arrPredios = Array.isArray(idsPredio) ? idsPredio : [idsPredio];

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
    var prediosNomes = {};
    if (sheetB && sheetB.getLastRow() > 1) {
      var bData = sheetB.getDataRange().getValues();
      bData.forEach(function(r) { prediosNomes[String(r[_BLDG_COL_ID])] = String(r[_BLDG_COL_NOME]); });
    }

    var devRes = getDevedoresCondominios(arrPredios, impersonateTarget);
    if (!devRes.success) return devRes;

    var targets = idFracao
      ? devRes.devedores.filter(function (d) { return String(d.idFracao) === String(idFracao); })
      : devRes.devedores;

    if (targets.length === 0) return { success: true, enviados: 0, message: "Nenhum devedor encontrado." };

    var enviados = 0; var erros = [];
    targets.forEach(function (d) {
      if (!d.email || !d.email.includes("@")) { erros.push(d.designacao + ": sem email"); return; }
      try {
        var nomePredio = d.predioNome || "Condomínio";
        var innerBody;
        if (mensagemCustom) {
            innerBody = "<div style='color:#475569;font-size:14px;line-height:1.7;'>" + 
                        mensagemCustom.replace(/{nome}/g, d.proprietario)
                                      .replace(/{fracao}/g, d.designacao)
                                      .replace(/{divida}/g, "€" + d.totalDivida.toFixed(2))
                                      .replace(/{predio}/g, nomePredio)
                                      .replace(/\n/g, '<br>') + 
                        "</div>";
        } else {
            innerBody = "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Exmo(a) Condómino <strong>" + d.proprietario + "</strong>,</p>" +
                        "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Verificamos que existem quotas em atraso referentes à sua fração <strong>" + d.designacao + "</strong> no <strong>" + nomePredio + "</strong>.</p>" +
                        "<div style='background:#FEF2F2;border-left:4px solid #EF4444;padding:16px;margin-bottom:24px;'>" +
                        "<p style='color:#991B1B;font-size:16px;font-weight:bold;margin:0;'>Valor em dívida: €" + d.totalDivida.toFixed(2) + "</p>" +
                        "<p style='color:#B91C1C;font-size:13px;margin:4px 0 0;'>" + d.quotasAtraso + " quota(s) em atraso</p>" +
                        "</div>" +
                        "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;'>Pedimos que regularize a situação o mais breve possível.</p>" +
                        "<p style='color:#475569;font-size:14px;line-height:1.7;margin:0;'>Para mais informações, contacte a administração do condomínio.</p>";
        }
        var htmlBody = _buildStandardEmailHTML("Quotas em Atraso", innerBody);
        let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
        var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
        GmailApp.sendEmail(d.email, "[" + nomePredio + "] Quotas em Atraso — Fração " + d.designacao, "", options);
        enviados++;
      } catch (err) { erros.push(d.designacao + ": " + err.message); }
    });

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Notificacao", arrPredios.join(","), "ENVIAR_EMAIL_DEVEDORES", "", "", "Enviados: " + enviados + " | Erros: " + erros.length);
    return { success: true, enviados: enviados, erros: erros, message: "📧 " + enviados + " email(s) enviado(s)" + (erros.length > 0 ? " | ⚠️ " + erros.length + " erro(s)" : "") };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Sumário Semanal — envia relatório de quotas/devedores ao gestor por email.
 */
function enviarSumarioSemanal(idsPredio, emailGestor, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idsPredio || !emailGestor) return { success: false, error: "Prédios e emailGestor são obrigatórios." };
    var arrPredios = Array.isArray(idsPredio) ? idsPredio : [idsPredio];

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var hoje = new Date();
    var mesAtual = hoje.getMonth() + 1;
    var anoAtual = hoje.getFullYear();

    var devRes = getDevedoresCondominios(arrPredios, impersonateTarget);
    var quotasRes = getQuotasCC(arrPredios, { meses: [String(mesAtual)], anos: [String(anoAtual)] }, impersonateTarget);

    var devedores = devRes.success ? devRes.devedores : [];
    var totais = quotasRes.success ? quotasRes.totais : { pendente: 0, pago: 0, totalPendente: 0, totalPago: 0, emAtraso: 0, valorAtraso: 0 };

    var tableBody = devedores.length > 0
      ? devedores.map(function(d) { return "<tr><td style='border-bottom:1px solid #E2E8F0;padding:8px;'>" + d.designacao + "<br><span style='font-size:11px;color:#94A3B8;'>" + (d.predioNome || d.idPredio) + "</span></td><td style='border-bottom:1px solid #E2E8F0;padding:8px;'>" + d.proprietario + "</td><td style='border-bottom:1px solid #E2E8F0;padding:8px;color:#EF4444;font-weight:bold;'>€" + d.totalDivida.toFixed(2) + " (" + d.quotasAtraso + ")</td></tr>"; }).join("")
      : "<tr><td colspan='3' style='padding:12px;text-align:center;color:#10B981;font-weight:bold;background:#ECFDF5;'>Nenhum devedor em atraso. ✅</td></tr>";

    var innerBody = "<div style='background:#F8FAFC;padding:16px;border-radius:8px;margin-bottom:24px;border:1px solid #E2E8F0;'>" +
                    "<p style='margin:0 0 8px;font-size:13px;color:#64748B;'><strong>Gerado em:</strong> " + Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + "</p>" +
                    "<p style='margin:0;font-size:13px;color:#64748B;'><strong>Prédios Analisados:</strong> " + arrPredios.length + "</p>" +
                    "</div>" +
                    "<h3 style='color:#0F172A;font-size:16px;margin:0 0 16px;'>Mês Atual (" + (["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][mesAtual-1]) + "/" + anoAtual + ")</h3>" +
                    "<ul style='color:#475569;font-size:14px;line-height:1.7;margin:0 0 32px;padding-left:20px;'>" +
                    "<li><strong>Quotas pagas:</strong> " + totais.pago + " (€" + totais.totalPago.toFixed(2) + ")</li>" +
                    "<li><strong>Quotas pendentes:</strong> " + totais.pendente + " (€" + totais.totalPendente.toFixed(2) + ")</li>" +
                    "<li><strong>Em atraso:</strong> " + totais.emAtraso + " <span style='color:#EF4444;'>(€" + (totais.valorAtraso || 0).toFixed(2) + ")</span></li>" +
                    "</ul>" +
                    "<h3 style='color:#0F172A;font-size:16px;margin:0 0 16px;'>Devedores Atuais</h3>" +
                    "<table style='width:100%;border-collapse:collapse;font-size:13px;color:#334155;margin-bottom:24px;'>" +
                    "<thead><tr><th style='border-bottom:2px solid #CBD5E1;padding:8px;text-align:left;'>Fração / Prédio</th><th style='border-bottom:2px solid #CBD5E1;padding:8px;text-align:left;'>Proprietário</th><th style='border-bottom:2px solid #CBD5E1;padding:8px;text-align:left;'>Dívida</th></tr></thead>" +
                    "<tbody>" + tableBody + "</tbody>" +
                    "</table>" +
                    "<p style='color:#94A3B8;font-size:11px;margin:0;border-top:1px solid #E2E8F0;padding-top:16px;'>Este email foi gerado automaticamente pelo Flowly Condomínios.</p>";

    var htmlBody = _buildStandardEmailHTML("Sumário Semanal", innerBody);
    let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
    var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
    GmailApp.sendEmail(emailGestor, "[Flowly] Sumário Semanal Multi-Prédio — " + Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy"), "", options);

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Sumario", arrPredios.join(","), "ENVIAR_SUMARIO_SEMANAL", "", "", "Enviado para: " + emailGestor);
    return { success: true, message: "📧 Sumário semanal enviado para " + emailGestor };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Cancela/anula uma quota (marca como Cancelado).
 */
function cancelarQuota(idQuota, motivo, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;
    if (!idQuota) return { success: false, error: "idQuota obrigatório." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(QUOTAS_CC_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Quota não encontrada." };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_QUOTA_COL_ID] || "").trim() === String(idQuota).trim()) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, error: "Quota não encontrada." };

    var sr = rowIndex + 1;
    sheet.getRange(sr, _QUOTA_COL_STATUS + 1).setValue("Cancelado");
    if (motivo) sheet.getRange(sr, _QUOTA_COL_NOTAS + 1).setValue(String(motivo).trim());

    var userEmail = Session.getActiveUser().getEmail() || impersonateTarget || "desconhecido";
    _auditCondominioAction(ss, userEmail, "Quota", idQuota, "CANCELAR_QUOTA", "Pendente", "Cancelado", motivo || "");
    return { success: true, message: "Quota " + idQuota + " cancelada." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _auditCondominioAction(ss, utilizador, tipoEntidade, idEntidade, acao, valorAnterior, valorNovo, motivo) {
  try {
    var sheetAudit = ss.getSheetByName(AUDIT_DB_TAB);
    if (!sheetAudit) {
      sheetAudit = ss.insertSheet(AUDIT_DB_TAB);
      sheetAudit.appendRow(AUDIT_HEADERS.concat(["Tipo_Entidade", "Valor_Anterior", "Valor_Novo"]));
      sheetAudit.getRange(1, 1, 1, AUDIT_HEADERS.length + 3).setFontWeight("bold").setBackground("#E2E8F0");
    }
    sheetAudit.appendRow([new Date(), utilizador, idEntidade, acao, motivo || "", tipoEntidade, valorAnterior || "", valorNovo || ""]);
  } catch (e) { }
}

function getUnitsSafe(idPredio, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    var guard = _assertCondominium(ctx);
    if (guard.success === false) return guard;

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(UNITS_DB_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, units: [] };

    _ensureUnitsIdEntidadeCol(sheet);

    var lastCol = Math.max(sheet.getLastColumn(), 12);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    var units = data.filter(function (row) {
      if (idPredio) return String(row[_UNIT_COL_ID_PREDIO] || "").trim() === String(idPredio).trim();
      return true;
    }).map(function (row) {
      return {
        ID_Fracao: row[_UNIT_COL_ID], ID_Predio: row[_UNIT_COL_ID_PREDIO], Designacao: row[_UNIT_COL_DESIGNACAO],
        Proprietario: row[_UNIT_COL_PROPRIETARIO], NIF_Proprietario: _maskSensitiveField(String(row[_UNIT_COL_NIF_PROP] || ""), 2, 2),
        Email_Proprietario: row[_UNIT_COL_EMAIL_PROP], Telefone: row[_UNIT_COL_TELEFONE], Quota_Mensal: row[_UNIT_COL_QUOTA],
        Quinhao_Perc: row[_UNIT_COL_QUINHAO], Status: row[_UNIT_COL_STATUS], Data_Registo: row[_UNIT_COL_DATA_REGISTO],
        ID_Entidade: row[_UNIT_COL_ID_ENTIDADE] || ""
      };
    });

    return { success: true, units: units };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _calcCondominiumFinancials(ss, dataLog, start, end) {
  var result = { byBuilding: [], totalGlobalCosts: 0, rateioPorPredio: 0, activeBuildingsCount: 0 };
  try {
    var sheetB = ss.getSheetByName(BUILDINGS_DB_TAB);
    if (!sheetB || sheetB.getLastRow() < 2) return result;

    var bData = sheetB.getRange(2, 1, sheetB.getLastRow() - 1, BUILDINGS_HEADERS.length).getValues();
    var buildings = bData.filter(function (r) { return String(r[7] || "").trim() === "Ativo"; });
    result.activeBuildingsCount = buildings.length;
    if (buildings.length === 0) return result;

    var acc = {};
    buildings.forEach(function (b) {
      var id = String(b[_BLDG_COL_ID] || "").trim();
      if (!id) return;
      acc[id] = { idPredio: id, nomePredio: String(b[_BLDG_COL_NOME] || "").trim(), receitas: 0, custosDiretos: 0 };
    });

    var totalGlobal = 0;

    if (dataLog && dataLog.length > 0) {
      dataLog.forEach(function (r) {
        var d = r[0];
        if (typeof d === "string") { var p = d.split("/"); if (p.length === 3) d = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
        if (!(d instanceof Date) || isNaN(d)) return;
        if (!isDateInRange(d, start, end)) return;

        var tipo = normalizeTipo(r[1]);
        var qty = parseFloat(r[5]) || 0;
        var cost = parseFloat(String(r[6] || "0").replace(",", ".")) || 0;
        var sell = parseFloat(String(r[7] || "0").replace(",", ".")) || 0;
        var idEntidade = String(r[24] || "").trim();
        var colZ = String(r[25] || "").trim();

        if (colZ === "GLOBAL_BUILDINGS") { totalGlobal += qty * (cost || sell); return; }
        if (!idEntidade || !acc[idEntidade]) return;

        if (tipo === "saida" || tipo === "fechocaixa") { acc[idEntidade].receitas += qty * sell; }
        else if (tipo === "entrada" || tipo === "despesas" || tipo === "despesa" || tipo === "consumo") { acc[idEntidade].custosDiretos += qty * (cost || sell); }
      });
    }

    result.totalGlobalCosts = Math.round(totalGlobal * 100) / 100;
    var rateioPorPredio = buildings.length > 0 ? totalGlobal / buildings.length : 0;
    result.rateioPorPredio = Math.round(rateioPorPredio * 100) / 100;

    var ids = Object.keys(acc);
    result.byBuilding = ids.map(function (id) {
      var b = acc[id];
      var lucro = b.receitas - b.custosDiretos - rateioPorPredio;
      return { idPredio: b.idPredio, nomePredio: b.nomePredio, receitas: Math.round(b.receitas * 100) / 100, custosDiretos: Math.round(b.custosDiretos * 100) / 100, fatiaRateada: Math.round(rateioPorPredio * 100) / 100, lucro: Math.round(lucro * 100) / 100 };
    });

  } catch (e) { }
  return result;
}

function revelarDadoSensivel(tipo, entidade, id, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Contexto de cliente não encontrado." };

    var guard = _assertCondominium(ctx);
    if (!guard.success) return guard;

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var activeUser = Session.getActiveUser().getEmail();
    var valorRevelado = '';

    if (entidade === 'predio') {
      var shPredios = ss.getSheetByName('Predios_DB') || ss.getSheetByName(BUILDINGS_DB_TAB);
      if (!shPredios) return { success: false, error: "Aba de prédios não encontrada." };
      var data = shPredios.getDataRange().getValues();
      var row = data.find(r => r[0] === id);
      if (!row) return { success: false, error: "Prédio não encontrado." };
      if (tipo === 'nif') valorRevelado = String(row[3] || '');
      else if (tipo === 'iban') valorRevelado = String(row[4] || '');
      else return { success: false, error: "Tipo de dado inválido para prédio." };
    } else if (entidade === 'fracao') {
      var shFracoes = ss.getSheetByName('Fracoes_DB') || ss.getSheetByName(UNITS_DB_TAB);
      if (!shFracoes) return { success: false, error: "Aba de frações não encontrada." };
      var data = shFracoes.getDataRange().getValues();
      var row = data.find(r => r[0] === id);
      if (!row) return { success: false, error: "Fração não encontrada." };
      if (tipo === 'email') valorRevelado = String(row[5] || '');
      else return { success: false, error: "Tipo de dado inválido para fração." };
    } else { return { success: false, error: "Entidade inválida." }; }

    var shAudit = ss.getSheetByName(AUDIT_DB_TAB);
    if (!shAudit) {
      shAudit = ss.insertSheet(AUDIT_DB_TAB);
      shAudit.appendRow(['Timestamp', 'User', 'Ação', 'Entidade', 'ID', 'Tipo Dado', 'IP']);
    }

    var timestamp = new Date();
    var acao = 'Visualização de Dado Sensível';
    var ip = activeUser || ctx.clientEmail || 'Desconhecido';

    shAudit.appendRow([timestamp, activeUser || ctx.clientEmail || 'Desconhecido', acao, entidade, id, tipo.toUpperCase(), ip]);
    return { success: true, value: valorRevelado };
  } catch (e) { return { success: false, error: e.toString() }; }
}