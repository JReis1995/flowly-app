// Ficheiro: MOD_CC_Dashboard.js
/// ==========================================
// 📊 MÓDULO DE DASHBOARD E CONTA CORRENTE
// ==========================================

function getMarginAlerts(impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget || null);
    if (!ctx.sheetId) return [];

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - 99);
    const data = sheet.getRange(startRow, 1, lastRow, 26).getValues();
    const alerts = [];

    data.forEach(function (r) {
      if (normalizeTipo(r[1]) !== "saida") return;
      const precoReal = parseFloat(String(r[7] || "").replace(",", ".")) || 0;
      const precoSugerido = parseFloat(String(r[25] || "").replace(",", ".")) || 0;

      if (precoSugerido <= 0 || precoReal >= precoSugerido * 0.95) return;

      const desvioPct = ((precoReal - precoSugerido) / precoSugerido) * 100;
      alerts.push({
        data: r[0],
        artigo: (r[4] || "").toString().trim(),
        precoReal: precoReal,
        precoSugerido: precoSugerido,
        desvioPct: desvioPct
      });
    });

    alerts.sort(function (a, b) { return a.desvioPct - b.desvioPct; });
    return alerts.slice(0, 10);
  } catch (e) {
    return [];
  }
}

function saveMarginSettings(margemDesejada, ircEstimado, metodoCalculo) {
  try {
    const m = parseFloat(margemDesejada);
    const i = parseFloat(ircEstimado);
    if (isNaN(m) || isNaN(i) || m < 0 || i < 0 || m > 100 || i > 100) {
      return { success: false, error: "Valores inválidos. Margem e IRC devem ser números entre 0 e 100." };
    }
    const props = PropertiesService.getDocumentProperties();
    const toSet = {
      MARGEM_DESEJADA: String(m),
      IRC_ESTIMADO: String(i)
    };
    if (metodoCalculo !== undefined && metodoCalculo !== null) {
      toSet.METODO_CALCULO = (metodoCalculo === "margem_real") ? "margem_real" : "markup";
    }
    props.setProperties(toSet);
    return { success: true, message: "Preferências de margem guardadas com sucesso." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getMarginSettings() {
  try {
    const props = PropertiesService.getDocumentProperties().getProperties();
    const margemDesejada = parseFloat(props.MARGEM_DESEJADA);
    const ircEstimado = parseFloat(props.IRC_ESTIMADO);
    const metodoCalculo = (props.METODO_CALCULO === "margem_real") ? "margem_real" : "markup";
    const autoOpExRate = parseFloat(props.AUTO_OPEX_RATE) || 0;
    return {
      margemDesejada: (!isNaN(margemDesejada) && margemDesejada >= 0) ? margemDesejada : 30,
      ircEstimado: (!isNaN(ircEstimado) && ircEstimado >= 0) ? ircEstimado : 20,
      metodoCalculo: metodoCalculo,
      autoOpExRate: autoOpExRate
    };
  } catch (e) {
    return { margemDesejada: 30, ircEstimado: 20, metodoCalculo: "markup", autoOpExRate: 0 };
  }
}

function calculateDynamicOpEx(impersonateTarget, startDate, endDate) {
  const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  const emptyResult = { rate: 0, label: "365d", ratePersisted: false };
  try {
    const ctx = getClientContext(impersonateTarget || null);
    if (!ctx.sheetId) {
      PropertiesService.getDocumentProperties().setProperty("AUTO_OPEX_RATE", "0");
      return emptyResult;
    }
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    const sheetStaff = getStaffSheet(ss);

    let totalRH = 0;
    if (sheetStaff && sheetStaff.getLastRow() >= 2) {
      const lastRowStaff = sheetStaff.getLastRow();
      const dataStaff = sheetStaff.getRange(2, 1, lastRowStaff, 14).getValues();
      for (let i = 0; i < dataStaff.length; i++) {
        if (String(dataStaff[i][8] || "").trim() === "Ativo") {
          const mensal = parsePTFloat(dataStaff[i][9]);
          const rescisao = parsePTFloat(dataStaff[i][12]);
          const formacao = parsePTFloat(dataStaff[i][13]);
          totalRH += (mensal + rescisao + formacao);
        }
      }
    }

    let start, end, diasNoPeriodo, label, usePeriodFilter;
    if (startDate && endDate && typeof startDate === "string" && typeof endDate === "string") {
      const mStart = startDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      start = mStart ? new Date(parseInt(mStart[1], 10), parseInt(mStart[2], 10) - 1, parseInt(mStart[3], 10)) : null;
      const mEnd = endDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      end = mEnd ? new Date(parseInt(mEnd[1], 10), parseInt(mEnd[2], 10) - 1, parseInt(mEnd[3], 10)) : null;
      if (!start || !end || isNaN(start.getTime()) || isNaN(end.getTime())) {
        start = null; end = null;
      }
    }

    if (start && end) {
      diasNoPeriodo = Math.max(1, Math.round((end.getTime() - start.getTime()) / 86400000) + 1);
      usePeriodFilter = true;
      if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
        label = (diasNoPeriodo === 1) ? (start.getDate() + " " + MESES[start.getMonth()]) : (start.getDate() + "-" + end.getDate() + " " + MESES[end.getMonth()]);
      } else {
        label = start.getDate() + "/" + (start.getMonth() + 1) + "-" + end.getDate() + "/" + (end.getMonth() + 1);
      }
    } else {
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - 365);
      cutoff.setHours(0, 0, 0, 0);
      start = cutoff;
      end = new Date();
      diasNoPeriodo = 365;
      label = "365d";
      usePeriodFilter = false;
    }

    const rhProporcional = (totalRH / 30) * diasNoPeriodo;

    let totalEntradas = 0;
    let totalOutrasDespesas = 0;
    if (sheetNew && sheetNew.getLastRow() >= 2) {
      const data = sheetNew.getRange(2, 1, sheetNew.getLastRow(), 8).getValues();
      for (let i = 0; i < data.length; i++) {
        const d = parseDate(data[i][0]);
        if (!d || isNaN(d.getTime())) continue;
        if (usePeriodFilter && !isDateInRange(d, start, end)) continue;
        if (!usePeriodFilter && d < start) continue;
        const tipo = normalizeTipo(data[i][1]);
        if (tipo === "entrada") {
          totalEntradas += parsePTFloat(data[i][5]) * parsePTFloat(data[i][6]);
        } else if (tipo === "despesa" || tipo === "despesas" || tipo === "consumo") {
          totalOutrasDespesas += parsePTFloat(data[i][5]) * (parsePTFloat(data[i][6]) || parsePTFloat(data[i][7]));
        }
      }
    }

    const rate = totalEntradas > 0 ? ((rhProporcional + totalOutrasDespesas) / totalEntradas) * 100 : 0;
    const rateStr = rate.toFixed(2);
    const ratePersisted = (diasNoPeriodo >= 30 && totalEntradas > 0);

    if (ratePersisted) {
      PropertiesService.getDocumentProperties().setProperty("AUTO_OPEX_RATE", rateStr);
    }

    return { rate: parseFloat(rateStr), label: label, ratePersisted: ratePersisted };
  } catch (e) {
    PropertiesService.getDocumentProperties().setProperty("AUTO_OPEX_RATE", "0");
    return emptyResult;
  }
}

function saveMarginSettingsAndHistory(margem, irc, updateHistory, impersonateTarget, metodoCalculo, startDate, endDate) {
  try {
    const saved = saveMarginSettings(margem, irc, metodoCalculo);
    if (!saved.success) return saved;
    if (updateHistory !== true) {
      return { success: true, message: "Preferências de margem guardadas." };
    }

    const ctx = getClientContext(impersonateTarget || null);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };

    const opExResult = calculateDynamicOpEx(impersonateTarget || null, startDate || "", endDate || "");
    const autoOpExRateOverride = (opExResult && opExResult.rate != null) ? opExResult.rate : 0;

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, message: "Preferências guardadas. Sem dados para atualizar.", ratePersisted: opExResult && opExResult.ratePersisted };
    }

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, 26).getValues()[0].map(function (h) { return (h || "").toString().trim(); });
    const idxPrecoSugerido = headers.indexOf("Preço Sugerido");
    const colW = (idxPrecoSugerido >= 0) ? (idxPrecoSugerido + 1) : 26;
    const data = sheet.getRange(2, 1, lastRow, 26).getValues();
    const array2D = [];

    for (let i = 0; i < data.length; i++) {
      const tipo = normalizeTipo(data[i][1]);
      const custo = data[i][6];
      const valorExistente = (colW <= data[i].length) ? data[i][colW - 1] : data[i][25];
      const novoValor = (tipo === "entrada") ? calcularPrecoSugerido(custo, autoOpExRateOverride) : (valorExistente != null && valorExistente !== "" ? valorExistente : "");
      array2D.push([novoValor]);
    }

    if (array2D.length > 0) {
      sheet.getRange(2, colW, array2D.length, 1).setValues(array2D);
      SpreadsheetApp.flush();
    }
    return { success: true, message: "Estratégia aplicada e histórico atualizado.", ratePersisted: opExResult && opExResult.ratePersisted };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function calcularPrecoSugerido(precoCusto_PT, autoOpExRateOverride) {
  if (precoCusto_PT === null || precoCusto_PT === undefined || precoCusto_PT === "") return "";
  const strNorm = String(precoCusto_PT).trim().replace(",", ".");
  const custoFloat = parseFloat(strNorm);
  if (isNaN(custoFloat) || custoFloat < 0) return "";

  const { margemDesejada, ircEstimado, metodoCalculo, autoOpExRate } = getMarginSettings();
  const opEx = (autoOpExRateOverride !== undefined && autoOpExRateOverride !== null) ? (autoOpExRateOverride || 0) : (autoOpExRate || 0);
  const totalTaxas = margemDesejada + ircEstimado + opEx;

  let precoSugerido;
  if (metodoCalculo === "margem_real") {
    if (totalTaxas >= 100) return "";
    precoSugerido = custoFloat / (1 - (totalTaxas / 100));
  } else {
    precoSugerido = custoFloat * (1 + (totalTaxas / 100));
  }

  if (!isFinite(precoSugerido) || precoSugerido < 0) return "";
  return precoSugerido.toFixed(2).replace(".", ",");
}

function ensureCCSheetWithHeaders(ss) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) sheet = ss.getSheetByName("New");
  if (!sheet) sheet = ss.insertSheet(SHEET_TAB_NAME);
  if (sheet.getLastRow() < 1) sheet.getRange(1, 1, 1, CC_HEADERS.length).setValues([CC_HEADERS]);

  ensureColumn(sheet, 21, "DocID");
  ensureColumn(sheet, 22, "Matricula");
  ensureColumn(sheet, 23, "Km_Atuais");
  ensureColumn(sheet, 24, "Litros");
  ensureColumn(sheet, CC_COL_ID_ENTIDADE, "ID_Entidade");
  ensureColumn(sheet, 26, "Preço Sugerido");

  try {
    const headerRange = sheet.getRange(1, 1, 1, 26);
    const protection = headerRange.protect().setDescription('Proteção de Estrutura Flowly');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) { protection.setDomainEdit(false); }
  } catch (e) { }

  const statusRange = sheet.getRange(2, 17, 5000, 17);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Aberto', 'Pago'], true)
    .setHelpText('Selecione Aberto ou Pago')
    .build();
  statusRange.setDataValidation(rule);

  return sheet;
}

function setupSheetStructure(impersonateTarget, sheetTabName) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, duplicatesRemoved: 0, idEntidadeAdded: false, error: "Infraestrutura não encontrada. Use Modo Espião." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var tabName = (sheetTabName && String(sheetTabName).trim()) ? String(sheetTabName).trim() : "New";
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, duplicatesRemoved: 0, idEntidadeAdded: false, error: "Aba '" + tabName + "' não encontrada." };

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { success: false, duplicatesRemoved: 0, idEntidadeAdded: false, error: "Folha vazia." };

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return (h || "").toString().trim(); });
    var seen = {};
    var colsToDelete = [];

    for (var c = 0; c < headers.length; c++) {
      var h = headers[c];
      if (!h) continue;
      var key = h.toLowerCase();
      if (seen[key] !== undefined) colsToDelete.push(c + 1);
      else seen[key] = c;
    }

    for (var d = colsToDelete.length - 1; d >= 0; d--) {
      sheet.deleteColumn(colsToDelete[d]);
      lastCol--;
    }

    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return (h || "").toString().trim(); });
    var idxIdEntidade = _indexOfHeader(headers, ["ID_Entidade", "ID Entidade"]);
    var idEntidadeAdded = false;
    var headerAtColY = (headers[CC_COL_ID_ENTIDADE - 1] || "").toString().trim();

    if (headerAtColY !== "ID_Entidade" && headerAtColY !== "ID Entidade") {
      if (sheet.getLastColumn() < CC_COL_ID_ENTIDADE) {
        ensureColumn(sheet, CC_COL_ID_ENTIDADE, "ID_Entidade");
      } else {
        sheet.insertColumnBefore(CC_COL_ID_ENTIDADE);
        sheet.getRange(1, CC_COL_ID_ENTIDADE).setValue("ID_Entidade");
        if (idxIdEntidade >= 0 && idxIdEntidade + 1 > CC_COL_ID_ENTIDADE) {
          var srcCol = idxIdEntidade + 2;
          for (var r = 2; r <= sheet.getLastRow(); r++) {
            sheet.getRange(r, CC_COL_ID_ENTIDADE).setValue(sheet.getRange(r, srcCol).getValue());
          }
          sheet.deleteColumn(srcCol);
        }
      }
      idEntidadeAdded = true;
    }

    return { success: true, duplicatesRemoved: colsToDelete.length, idEntidadeAdded: idEntidadeAdded };
  } catch (e) {
    return { success: false, duplicatesRemoved: 0, idEntidadeAdded: false, error: e.toString() };
  }
}

function syncAllExistingIds(impersonateTarget, sheetTabName) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, updated: 0, skipped: 0, error: "Infraestrutura não encontrada." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var tabName = (sheetTabName && String(sheetTabName).trim()) ? String(sheetTabName).trim() : "New";
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, updated: 0, skipped: 0, error: "Aba '" + tabName + "' não encontrada." };

    ensureColumn(sheet, CC_COL_ID_ENTIDADE, "ID_Entidade");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, updated: 0, skipped: 0 };

    var headers = sheet.getRange(1, 1, 1, Math.max(CC_COL_ID_ENTIDADE, sheet.getLastColumn())).getValues()[0].map(function (h) { return (h || "").toString().trim(); });
    var idxFornecedor = _indexOfHeader(headers, ["Fornecedor", "fornecedor"]);
    if (idxFornecedor === -1) idxFornecedor = 3;
    var idxIdEntidade = _indexOfHeader(headers, ["ID_Entidade", "ID Entidade"]);
    if (idxIdEntidade === -1) idxIdEntidade = CC_COL_ID_ENTIDADE - 1;
    var idxTipo = _indexOfHeader(headers, ["Tipo", "tipo"]);
    if (idxTipo === -1) idxTipo = 1;

    var data = sheet.getRange(2, 1, lastRow, Math.max(idxFornecedor + 1, CC_COL_ID_ENTIDADE)).getValues();
    var updated = 0, skipped = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var fornecedor = (row[idxFornecedor] || "").toString().trim();
      var idAtual = (row[idxIdEntidade] || "").toString().trim();
      if (!fornecedor) { skipped++; continue; }
      if (idAtual) { skipped++; continue; }

      var tipo = (row[idxTipo] || "").toString().trim();
      var tipoLanc = (tipo === "Saida" || tipo === "Saída") ? "Saída" : "Despesa";
      var id = getEntityIdByName(fornecedor, tipoLanc, impersonateTarget || null);
      if (id) {
        sheet.getRange(i + 2, CC_COL_ID_ENTIDADE).setValue(id);
        updated++;
      }
    }
    return { success: true, updated: updated, skipped: skipped };
  } catch (e) {
    return { success: false, updated: 0, skipped: 0, error: e.toString() };
  }
}

function fixMissingEntityIds(impersonateTarget, sheetTabName) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, updated: 0, skipped: 0, error: "Infraestrutura não encontrada. Use Modo Espião." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var tabName = (sheetTabName && String(sheetTabName).trim()) ? String(sheetTabName).trim() : SHEET_TAB_NAME;
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, updated: 0, skipped: 0, error: "Aba '" + tabName + "' não encontrada." };

    ensureColumn(sheet, CC_COL_ID_ENTIDADE, "ID_Entidade");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, updated: 0, skipped: 0 };

    var headers = sheet.getRange(1, 1, 1, Math.max(CC_COL_ID_ENTIDADE, sheet.getLastColumn())).getValues()[0].map(function (h) { return (h || "").toString().trim(); });
    var idxFornecedor = _indexOfHeader(headers, ["Fornecedor", "fornecedor"]);
    if (idxFornecedor === -1) idxFornecedor = 3;
    var idxIdEntidade = _indexOfHeader(headers, ["ID_Entidade", "ID Entidade"]);
    if (idxIdEntidade === -1) idxIdEntidade = CC_COL_ID_ENTIDADE - 1;
    var idxTipo = _indexOfHeader(headers, ["Tipo", "tipo"]);
    if (idxTipo === -1) idxTipo = 1;

    var data = sheet.getRange(2, 1, lastRow, Math.max(idxFornecedor + 1, idxIdEntidade + 1)).getValues();
    var updated = 0, skipped = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var fornecedor = (row[idxFornecedor] || "").toString().trim();
      var idEntidadeAtual = (row[idxIdEntidade] || "").toString().trim();
      if (!fornecedor) { skipped++; continue; }
      if (idEntidadeAtual) { skipped++; continue; }

      var tipo = (row[idxTipo] || "").toString().trim();
      var tipoLanc = (tipo === "Saida" || tipo === "Saída") ? "Saída" : "Despesa";
      var id = getEntityIdByName(fornecedor, tipoLanc, impersonateTarget || null);
      if (id) {
        sheet.getRange(i + 2, CC_COL_ID_ENTIDADE).setValue(id);
        updated++;
      }
    }
    return { success: true, updated: updated, skipped: skipped };
  } catch (e) {
    return { success: false, updated: 0, skipped: 0, error: e.toString() };
  }
}

function processPayment(rowIndex, paymentVal, paymentDate, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_TAB_NAME);

    const pVal = parseFloat(String(paymentVal).replace(',', '.')) || 0;
    const currentPaid = parseFloat(sheet.getRange(rowIndex, 19).getValue()) || 0;
    const newPaid = currentPaid + pVal;

    const tipo = normalizeTipo(sheet.getRange(rowIndex, 2).getValue());
    const qty = parseFloat(sheet.getRange(rowIndex, 6).getValue()) || 0;
    const custo = parseFloat(sheet.getRange(rowIndex, 7).getValue()) || 0;
    const venda = parseFloat(sheet.getRange(rowIndex, 8).getValue()) || 0;

    const isEntradaOuDespesa = (tipo === "entrada" || tipo === "despesa" || tipo === "despesas" || tipo === "consumo");
    const invoiceTotal = isEntradaOuDespesa ? (qty * (custo || venda)) : (qty * venda);

    sheet.getRange(rowIndex, 19).setValue(newPaid);

    if (invoiceTotal > 0 && newPaid >= invoiceTotal) {
      sheet.getRange(rowIndex, 17).setValue("Pago");
      sheet.getRange(rowIndex, 18).setValue(paymentDate);
    } else {
      sheet.getRange(rowIndex, 17).setValue("Aberto");
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function processBulkPayments(docIds, meta) {
  try {
    if (!Array.isArray(docIds) || docIds.length === 0) return { success: false, error: "docIds deve ser um array não vazio." };

    const dataPagamento = (meta && meta.dataPagamento) ? String(meta.dataPagamento).trim() : "";
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dataPagamento)) {
      return { success: false, error: "meta.dataPagamento obrigatório no formato DD/MM/AAAA." };
    }

    const docIdsSet = {};
    docIds.forEach(function (id) {
      if (id && String(id).trim()) docIdsSet[String(id).trim()] = true;
    });
    if (Object.keys(docIdsSet).length === 0) return { success: false, error: "Nenhum DocID válido." };

    const ctx = getClientContext(meta.impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "ERRO: Infraestrutura não encontrada." };

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, updated: 0 };

    const raw = sheet.getRange(2, 1, lastRow, 21).getValues();
    const colK = [], colQ = [], colR = [], colS = [];
    let updated = 0;

    raw.forEach(function (r) {
      const docId = (r[20] != null) ? String(r[20]).trim() : "";
      const match = docId && docIdsSet[docId];

      const qty = parseFloat(r[5]) || 0;
      const custo = parseFloat(r[6]) || 0;
      const venda = parseFloat(r[7]) || 0;
      const tipo = normalizeTipo(r[1]) || "entrada";
      const isEntradaOuDespesa = (tipo === "entrada" || tipo === "despesa" || tipo === "despesas" || tipo === "consumo");
      const total = isEntradaOuDespesa ? (qty * (custo || venda)) : (qty * venda);

      if (match) {
        colK.push(["Sim"]);
        colQ.push(["Pago"]);
        colR.push([dataPagamento]);
        colS.push([total]);
        updated++;
      } else {
        colK.push([r[10] != null ? r[10] : ""]);
        colQ.push([r[16] != null ? r[16] : ""]);
        colR.push([r[17] != null ? r[17] : ""]);
        colS.push([r[18] != null ? r[18] : ""]);
      }
    });

    if (updated === 0) return { success: true, updated: 0 };

    const startRow = 2;
    for (let i = 0; i < raw.length; i++) {
      const row = startRow + i;
      sheet.getRange(row, 11).setValue(colK[i][0]);
      sheet.getRange(row, 17).setValue(colQ[i][0]);
      sheet.getRange(row, 18).setValue(colR[i][0]);
      sheet.getRange(row, 19).setValue(colS[i][0]);
    }

    return { success: true, updated: updated };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveValidatedData(p) {
  try {
    const ctx = getClientContext(p.impersonateEmail);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    const userEmail = Session.getActiveUser().getEmail();

    let formattedDate = p.data;
    if (formattedDate && String(formattedDate).match(/^\d{4}-\d{2}-\d{2}$/)) {
      const x = String(formattedDate).split('-');
      formattedDate = `${x[2]}/${x[1]}/${x[0]}`;
    }

    const docId = Utilities.getUuid();
    const jaPago = p.jaPago === true;
    const dataPagamento = (p.dataPagamento && String(p.dataPagamento).trim()) ? String(p.dataPagamento).trim() : "";
    let tipoCol = "Entrada";
    const t = String(p.tipo || "").toLowerCase().trim();
    if (t === "saída" || t === "saida") tipoCol = "Saida";
    else if (t === "quebra") tipoCol = "Quebra";
    else if (t === "oferta") tipoCol = "Oferta";
    else if (t === "consumo") tipoCol = "Consumo";
    else if (t === "fecho de caixa/relatório" || t === "fechocaixa" || t === "fecho caixa") tipoCol = "FechoCaixa";
    else if (t === "despesas") tipoCol = "Despesas";
    else if (p.tipo && String(p.tipo).trim()) tipoCol = String(p.tipo).trim();

    const isSaidaOuFecho = (tipoCol === "Saida" || tipoCol === "Saída" || tipoCol === "Quebra" || tipoCol === "Oferta" || tipoCol === "Consumo" || tipoCol === "FechoCaixa" || tipoCol === "Despesas" || String(tipoCol).toLowerCase().indexOf("fecho") >= 0);
    const isFechoCaixa = (tipoCol === "FechoCaixa" || String(tipoCol).toLowerCase().indexOf("fecho") >= 0);
    const fornecedorVal = isFechoCaixa ? "" : (p.fornecedor || "");
    const idEntidade = isFechoCaixa ? "" : getEntityIdByName(p.fornecedor || "", p.tipo || tipoCol, p.impersonateEmail);

    p.artigos.forEach(i => {
      const qtd = parseFloat(String(i.quantidade).replace(',', '.')) || 0;
      const custo = parseFloat(String(i.preco_custo).replace(',', '.')) || 0;
      const venda = parseFloat(String(i.preco_venda || '').replace(',', '.')) || 0;
      const taxa = normalizeTaxaIva(i.taxa_iva);
      const colG = isSaidaOuFecho ? 0 : custo;
      const colH = isSaidaOuFecho ? (venda || custo) : 0;
      const valorSugerido = !isSaidaOuFecho ? calcularPrecoSugerido(colG) : "";
      const valorBase = isSaidaOuFecho ? (venda || custo) : custo;
      const valIVA = (i.valor_iva != null && !isNaN(parseFloat(String(i.valor_iva).replace(",", ".")))) ? parseFloat(String(i.valor_iva).replace(",", ".")) : (qtd * valorBase) * (taxa / 100);
      const totalLinha = qtd * valorBase;
      const status = jaPago ? "Pago" : "Aberto";
      const valorPago = jaPago ? totalLinha : 0;
      const faturaPaga = jaPago ? "Sim" : "Não";
      const contaStock = (p.contaStock != null && String(p.contaStock).trim() !== "") ? String(p.contaStock).trim() : "Sim";

      const rowData = [
        formattedDate, tipoCol, p.metodo || "A Definir", fornecedorVal, (i.artigo || "").toString(),
        qtd, colG, colH, taxa + "%", valIVA, faturaPaga, "VALIDADO", p.observacoes || "", "",
        new Date(), userEmail, status, jaPago ? dataPagamento : "", valorPago, contaStock,
        docId, "", "", "", idEntidade, valorSugerido
      ];
      sheet.appendRow(rowData);
    });

    if (tipoCol === "Entrada") {
        try { if (typeof invalidateFTCostCache === "function") invalidateFTCostCache(p.impersonateEmail); } catch(e){}
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveDocumentToDB(tipoDocumento, cabecalho, linhas, fileUrl, impersonateTarget, opcoes) {
  try {
    const docId = Utilities.getUuid();
    let tipoCol = "Entrada";
    const t = String(tipoDocumento || "").toLowerCase();
    if (t === "saída" || t === "saida") tipoCol = "Saida";
    else if (t === "quebra") tipoCol = "Quebra";
    else if (t === "oferta") tipoCol = "Oferta";
    else if (t === "consumo") tipoCol = "Consumo";
    else if (t === "fecho de caixa/relatório" || t === "fechocaixa" || t === "fecho caixa") tipoCol = "FechoCaixa";
    else if (tipoDocumento === "FechoCaixa") tipoCol = "FechoCaixa";
    else if (tipoDocumento === "RelatorioSaidas") tipoCol = "Saida";
    else if (t === "despesas") tipoCol = "Despesas";
    else if (tipoDocumento === "Entrada" || t === "entrada") tipoCol = "Entrada";

    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "ERRO: Infraestrutura não encontrada." };

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date();
    const fornecedor = (cabecalho && cabecalho.fornecedor) ? cabecalho.fornecedor : (tipoCol === "Entrada" ? "" : "Caixa/Vendas");
    const nif = (cabecalho && cabecalho.nif) ? cabecalho.nif : "";
    const dataStr = (cabecalho && cabecalho.data) ? cabecalho.data : "";

    const jaPago = opcoes && opcoes.jaPago === true;
    const dataPagamento = (opcoes && opcoes.dataPagamento) ? String(opcoes.dataPagamento) : "";
    const contaStock = (opcoes && opcoes.contaStock != null) ? String(opcoes.contaStock).trim() : "Sim";

    const isSaidaOuFecho = (tipoCol === "Saida" || tipoCol === "Quebra" || tipoCol === "Oferta" || tipoCol === "Consumo" || tipoCol === "FechoCaixa" || tipoCol === "Despesas");

    for (let i = 0; i < linhas.length; i++) {
      const linha = linhas[i];
      const qty = parseFloat(linha.quantidade) || 1;
      const taxa = normalizeTaxaIva(linha.taxa_iva);
      let precoCusto = 0, precoVenda = 0;
      if (isSaidaOuFecho) {
        precoCusto = 0;
        precoVenda = parseFloat(linha.preco_venda) != null ? parseFloat(linha.preco_venda) : (parseFloat(linha.preco_custo) || 0);
      } else {
        precoCusto = parseFloat(linha.preco_custo) != null ? parseFloat(linha.preco_custo) : 0;
        precoVenda = 0;
      }

      const valorSugerido = !isSaidaOuFecho ? calcularPrecoSugerido(precoCusto) : "";
      const totalLinha = isSaidaOuFecho ? (qty * precoVenda) : (qty * precoCusto);
      const valorPagoLinha = jaPago ? totalLinha : 0;
      const status = jaPago ? "Pago" : "Aberto";
      const faturaPaga = jaPago ? "Sim" : "Não";
      const valorIvaCalc = (linha.valor_iva != null && !isNaN(parseFloat(linha.valor_iva))) ? parseFloat(linha.valor_iva) : (qty * (isSaidaOuFecho ? precoVenda : precoCusto)) * (taxa / 100);
      const idEnt = getEntityIdByName(fornecedor, tipoCol, impersonateTarget);

      const rowData = [
        dataStr, tipoCol, tipoCol === "Entrada" ? "A Definir" : "Caixa", fornecedor, (linha.artigo || "").toString(),
        qty, precoCusto, precoVenda, taxa + "%", valorIvaCalc, faturaPaga, "VALIDADO", "NIF: " + nif, fileUrl,
        now, userEmail, status, jaPago ? dataPagamento : "", valorPagoLinha, contaStock,
        docId, "", "", "", idEnt, valorSugerido
      ];
      sheet.appendRow(rowData);
    }
    if (tipoCol === "Entrada") {
        try { if (typeof invalidateFTCostCache === "function") invalidateFTCostCache(impersonateTarget); } catch(e){}
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveInvoiceToDB(cabecalhoOrInvoiceData, linhasOrFileUrl, fileUrlOrImpersonate, impersonateTarget) {
  try {
    let cabecalho, linhas, fileUrl, impersonate;
    if (arguments.length >= 4 && Array.isArray(linhasOrFileUrl)) {
      cabecalho = cabecalhoOrInvoiceData;
      linhas = linhasOrFileUrl;
      fileUrl = fileUrlOrImpersonate;
      impersonate = impersonateTarget;
    } else {
      const invoiceData = cabecalhoOrInvoiceData;
      fileUrl = linhasOrFileUrl;
      impersonate = fileUrlOrImpersonate;
      cabecalho = { nif: invoiceData.nif || "", fornecedor: invoiceData.fornecedor || "", data: invoiceData.data || "" };
      linhas = [{
        artigo: invoiceData.artigo || "",
        quantidade: 1,
        preco_custo: parseFloat(invoiceData.valor_base) || 0,
        taxa_iva: parseInt(invoiceData.taxa_iva) || 23,
        valor_iva: parseFloat(invoiceData.valor_iva) || 0
      }];
    }

    const ctx = getClientContext(impersonate);
    if (!ctx.sheetId) return { success: false, error: "ERRO: Infraestrutura não encontrada." };

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date();
    const docId = Utilities.getUuid();

    for (let i = 0; i < linhas.length; i++) {
      const linha = linhas[i];
      const qty = parseFloat(linha.quantidade) || 1;
      const preco = parseFloat(linha.preco_custo) != null ? parseFloat(linha.preco_custo) : 0;
      const valorSugerido = calcularPrecoSugerido(preco);
      const taxa = normalizeTaxaIva(linha.taxa_iva);
      const valorIva = parseFloat(linha.valor_iva) != null ? parseFloat(linha.valor_iva) : 0;
      const idEnt = getEntityIdByName(cabecalho.fornecedor || "", "Entrada", impersonate);

      const rowData = [
        cabecalho.data || "", "Entrada", "A Definir", cabecalho.fornecedor || "", (linha.artigo || "").toString(),
        qty, preco, "", taxa + "%", valorIva, "Sim", "VALIDADO", "NIF: " + (cabecalho.nif || ""), fileUrl,
        now, userEmail, "Aberto", "", 0, "Sim", docId, "", "", "", idEnt, valorSugerido
      ];
      sheet.appendRow(rowData);
    }
    try { if (typeof invalidateFTCostCache === "function") invalidateFTCostCache(impersonateTarget); } catch(e){}
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function _batchLoadDashboardSheets(impersonateTarget) {
  var out = { dataNew: [], dataFrota: [], dataVasilhame: [], dataAI: [], dataStaff: [], headersNew: [] };
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return out;

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    var sheetFrota = ss.getSheetByName(FROTA_VEICULOS_TAB);
    var sheetVasilhame = ss.getSheetByName(LOGISTICA_VASILHAME_TAB);
    var sheetAI = ss.getSheetByName(AI_HISTORY_TAB);
    var sheetStaff = getStaffSheet(ss);

    if (sheetNew && sheetNew.getLastRow() >= 2) {
      var fullNew = sheetNew.getRange(1, 1, sheetNew.getLastRow(), Math.max(26, sheetNew.getLastColumn())).getValues();
      out.headersNew = fullNew[0] ? fullNew[0].map(function (h) { return (h || "").toString().trim(); }) : [];
      out.dataNew = fullNew.slice(1);
    }
    if (sheetFrota && sheetFrota.getLastRow() >= 2) {
      out.dataFrota = sheetFrota.getRange(2, 1, sheetFrota.getLastRow(), Math.max(12, sheetFrota.getLastColumn())).getValues();
    }
    if (sheetVasilhame && sheetVasilhame.getLastRow() >= 2) {
      out.dataVasilhame = sheetVasilhame.getRange(2, 1, sheetVasilhame.getLastRow(), Math.max(7, sheetVasilhame.getLastColumn())).getValues();
    }
    if (sheetAI && sheetAI.getLastRow() >= 2) {
      out.dataAI = sheetAI.getRange(2, 1, sheetAI.getLastRow(), 4).getValues();
    }
    if (sheetStaff && sheetStaff.getLastRow() >= 2) {
      out.dataStaff = sheetStaff.getRange(2, 1, sheetStaff.getLastRow(), Math.max(20, sheetStaff.getLastColumn())).getValues();
    }
  } catch (e) { }
  return out;
}

function getUnifiedDashboardData(impersonateTarget, startDate, endDate, chartStartDate, chartEndDate, forceAI) {
  try {
    var batch = _batchLoadDashboardSheets(impersonateTarget);
    var lastAIInsight = _getLastAIInsightsFromBatch(batch.dataAI, impersonateTarget || Session.getActiveUser().getEmail() || "");
    var preloaded = { batch: batch, lastAIInsight: lastAIInsight };
    return getDashboardData(impersonateTarget, startDate, endDate, chartStartDate, chartEndDate, forceAI, preloaded);
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getDashboardData(impersonateTarget, startDate, endDate, chartStartDate, chartEndDate, forceAI, preloaded) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (ctx.isMaster && !ctx.sheetId) return { success: false };

    const email = impersonateTarget || Session.getActiveUser().getEmail() || "";
    const runAI = forceAI === true;
    const ss = SpreadsheetApp.openById(ctx.sheetId);

    const datesEmpty = !startDate && !endDate;
    const start = datesEmpty ? new Date(1970, 0, 1) : (startDate ? new Date(startDate) : new Date(new Date().getFullYear(), new Date().getMonth(), 1));
    const end = datesEmpty ? new Date(2099, 11, 31) : (endDate ? new Date(endDate) : new Date());

    const chartStart = chartStartDate ? new Date(chartStartDate) : start;
    const chartEnd = chartEndDate ? new Date(chartEndDate) : end;

    const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime();
    const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate()).getTime();
    const diasNoPeriodo = Math.max(1, Math.round((endDay - startDay) / 86400000) + 1);
    const now = new Date();
    const endIsCurrentMonth = (end.getMonth() === now.getMonth() && end.getFullYear() === now.getFullYear());
    const shouldProrateFixed = (diasNoPeriodo <= 7) || (diasNoPeriodo <= 31 && endIsCurrentMonth);

    let kpi = {
      fat: 0, stock_val: 0, rh: 0, iva_p: 0, iva_d: 0, compras: 0, despesas_opex: 0, transacoes: 0,
      profit_chart: {}, top_products: {}, iva_por_trimestre: {}, iva_por_trimestre_no_periodo: {}, stock_por_artigo: {}
    };
    let erosionCount = 0;
    let erosionItems = [];
    let chart_fat_by_day = {};
    let chart_desp_by_day = {};
    let chart_profit_by_day = {};

    let ftCostMap = {};
    try { if (typeof getFTCostMap === "function") ftCostMap = getFTCostMap(impersonateTarget); } catch(e){}

    const sheetLog = ss.getSheetByName(SHEET_TAB_NAME);
    const dataLog = (preloaded && preloaded.batch && preloaded.batch.dataNew && preloaded.batch.dataNew.length > 0)
      ? preloaded.batch.dataNew
      : (sheetLog && sheetLog.getLastRow() > 1 ? sheetLog.getRange(2, 1, sheetLog.getLastRow(), 26).getValues() : []);

    if (dataLog.length > 0) {
      const diasComSaida = {};
      dataLog.forEach(r => {
        const tipo = normalizeTipo(r[1]);
        const contaStock = (r[19] != null && String(r[19]).trim() !== "") ? String(r[19]).trim() : "Sim";
        if (tipo === "saida" && contaStock !== "Não") {
          let d = r[0];
          if (typeof d === 'string') { const p = d.split('/'); if (p.length === 3) d = new Date(`${p[2]}-${p[1]}-${p[0]}`); }
          if (d instanceof Date && !isNaN(d)) diasComSaida[d.toLocaleDateString('pt-PT')] = true;
        }
      });

      dataLog.forEach(r => {
        let d = r[0];
        if (typeof d === 'string') { const p = d.split('/'); if (p.length === 3) d = new Date(`${p[2]}-${p[1]}-${p[0]}`); }
        if (!(d instanceof Date) || isNaN(d)) return;

        const tipo = normalizeTipo(r[1]);
        const qty = parseFloat(r[5]) || 0;
        let cost = parseFloat(r[6]) || 0;
        const sell = parseFloat(r[7]) || 0;
        const artigo = (r[4] != null && r[4] !== '') ? String(r[4]).trim() : null;

        if (artigo && (tipo === "saida" || tipo === "fechocaixa")) {
           const lowArtigo = artigo.toLowerCase();
           if (ftCostMap[lowArtigo] !== undefined && ftCostMap[lowArtigo] > 0) {
                 cost = ftCostMap[lowArtigo];
           } else if (cost <= 0 && ftCostMap[lowArtigo] !== undefined) {
                 cost = ftCostMap[lowArtigo];
           }
        }

        const contaStock = (r[19] != null && String(r[19]).trim() !== "") ? String(r[19]).trim() : "Sim";
        const taxaRaw = parseFloat(String(r[8] || '').replace('%', '').trim());
        const taxa = isNaN(taxaRaw) ? 0 : (taxaRaw > 0 && taxaRaw <= 1 ? taxaRaw : taxaRaw / 100);
        const dedutivel = String(r[10] || '').trim().toLowerCase() === "sim";
        const valorIvaSheet = parsePTFloat(r[9]);

        const valor = (tipo === "entrada" ? qty * cost : (tipo === "saida" || tipo === "fechocaixa" ? qty * sell : 0));

        if (artigo) {
          if (!kpi.stock_por_artigo[artigo]) kpi.stock_por_artigo[artigo] = { qty: 0, valor: 0, lastEntryDate: null };
          if (tipo === "entrada") {
            kpi.stock_por_artigo[artigo].qty += qty;
            kpi.stock_por_artigo[artigo].valor += qty * cost;
            kpi.stock_por_artigo[artigo].lastEntryDate = d;
          }
          if (tipo === "saida" || tipo === "fechocaixa") {
            const st = kpi.stock_por_artigo[artigo];
            const stockQty = st.qty;
            const stockValor = st.valor;
            const costToUse = (cost > 0) ? cost : (stockQty > 0 ? stockValor / stockQty : 0);
            if (contaStock !== "Não" && sell > 0 && costToUse > 0) {
              const currentIdealPrice = calcularPrecoSugerido(costToUse);
              const sugeridoNum = parsePTFloat(currentIdealPrice);
              const vendaNum = parsePTFloat(r[7]);
              if (sugeridoNum > 0 && vendaNum > 0 && vendaNum < sugeridoNum) {
                erosionCount++;
                erosionItems.push({ artigo: artigo, venda: vendaNum, sugerido: sugeridoNum, data: d.toLocaleDateString ? d.toLocaleDateString("pt-PT") : String(d) });
              }
            }
            kpi.stock_por_artigo[artigo].qty -= qty;
            kpi.stock_por_artigo[artigo].valor -= qty * costToUse;
            if (kpi.stock_por_artigo[artigo].qty <= 0) kpi.stock_por_artigo[artigo].lastEntryDate = null;
          }
        }

        if (tipo === "entrada") kpi.stock_val += (qty * cost);
        if (tipo === "saida" || tipo === "fechocaixa") kpi.stock_val -= (qty * cost);

        const month = d.getMonth();
        const quarter = month <= 2 ? "T1" : (month <= 5 ? "T2" : (month <= 8 ? "T3" : "T4"));
        const yearKey = d.getFullYear() + "-" + quarter;
        const ivaSaida = (valorIvaSheet > 0 && (tipo === "saida" || tipo === "fechocaixa")) ? valorIvaSheet : (qty * sell) * taxa;
        const ivaEntrada = (valorIvaSheet > 0 && tipo === "entrada") ? valorIvaSheet : (qty * cost) * taxa;
        const ivaDespesa = (valorIvaSheet > 0 && (tipo === "despesas" || tipo === "consumo" || tipo === "despesa")) ? valorIvaSheet : (qty * (cost || sell)) * taxa;

        if ((tipo === "saida" || tipo === "fechocaixa") && contaStock !== "Não") {
          kpi.iva_por_trimestre[yearKey] = (kpi.iva_por_trimestre[yearKey] || 0) + ivaSaida;
        }
        if (tipo === "entrada" && dedutivel) {
          kpi.iva_por_trimestre[yearKey] = (kpi.iva_por_trimestre[yearKey] || 0) - ivaEntrada;
        }
        if ((tipo === "despesas" || tipo === "consumo" || tipo === "despesa") && dedutivel) {
          kpi.iva_por_trimestre[yearKey] = (kpi.iva_por_trimestre[yearKey] || 0) - ivaDespesa;
        }

        if (isDateInRange(d, start, end)) {
          const dayKey = d.toLocaleDateString('pt-PT');

          if ((tipo === "saida" || tipo === "fechocaixa") && contaStock !== "Não") {
            kpi.iva_por_trimestre_no_periodo[yearKey] = (kpi.iva_por_trimestre_no_periodo[yearKey] || 0) + ivaSaida;
          }
          if (tipo === "entrada" && dedutivel) {
            kpi.iva_por_trimestre_no_periodo[yearKey] = (kpi.iva_por_trimestre_no_periodo[yearKey] || 0) - ivaEntrada;
          }
          if ((tipo === "despesas" || tipo === "consumo" || tipo === "despesa") && dedutivel) {
            kpi.iva_por_trimestre_no_periodo[yearKey] = (kpi.iva_por_trimestre_no_periodo[yearKey] || 0) - ivaDespesa;
          }

          if (tipo === "saida" && contaStock !== "Não") {
            const valVenda = qty * sell;
            kpi.fat += valVenda;
            kpi.iva_p += (valorIvaSheet > 0 ? valorIvaSheet : valVenda * taxa);
            kpi.transacoes++;
            if (artigo) {
              if (!kpi.top_products[artigo]) kpi.top_products[artigo] = 0;
              kpi.top_products[artigo] += qty;
            }
            kpi.profit_chart[dayKey] = (kpi.profit_chart[dayKey] || 0) + (valVenda - (qty * cost));
          }
          if (tipo === "fechocaixa" && contaStock !== "Não" && !diasComSaida[dayKey]) {
            const valVenda = qty * sell;
            kpi.fat += valVenda;
            kpi.iva_p += (valorIvaSheet > 0 ? valorIvaSheet : valVenda * taxa);
            kpi.transacoes++;
            if (artigo) {
              if (!kpi.top_products[artigo]) kpi.top_products[artigo] = 0;
              kpi.top_products[artigo] += qty;
            }
            kpi.profit_chart[dayKey] = (kpi.profit_chart[dayKey] || 0) + (valVenda - (qty * cost));
          }
          if (tipo === "entrada") {
            const valCompra = qty * cost;
            kpi.compras += valCompra;
            if (dedutivel) kpi.iva_d += (valorIvaSheet > 0 ? valorIvaSheet : valCompra * taxa);
          }
          if (tipo === "despesas" || tipo === "consumo" || tipo === "despesa") {
            const valDesp = qty * (cost || sell);
            const obs = (r[12] != null) ? String(r[12]).trim() : "";
            let contrib = valDesp;
            if (shouldProrateFixed && isProratableFixedExpense(artigo, obs)) {
              const dailyExpense = valDesp / 30;
              contrib = Math.round(dailyExpense * diasNoPeriodo * 100) / 100;
            }
            kpi.despesas_opex += contrib;
            kpi.profit_chart[dayKey] = (kpi.profit_chart[dayKey] || 0) - contrib;
            if (dedutivel) kpi.iva_d += ivaDespesa;
          }
        }

        if (isDateInRange(d, chartStart, chartEnd)) {
          const dayKey = d.toLocaleDateString('pt-PT');
          if (tipo === "saida" && contaStock !== "Não") {
            const valVenda = qty * sell;
            chart_fat_by_day[dayKey] = (chart_fat_by_day[dayKey] || 0) + valVenda;
            chart_profit_by_day[dayKey] = (chart_profit_by_day[dayKey] || 0) + (valVenda - (qty * cost));
          }
          if (tipo === "fechocaixa" && contaStock !== "Não" && !diasComSaida[dayKey]) {
            const valVenda = qty * sell;
            chart_fat_by_day[dayKey] = (chart_fat_by_day[dayKey] || 0) + valVenda;
            chart_profit_by_day[dayKey] = (chart_profit_by_day[dayKey] || 0) + (valVenda - (qty * cost));
          }
          if (tipo === "despesas" || tipo === "consumo" || tipo === "despesa") {
            const valDesp = qty * (cost || sell);
            chart_desp_by_day[dayKey] = (chart_desp_by_day[dayKey] || 0) + valDesp;
            chart_profit_by_day[dayKey] = (chart_profit_by_day[dayKey] || 0) - valDesp;
          }
        }
      });
    }

    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const trimestreAtual = currentMonth <= 2 ? "T1" : (currentMonth <= 5 ? "T2" : (currentMonth <= 8 ? "T3" : "T4"));
    const valorTrimestreAtual = kpi.iva_por_trimestre[currentYear + "-" + trimestreAtual] || 0;

    const toQuarterKey = function (dt) {
      const m = dt.getMonth();
      const q = m <= 2 ? "T1" : (m <= 5 ? "T2" : (m <= 8 ? "T3" : "T4"));
      return dt.getFullYear() + "-" + q;
    };
    const coveredKeysMap = {};
    if (datesEmpty) {
      coveredKeysMap[currentYear + "-" + trimestreAtual] = true;
    } else {
      const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
      while (cursor <= end) {
        coveredKeysMap[toQuarterKey(cursor)] = true;
        cursor.setMonth(cursor.getMonth() + 1);
      }
      coveredKeysMap[toQuarterKey(end)] = true;
    }
    const coveredKeys = Object.keys(coveredKeysMap).sort();
    let ivaTrimestralFiltrado = 0;
    coveredKeys.forEach(function (k) { ivaTrimestralFiltrado += (kpi.iva_por_trimestre_no_periodo[k] || 0); });
    const trimestresLabel = coveredKeys.length === 1 ? coveredKeys[0] : coveredKeys.join(" + ");

    const stockPorArtigoList = Object.keys(kpi.stock_por_artigo).map(art => ({
      artigo: art,
      quantidadeEmStock: kpi.stock_por_artigo[art].qty,
      valorEmStock: Math.max(0, kpi.stock_por_artigo[art].valor)
    })).filter(x => x.quantidadeEmStock !== 0 || x.valorEmStock > 0).sort((a, b) => b.valorEmStock - a.valorEmStock);

    const todayMs = new Date().setHours(0, 0, 0, 0);
    const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
    let deadStockValue = 0;
    const deadStockList = [];
    Object.keys(kpi.stock_por_artigo).forEach(art => {
      const st = kpi.stock_por_artigo[art];
      if (st.qty > 0 && st.valor > 0 && st.lastEntryDate) {
        const entryMs = (st.lastEntryDate instanceof Date ? st.lastEntryDate : new Date(st.lastEntryDate)).setHours(0, 0, 0, 0);
        if ((todayMs - entryMs) > thirtyDaysMs) {
          deadStockValue += st.valor;
          deadStockList.push({ artigo: art });
        }
      }
    });

    const sheetRH = getStaffSheet(ss);
    let activeStaff = 0;
    let rhEstimado = 0;
    if (sheetRH && sheetRH.getLastRow() > 1) {
      const rhLastCol = Math.max(sheetRH.getLastColumn(), 20);
      const dataRH = sheetRH.getRange(2, 1, sheetRH.getLastRow() - 1, rhLastCol).getValues();
      const mesAtual = now.getMonth() + 1;
      const anoAtual = now.getFullYear();
      dataRH.forEach(r => {
        const status = String(r[8] || "");
        const custoBase = parsePTFloat(r[9]);
        if (custoBase <= 0) return;
        const fator = getStaffCostFactor(status, r[14], r[19] || null, mesAtual, anoAtual);
        if (fator > 0) {
          kpi.rh += Math.round(custoBase * fator * 100) / 100;
          rhEstimado += custoBase;
          if (status === "Ativo") activeStaff++;
        }
      });
    }

    const diasNoMes = new Date(end.getFullYear(), end.getMonth() + 1, 0).getDate();
    const diaAtual = end.getDate();
    const rh_period = (endIsCurrentMonth && diasNoPeriodo <= 31) ? Math.round((kpi.rh / diasNoMes) * diaAtual * 100) / 100 : Math.round((kpi.rh / 30) * diasNoPeriodo * 100) / 100;
    const dailyHR = diasNoPeriodo > 0 ? rh_period / diasNoPeriodo : 0;

    const margemVal = kpi.fat - kpi.despesas_opex - rh_period;

    const totalFixed = (kpi.rh) + (parseFloat(PropertiesService.getDocumentProperties().getProperty("MANUAL_FIXED_EXPENSES")) || 0);
    const contribMargin = kpi.fat > 0 ? ((margemVal + totalFixed) / kpi.fat) : 0;
    const breakEvenRevenue = (contribMargin > 0) ? (totalFixed / contribMargin) : 0;
    const projectedRev = (diasNoPeriodo > 0) ? (kpi.fat * 30 / diasNoPeriodo) : 0;
    const gap = breakEvenRevenue - projectedRev;

    const irc_est = margemVal <= 0 ? 0 : round2(Math.min(margemVal, 50000) * 0.17 + Math.max(0, margemVal - 50000) * 0.21);
    const top5Labels = Object.keys(kpi.top_products).sort((a, b) => kpi.top_products[b] - kpi.top_products[a]).slice(0, 5);

    Object.keys(chart_profit_by_day).forEach(function (dayKey) {
      const pa = dayKey.split('/');
      if (pa.length >= 3) {
        const dayDate = new Date(parseInt(pa[2], 10), parseInt(pa[1], 10) - 1, parseInt(pa[0], 10));
        if (isDateInRange(dayDate, start, end)) {
          chart_profit_by_day[dayKey] = (chart_profit_by_day[dayKey] || 0) - dailyHR;
        }
      }
    });

    const allChartKeys = Array.from(new Set([
      ...Object.keys(chart_fat_by_day),
      ...Object.keys(chart_desp_by_day),
      ...Object.keys(chart_profit_by_day)
    ])).sort(function (a, b) {
      const pa = a.split('/'), pb = b.split('/');
      return new Date(pa[2], pa[1] - 1, pa[0]) - new Date(pb[2], pb[1] - 1, pb[0]);
    });

    const toYMD = function (d) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return y + "-" + m + "-" + day;
    };
    const taxaAbsorcao = calculateDynamicOpEx(impersonateTarget, toYMD(start), toYMD(end));

    let entradasCaixa = 0;
    let saidasCaixa = 0;
    let faturasAPagamento = 0;
    let contasAReceber = 0;
    let fatRecebido = 0, fatPendente = 0, despesasPago = 0, despesasPorPagar = 0;
    const devedoresMap = {};
    const dataForTesouraria = (preloaded && preloaded.batch && preloaded.batch.dataNew && preloaded.batch.dataNew.length > 0)
      ? preloaded.batch.dataNew
      : (sheetLog && sheetLog.getLastRow() > 1 ? sheetLog.getRange(2, 1, sheetLog.getLastRow(), 26).getValues() : []);

    if (dataForTesouraria.length > 0) {
      dataForTesouraria.forEach(function (r) {
        let d = r[0];
        if (typeof d === "string") { const p = d.split("/"); if (p.length === 3) d = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
        const inPeriod = (d instanceof Date && !isNaN(d)) && isDateInRange(d, start, end);
        const tipo = normalizeTipo(r[1]) || "entrada";
        const qty = parseFloat(r[5]) || 0;
        const cost = parseFloat(String(r[6] || "0").replace(",", ".")) || 0;
        const sell = parseFloat(String(r[7] || "0").replace(",", ".")) || 0;
        const status = (r[16] != null && String(r[16]).trim() !== "") ? String(r[16]).trim().toLowerCase() : "aberto";
        const isEntradaOuDespesa = (tipo === "entrada" || tipo === "despesas" || tipo === "despesa" || tipo === "consumo");
        const isDespesa = (tipo === "despesas" || tipo === "despesa" || tipo === "consumo");
        const totalLinha = (tipo === "saida" || tipo === "fechocaixa") ? (qty * sell) : (isEntradaOuDespesa ? (qty * (cost || sell)) : 0);

        let contribDespesa = totalLinha;
        if (inPeriod && isDespesa) {
          const artigo = (r[4] != null && r[4] !== "") ? String(r[4]).trim() : null;
          const obs = (r[12] != null) ? String(r[12]).trim() : "";
          if (shouldProrateFixed && isProratableFixedExpense(artigo, obs)) {
            const dailyExpense = totalLinha / 30;
            contribDespesa = Math.round(dailyExpense * diasNoPeriodo * 100) / 100;
          }
        }

        if (status === "pago") {
          if (tipo === "saida" || tipo === "fechocaixa") { entradasCaixa += totalLinha; if (inPeriod) fatRecebido += totalLinha; }
          else if (isEntradaOuDespesa) { saidasCaixa += totalLinha; if (inPeriod && isDespesa) despesasPago += contribDespesa; }
        } else if (status === "aberto" || status === "pendente") {
          if (isEntradaOuDespesa) { faturasAPagamento += totalLinha; if (inPeriod && isDespesa) despesasPorPagar += contribDespesa; }
          else if (tipo === "saida") {
            const valorPago = parseFloat(String(r[18] || "0").replace(",", ".")) || 0;
            const totalPendente = Math.max(0, totalLinha - valorPago);
            contasAReceber += totalPendente;
            if (inPeriod) fatPendente += totalPendente;
            const entity = (r[3] || r[12] || r[4] || "Sem Cliente").toString().trim() || "Sem Cliente";
            devedoresMap[entity] = (devedoresMap[entity] || 0) + totalPendente;
          }
        }
      });
    }

    const listaDevedores = Object.entries(devedoresMap).map(function (e) { return { nome: e[0], valor: Math.round(e[1] * 100) / 100 }; }).sort(function (a, b) { return b.valor - a.valor; }).slice(0, 5);
    const saldoBancario = entradasCaixa - saidasCaixa - rh_period;
    const compromissosRH = rhEstimado;
    const caixaLivre = entradasCaixa - saidasCaixa;
    const rhPago = 0;
    const despesasAbertoMes = 0;

    const resObj = {
      success: true,
      fat: kpi.fat,
      taxaAbsorcao: { rate: taxaAbsorcao.rate, label: taxaAbsorcao.label || "365d" },
      financeiro: {
        faturacao: (Math.round((fatRecebido + fatPendente) * 100) / 100).toFixed(2),
        lucro_liq: margemVal.toFixed(2),
        margem_perc: kpi.fat > 0 ? ((margemVal / kpi.fat) * 100).toFixed(1) : "0.0",
        iva_pagar: (kpi.iva_p - kpi.iva_d).toFixed(2),
        irc_est: irc_est.toFixed(2),
        despesas: (Math.round((despesasPago + despesasPorPagar) * 100) / 100 + rh_period).toFixed(2),
        ticket_medio: kpi.transacoes > 0 ? (kpi.fat / kpi.transacoes).toFixed(2) : "0.00",
        transacoes: kpi.transacoes,
        saldoBancario: saldoBancario.toFixed(2),
        faturasAPagamento: faturasAPagamento.toFixed(2),
        contasAReceber: contasAReceber.toFixed(2),
        listaDevedores: listaDevedores,
        despesasAbertoMes: despesasAbertoMes.toFixed(2),
        compromissosRH: compromissosRH.toFixed(2),
        caixaLivre: caixaLivre.toFixed(2),
        fatRecebido: Math.round(fatRecebido * 100) / 100,
        fatPendente: Math.round(fatPendente * 100) / 100,
        despesasPago: Math.round(despesasPago * 100) / 100,
        despesasPorPagar: Math.round(despesasPorPagar * 100) / 100,
        rhProcessado: rh_period.toFixed(2),
        rhPago: rhPago.toFixed(2),
        rhAVencer: Math.max(0, rhEstimado - rh_period).toFixed(2)
      },
      chartFinanceiro: { chartLabels: allChartKeys, chartData: allChartKeys.map(k => chart_profit_by_day[k] || 0), chartFatData: allChartKeys.map(k => chart_fat_by_day[k] || 0), chartDespData: allChartKeys.map(k => chart_desp_by_day[k] || 0) },
      stock: {
        valor: kpi.stock_val.toFixed(2),
        rotatividade: kpi.stock_val > 0 ? (kpi.fat / kpi.stock_val).toFixed(2) : "0.0",
        compras: kpi.compras.toFixed(2),
        diasCobertura: (kpi.fat > 0 && diasNoPeriodo > 0) ? (kpi.stock_val * diasNoPeriodo / kpi.fat).toFixed(1) : null,
        topLabels: top5Labels,
        topData: top5Labels.map(k => kpi.top_products[k]),
        stockPorArtigo: stockPorArtigoList,
        vasilhame: { saldos: getSaldosVasilhameComValor(impersonateTarget, preloaded ? preloaded.batch : null).saldos || [], valorTotalNaRua: (getSaldosVasilhameComValor(impersonateTarget, preloaded ? preloaded.batch : null).valorTotalNaRua || 0).toFixed(2) }
      },
      alertaIvaTrimestral: { T1: (kpi.iva_por_trimestre[currentYear + "-T1"] || 0).toFixed(2), T2: (kpi.iva_por_trimestre[currentYear + "-T2"] || 0).toFixed(2), T3: (kpi.iva_por_trimestre[currentYear + "-T3"] || 0).toFixed(2), T4: (kpi.iva_por_trimestre[currentYear + "-T4"] || 0).toFixed(2), trimestreAtual: trimestreAtual, valorTrimestreAtual: valorTrimestreAtual.toFixed(2), ivaTrimestralFiltrado: ivaTrimestralFiltrado.toFixed(2), trimestresLabel: trimestresLabel },
      rh: { custo_total: rh_period.toFixed(2), custo_mensal: rhEstimado.toFixed(2), custo_estimado: rhEstimado.toFixed(2), ativos: activeStaff, custo_medio: activeStaff > 0 ? (rh_period / activeStaff).toFixed(0) : 0, ideal: Math.ceil(kpi.fat / 4000) || 1 },
      breakEven: { gap: gap, reached: gap <= 0, target: breakEvenRevenue, projectedRev: projectedRev },
      deadStockValue: (deadStockValue || 0).toFixed(2),
      deadStockList: deadStockList || [],
      erosionCount: erosionCount || 0,
      erosionItems: erosionItems || [],
      frota: getDashboardFrotaData(impersonateTarget, preloaded ? preloaded.batch : null)
    };

    if (ctx.businessVertical === "Condominium") {
      if (typeof _calcCondominiumFinancials === "function") {
        resObj.condominiumData = _calcCondominiumFinancials(ss, dataLog, start, end);
      }
    }

    resObj.aiAuto = getAIAutoPreference(email);
    if (runAI) {
      try { resObj.aiInsight = getFlowlyAIInsight(resObj, impersonateTarget); } catch (aiErr) { resObj.aiInsight = { success: false, error: "Análise IA indisponível. Tente novamente.", currentCredits: getAiCredits(impersonateTarget) }; }
    } else {
      var lastInsight = (preloaded && preloaded.lastAIInsight) ? preloaded.lastAIInsight : getLastAIHistoryEntry(impersonateTarget || email);
      if (lastInsight && typeof lastInsight === "object") { lastInsight.currentCredits = getAiCredits(impersonateTarget); resObj.aiInsight = lastInsight; } else { resObj.aiInsight = { atRest: true, currentCredits: getAiCredits(impersonateTarget) }; }
    }

    resObj.aiCredits = (resObj.aiInsight && resObj.aiInsight.currentCredits != null) ? resObj.aiInsight.currentCredits : getAiCredits(impersonateTarget);
    resObj.clientConfig = ctx.planConfig || Object.assign({}, DEFAULT_PLAN_CONFIG);

    var planCfg = resObj.clientConfig;
    var hasFin = planCfg.dashFinanceiro === true || (planCfg.financeiro === true);
    var hasStock = planCfg.dashStocks === true || (planCfg.stocks === true);
    var hasRH = planCfg.dashRH === true || (planCfg.rh === true);

    if (!hasFin) {
      resObj.financeiro = { faturacao: "0", lucro_liq: "0", margem_perc: "0", iva_pagar: "0", irc_est: "0", despesas: "0", ticket_medio: "0", transacoes: 0, saldoBancario: "0", faturasAPagamento: "0", contasAReceber: "0", listaDevedores: [], despesasAbertoMes: "0", compromissosRH: "0", caixaLivre: "0", fatRecebido: 0, fatPendente: 0, despesasPago: 0, despesasPorPagar: 0, rhProcessado: "0", rhPago: "0", rhAVencer: "0" };
      resObj.chartFinanceiro = { chartLabels: [], chartData: [], chartFatData: [], chartDespData: [] };
    }
    if (!hasStock) {
      resObj.stock = { valor: "0", rotatividade: "0", compras: "0", diasCobertura: null, topLabels: [], topData: [], stockPorArtigo: [] };
      resObj.deadStockValue = "0"; resObj.deadStockList = []; resObj.erosionCount = 0; resObj.erosionItems = [];
    }
    if (!hasRH) resObj.rh = { custo_total: "0", custo_mensal: "0", custo_estimado: "0", ativos: 0, custo_medio: 0, ideal: 0 };

    return resObj;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function isProratableFixedExpense(artigo, obs) {
  const a = (artigo || "").toLowerCase();
  const o = (obs || "").toLowerCase();
  const fixedPatterns = ["renda", "aluguer", "rent", "fixos", "operacional"];
  return fixedPatterns.some(function (p) { return a.indexOf(p) >= 0 || o.indexOf(p) >= 0; });
}

function getMasterData(impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (ctx.isMaster && !ctx.sheetId) return getSuperAdminDashboardData(impersonateTarget);

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheetLog = ss.getSheetByName(SHEET_TAB_NAME);
    let fornecedores = [], artigos = [];

    if (sheetLog && sheetLog.getLastRow() > 1) {
      const d = sheetLog.getRange(2, 1, sheetLog.getLastRow(), 20).getValues();
      artigos = [...new Set(d.map(r => (r[4] || "").toString().trim()))].filter(v => v !== "").sort();
      fornecedores = ["Flowly SaaS", "Fornecedor Geral"];
    }

    const sheetRH = getStaffSheet(ss);
    let staff = [];
    if (sheetRH && sheetRH.getLastRow() > 1) {
      const lastCol = Math.max(sheetRH.getLastColumn(), 20);
      const now = new Date();
      const mesRef = now.getMonth() + 1;
      const anoRef = now.getFullYear();
      const diasUteis = getWorkingDaysInMonth(mesRef, anoRef);

      staff = sheetRH.getRange(2, 1, sheetRH.getLastRow() - 1, lastCol).getValues().map(r => {
        const estado = String(r[8] || '');
        const custoCalc = parsePTFloat(r[9]);
        const admissaoRaw = r[14];
        const admissaoStr = (admissaoRaw instanceof Date) ? admissaoRaw.toLocaleDateString('pt-PT') : String(admissaoRaw || '');
        const dataSaidaRaw = r[19] || null;
        const dataSaidaStr = (dataSaidaRaw instanceof Date) ? dataSaidaRaw.toLocaleDateString('pt-PT') : (dataSaidaRaw ? String(dataSaidaRaw) : '');
        const fator = getStaffCostFactor(estado, admissaoRaw, dataSaidaRaw, mesRef, anoRef);
        const custoEfetivo = Math.round(custoCalc * fator * 100) / 100;
        const vencimento = parseFloat(r[4]) || 0;
        const subAlim = parseFloat(r[5]) || 0;
        const provRescisao = parsePTFloat(r[12]);
        const provFormacao = parsePTFloat(r[13]);
        const custoPonderado = Math.round((vencimento + provFormacao + (subAlim * diasUteis) + provRescisao) * fator * 100) / 100;

        return {
          id: String(r[0] || ''), nome: String(r[1] || ''), nif: String(r[2] || ''), cargo: String(r[3] || ''),
          vencimento: vencimento, subAlim: subAlim, seguro: parseFloat(r[6]) || 0, tsuPct: parseFloat(r[7]) || 23.75,
          status: estado, custoMensalReal: custoEfetivo, custoBase: custoCalc, fatorProporcional: Math.round(fator * 100),
          provFerias: parseFloat(r[10]) || 0, provNatal: parseFloat(r[11]) || 0, provRescisao: provRescisao, provFormacao: provFormacao,
          admissao: admissaoStr, diasContrato: parseInt(r[15]) || 0, premios: parsePTFloat(r[16]) || 0,
          email: String(r[17] || ''), dataSaida: dataSaidaStr, custoTotal: custoEfetivo, custoPonderado: custoPonderado
        };
      }).filter(s => s.id !== "");
    }

    const sheetUsers = ss.getSheetByName(SHEET_USERS_NAME);
    let users = [];
    if (sheetUsers && sheetUsers.getLastRow() > 1) {
      users = sheetUsers.getRange(2, 1, sheetUsers.getLastRow() - 1, 6).getValues().map(r => ({
        email: String(r[0] || ''), name: String(r[2] || ''), role: String(r[3] || ''), status: String(r[4] || ''),
        permissions: (r[5] && typeof r[5] === 'string' && r[5].startsWith('{')) ? (() => { try { return JSON.parse(r[5]); } catch (e) { return { logistica: true }; } })() : { logistica: true }
      }));
    }

    let ccData = [];
    const diasComSaida = [];
    if (sheetLog && sheetLog.getLastRow() > 1) {
      const raw = sheetLog.getRange(2, 1, sheetLog.getLastRow(), Math.max(26, sheetLog.getLastColumn())).getValues();
      raw.forEach(r => {
        const tipo = normalizeTipo(r[1]) || "entrada";
        const contaStock = (r[19] != null && String(r[19]).trim() !== "") ? String(r[19]).trim() : "Sim";
        if (tipo === "saida" && contaStock !== "Não") {
          let d = r[0];
          if (typeof d === "string") { const p = d.split("/"); if (p.length === 3) d = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
          if (d instanceof Date && !isNaN(d)) {
            const key = d.toLocaleDateString("pt-PT");
            if (diasComSaida.indexOf(key) === -1) diasComSaida.push(key);
          }
        }
      });
      ccData = raw.map((r, i) => {
        const qty = parseFloat(r[5]) || 0;
        const custo = parseFloat(r[6]) || 0;
        const venda = parseFloat(r[7]) || 0;
        const tipo = normalizeTipo(r[1]) || "entrada";
        const isEntradaOuDespesa = (tipo === "entrada" || tipo === "despesa" || tipo === "despesas" || tipo === "consumo");
        const total = isEntradaOuDespesa ? (qty * (custo || venda)) : (qty * venda);
        const valorPago = parseFloat(r[18]) || 0;
        const totalPendente = Math.max(0, total - valorPago);
        const contaStock = (r[19] != null && String(r[19]).trim() !== "") ? String(r[19]).trim() : "Sim";

        let d = r[0];
        if (typeof d === "string") { const p = d.split("/"); if (p.length === 3) d = new Date(p[2] + "-" + p[1] + "-" + p[0]); }
        const dataNorm = (d instanceof Date && !isNaN(d)) ? d.toLocaleDateString("pt-PT") : (r[0] != null ? String(r[0]) : "");
        const docIdVal = (r[20] != null && String(r[20]).trim() !== "") ? String(r[20]).trim() : "sem-id";
        const idEntidadeVal = (r[24] != null && String(r[24]).trim() !== "") ? String(r[24]).trim() : "";
        const taxaRaw = parseFloat(String(r[8] || "").replace("%", "").trim());
        const taxaIvaPct = (!isNaN(taxaRaw) && taxaRaw > 0) ? (taxaRaw > 1 ? Math.round(taxaRaw) : Math.round(taxaRaw * 100)) : null;
        const linkFotoVal = (r[13] != null && String(r[13]).trim() !== "") ? String(r[13]).trim() : "";

        return {
          rowIndex: i + 2, data: r[0], dataNorm: dataNorm, tipo: tipo, fornecedor: canonicalizeEntityForDisplay(r[3] || "") || (r[3] || "N/A"),
          artigo: (r[4] || "").toString().trim(), qtd: parseFloat(r[5]) || 0, taxaIva: taxaIvaPct != null ? taxaIvaPct : "", total: isNaN(total) ? "0.00" : total.toFixed(2),
          totalPendente: isNaN(totalPendente) ? "0.00" : totalPendente.toFixed(2), estado: (r[16] != null && r[16] !== "") ? String(r[16]).trim() : "Aberto",
          contaStock: contaStock, docId: docIdVal, idEntidade: idEntidadeVal, linkFoto: linkFotoVal
        };
      }).filter(x => parseFloat(x.total) > 0);
    }

    const safeCcData = ccData.map(item => Object.assign({}, item, { data: (item.data instanceof Date) ? item.data.toLocaleDateString('pt-PT') : String(item.data || '') }));
    const safeStaff = staff.map(s => Object.assign({}, s, { admissao: (s.admissao instanceof Date) ? s.admissao.toLocaleDateString('pt-PT') : String(s.admissao || '') }));

    return {
      success: true, isMaster: false, isImpersonating: ctx.isImpersonating, clientName: ctx.clientName, planConfig: ctx.planConfig,
      clientConfig: ctx.planConfig, fornecedores, artigos, staff: safeStaff, users, cc: safeCcData, diasComSaida: diasComSaida
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getContaCorrenteData(impersonateTarget) {
  const data = getMasterData(impersonateTarget);
  return data.cc || [];
}

function _calcularSaldoDevedorCliente(ss, idEntidade) {
  if (!idEntidade || !ss) return { saldo: 0, detalhe: "" };
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { saldo: 0, detalhe: "" };
  var lastCol = Math.max(26, sheet.getLastColumn());
  var raw = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var idNorm = String(idEntidade).trim();
  var saldo = 0;
  var linhas = [];
  raw.forEach(function (r) {
    var tipo = normalizeTipo(r[1]) || "entrada";
    if (tipo !== "saida" && tipo !== "fechocaixa") return;
    var status = (r[16] != null && r[16] !== "") ? String(r[16]).trim() : "Aberto";
    if (status !== "Aberto") return;
    var idEnt = (r[24] != null && String(r[24]).trim() !== "") ? String(r[24]).trim() : "";
    if (idEnt !== idNorm) return;

    var qty = parseFloat(r[5]) || 0;
    var venda = parseFloat(r[7]) || 0;
    var valorPago = parseFloat(r[18]) || 0;
    var total = qty * venda;
    var pendente = Math.max(0, total - valorPago);
    saldo += pendente;

    var dataStr = r[0] ? (r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(r[0])) : "";
    linhas.push({ data: dataStr, total: total, pendente: pendente });
  });
  var detalhe = linhas.map(function (l) { return l.data + ": " + l.pendente.toFixed(2) + " €"; }).join("; ");
  return { saldo: saldo, detalhe: detalhe };
}

function _buildEmailCobrancaHtml(nomeEmpresa, valorStr, stripeLink) {
  var payUrl = (stripeLink && /^https?:\/\//i.test(String(stripeLink).trim())) ? String(stripeLink).trim() : "";
  var btnHtml = "";
  if (payUrl) {
    btnHtml = '<p style="text-align:center;margin:24px 0;"><a href="' + payUrl.replace(/"/g, "&quot;") + '" style="display:inline-block;background:#06B6D4;color:#FFFFFF;padding:16px 32px;border-radius:8px;font-weight:bold;font-size:16px;text-decoration:none;">Pagar Agora via Stripe</a></p>';
  } else {
    btnHtml = '<p style="text-align:center;margin:24px 0;"><a href="https://flowly.pt" style="display:inline-block;background:#06B6D4;color:#FFFFFF;padding:16px 32px;border-radius:8px;font-weight:bold;font-size:16px;text-decoration:none;">Pagar Agora</a></p>';
  }
  
  var innerBody = '<p style="font-size:14px;color:#334155;line-height:1.6;">Exmo(a) Sr(a) <strong>' + (nomeEmpresa || "Cliente").replace(/</g, "&lt;").replace(/>/g, "&gt;") + '</strong>,</p>' +
                  '<p style="font-size:14px;color:#334155;line-height:1.6;">Informamos que na nossa contabilidade consta um saldo devedor no valor de:</p>' +
                  '<p style="color:#10B981;font-weight:bold;font-size:24px;margin:16px 0;">' + (valorStr || "0,00").replace(/</g, "&lt;") + ' €</p>' +
                  btnHtml +
                  '<p style="font-size:14px;color:#334155;line-height:1.6;">Solicitamos a regularização deste valor no mais breve prazo possível.</p>';
                  
  return _buildStandardEmailHTML("Lembrete de Cobrança", innerBody);
}

function _registarAuditAvisoCobranca(ss, userEmail, idCliente, valor, sucesso) {
  try {
    var sheet = ss.getSheetByName(AUDIT_DB_TAB);
    if (!sheet) return;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var motivo = "Aviso cobrança " + (sucesso ? "enviado" : "falhou") + " - Valor: " + (valor || 0).toFixed(2) + " €";
    sheet.appendRow([now, userEmail || "", idCliente || "", "Aviso_Cobranca_Email", motivo]);
  } catch (e) { }
}

function processarEmailsCobranca(listaIds, impersonateTarget) {
  try {
    if (!Array.isArray(listaIds) || listaIds.length === 0) return { success: false, enviados: 0, falhas: 0, erros: [], error: "Lista de IDs vazia." };

    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, enviados: 0, falhas: 0, erros: [], error: "Infraestrutura não encontrada." };

    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var userEmail = Session.getActiveUser().getEmail();
    var idsUnicos = [];
    var seen = {};
    listaIds.forEach(function (id) {
      var s = (id || "").toString().trim();
      if (s && !seen[s]) { seen[s] = true; idsUnicos.push(s); }
    });

    var enviados = 0;
    var falhas = 0;
    var erros = [];

    idsUnicos.forEach(function (idEntidade) {
      var entidade = _getEntidadeByEmail(idEntidade, ss);
      if (!entidade || !entidade.email) {
        falhas++;
        var prefixo = (String(idEntidade).trim().toUpperCase().indexOf("FORN") === 0) ? "Fornecedores" : "Clientes_DB";
        erros.push("Entidade " + idEntidade + ": sem email registado na aba '" + prefixo + "'.");
        _registarAuditAvisoCobranca(ss, userEmail, idEntidade, 0, false);
        return;
      }

      var calc = _calcularSaldoDevedorCliente(ss, idEntidade);
      if (calc.saldo <= 0) {
        falhas++;
        erros.push(entidade.nomeEmpresa + ": saldo zero ou negativo.");
        return;
      }

      var valorStr = calc.saldo.toFixed(2).replace(".", ",");
      var assunto = "Lembrete de cobrança — " + valorStr + " € em dívida";
      var htmlBody = _buildEmailCobrancaHtml(entidade.nomeEmpresa, valorStr, entidade.stripeLink || "");

      try {
        let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
        var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
        GmailApp.sendEmail(entidade.email, assunto, "", options);
        enviados++;
        _registarAuditAvisoCobranca(ss, userEmail, idEntidade, calc.saldo, true);
      } catch (mailErr) {
        falhas++;
        erros.push(entidade.nomeEmpresa + ": " + (mailErr.message || mailErr.toString()));
        _registarAuditAvisoCobranca(ss, userEmail, idEntidade, calc.saldo, false);
      }
    });

    return { success: true, enviados: enviados, falhas: falhas, erros: erros };
  } catch (e) {
    return { success: false, enviados: 0, falhas: 0, erros: [e.toString()], error: e.toString() };
  }
}