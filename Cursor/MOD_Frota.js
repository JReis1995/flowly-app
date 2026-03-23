// Ficheiro: MOD_Frota.js
/// ==========================================
// 🚐 MÓDULO DE GESTÃO DE FROTA
// ==========================================

function initFrotaDB(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(FROTA_VEICULOS_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(FROTA_VEICULOS_TAB);
      sheet.appendRow(["ID_Veiculo", "Matricula", "Marca_Modelo", "Ano", "Combustivel", "Categoria", "Lotacao", "Data_Inspecao", "Validade_Seguro", "Proxima_Revisao", "Status", "Ultima_Lavagem"]);
      sheet.getRange(1, 1, 1, 12).setFontWeight("bold");
    } else {
      var headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 13)).getValues()[0];
      var idxCombustivel = headers.indexOf("Combustivel");
      var idxCategoria = headers.indexOf("Categoria");
      if (idxCombustivel >= 0 && idxCategoria < 0) {
        sheet.insertColumnAfter(idxCombustivel + 1);
        sheet.getRange(1, idxCombustivel + 2).setValue("Categoria");
        sheet.getRange(1, idxCombustivel + 2).setFontWeight("bold");
      }
      var idxUltimaLavagem = headers.indexOf("Ultima_Lavagem");
      if (idxUltimaLavagem < 0) {
        var lastCol = sheet.getLastColumn();
        sheet.insertColumnAfter(lastCol);
        sheet.getRange(1, lastCol + 1).setValue("Ultima_Lavagem");
        sheet.getRange(1, lastCol + 1).setFontWeight("bold");
      }
    }
    return { success: true, sheet: sheet };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getVeiculos(impersonateTarget) {
  try {
    var init = initFrotaDB(impersonateTarget);
    if (!init.success || !init.sheet) return [];
    var sheet = init.sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var numCols = Math.max(12, sheet.getLastColumn());
    var rawData = sheet.getRange(2, 1, lastRow, numCols).getValues();
    var validData = rawData.filter(function (row) { return row[0] && String(row[0]).trim() !== ""; });
    var out = [];
    for (var i = 0; i < validData.length; i++) {
      var r = validData[i];
      out.push({
        id: r[0] || "", matricula: r[1] || "", marcaModelo: r[2] || "", ano: r[3] || "", combustivel: r[4] || "",
        categoria: r[5] || "", lotacao: r[6] || "",
        dataInspecao: r[7] ? (r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[7])) : "",
        validadeSeguro: r[8] ? (r[8] instanceof Date ? Utilities.formatDate(r[8], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[8])) : "",
        proximaRevisao: r[9] ? (r[9] instanceof Date ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[9])) : "",
        status: r[10] || "Ativo",
        ultimaLavagem: r[11] ? (r[11] instanceof Date ? Utilities.formatDate(r[11], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[11]).trim()) : ""
      });
    }
    return out;
  } catch (e) { return []; }
}

function saveVeiculo(veiculoData, impersonateTarget) {
  try {
    var init = initFrotaDB(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Frota." };
    var sheet = init.sheet;
    var id = (veiculoData && veiculoData.id) ? String(veiculoData.id).trim() : "";
    var isUpdate = !!id;
    if (!isUpdate) id = Utilities.getUuid();
    var ultimaLavagemVal = (veiculoData.ultimaLavagem || "").trim();
    var row = [
      id, (veiculoData.matricula || "").trim(), (veiculoData.marcaModelo || "").trim(), veiculoData.ano || "",
      (veiculoData.combustivel || "").trim(), (veiculoData.categoria || "").trim(), (veiculoData.lotacao || "").trim(),
      (veiculoData.dataInspecao || "").trim(), (veiculoData.validadeSeguro || "").trim(), (veiculoData.proximaRevisao || "").trim(),
      (veiculoData.status || "Ativo").trim(), ultimaLavagemVal
    ];
    if (isUpdate) {
      var lastRow = sheet.getLastRow();
      var numCols = Math.max(sheet.getLastColumn(), row.length);
      var data = sheet.getRange(2, 1, lastRow, numCols).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0] || "").trim() === id) {
          if (!ultimaLavagemVal && data[i][11]) row[11] = data[i][11] instanceof Date ? Utilities.formatDate(data[i][11], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][11] || "").trim();
          var rowIndex = i + 2;
          while (row.length < numCols) row.push("");
          sheet.getRange(rowIndex, 1, 1, numCols).setValues([row]);
          return { success: true, id: id };
        }
      }
      return { success: false, error: "Veículo não encontrado para atualização." };
    } else {
      var newRow = sheet.getLastRow() + 1;
      var numCols = Math.max(sheet.getLastColumn(), row.length);
      while (row.length < numCols) row.push("");
      sheet.getRange(newRow, 1, 1, numCols).setValues([row]);
      return { success: true, id: id };
    }
  } catch (e) { return { success: false, error: e.toString() }; }
}

function registarLavagem(matricula, impersonateTarget) {
  try {
    var init = initFrotaDB(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Frota." };
    var sheet = init.sheet;
    var matNorm = (matricula || "").toString().trim().toUpperCase();
    if (!matNorm) return { success: false, error: "Matrícula inválida." };
    var hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "Nenhum veículo encontrado." };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idxMatricula = headers.indexOf("Matricula");
    var idxUltimaLavagem = headers.indexOf("Ultima_Lavagem");
    if (idxMatricula < 0 || idxUltimaLavagem < 0) return { success: false, error: "Estrutura da tabela inválida." };
    var data = sheet.getRange(2, idxMatricula + 1, lastRow, idxMatricula + 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toUpperCase() === matNorm) {
        sheet.getRange(i + 2, idxUltimaLavagem + 1).setValue(hoje);
        return { success: true };
      }
    }
    return { success: false, error: "Viatura com matrícula " + matricula + " não encontrada." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteVeiculo(idVeiculo, impersonateTarget) {
  try {
    var init = initFrotaDB(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Frota." };
    var sheet = init.sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "Nenhum veículo encontrado." };
    var data = sheet.getRange(2, 1, lastRow, 1).getValues();
    var idNorm = (idVeiculo || "").toString().trim();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === idNorm) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: "Veículo não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function initCustosDB(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(FROTA_CUSTOS_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(FROTA_CUSTOS_TAB);
      sheet.appendRow(["ID_Custo", "Data", "Matricula", "Motorista", "Tipo", "Km_Atuais", "Litros", "Custo_Total", "Fatura_Link", "Observacoes", "Consumo_Medio"]);
      sheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    }
    return { success: true, sheet: sheet };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _getUltimoKmAbastecimento(sheet, matricula) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow, 11).getValues();
  var matNorm = (matricula || "").toString().trim().toUpperCase();
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var tipo = (row[4] || "").toString().trim();
    var mat = (row[2] || "").toString().trim().toUpperCase();
    if (tipo === "Abastecimento" && mat === matNorm) {
      var km = parseFloat(row[5]);
      return (!isNaN(km) && km > 0) ? km : null;
    }
  }
  return null;
}

function saveCusto(custoData, impersonateTarget) {
  try {
    var init = initCustosDB(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Frota_Custos." };
    var sheet = init.sheet;
    var id = Utilities.getUuid();
    var data = (custoData.data || "").trim();
    if (!data) data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var matricula = (custoData.matricula || "").trim();
    var motorista = (custoData.motorista || "").trim();
    var tipo = (custoData.tipo || "").trim();
    var kmAtuais = parseFloat(custoData.kmAtuais);
    var litros = parseFloat(custoData.litros);
    var custoTotal = parseFloat(custoData.custoTotal);
    var faturaLink = (custoData.faturaLink || "").trim();
    var observacoes = (custoData.observacoes || "").trim();
    var consumoMedio = 0;
    if (tipo === "Abastecimento" && !isNaN(litros) && litros > 0 && !isNaN(kmAtuais) && kmAtuais > 0) {
      var kmAnteriores = _getUltimoKmAbastecimento(sheet, matricula);
      if (kmAnteriores !== null && kmAtuais > kmAnteriores) consumoMedio = (litros / (kmAtuais - kmAnteriores)) * 100;
    }
    var consumoStr = (consumoMedio > 0) ? String(Math.round(consumoMedio * 100) / 100) : "N/A";
    var row = [id, data, matricula, motorista, tipo, isNaN(kmAtuais) ? "" : kmAtuais, isNaN(litros) ? "" : litros, isNaN(custoTotal) ? "" : custoTotal, faturaLink, observacoes, consumoStr];
    sheet.appendRow(row);
    return { success: true, id: id };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getCustos(impersonateTarget, limit) {
  try {
    var init = initCustosDB(impersonateTarget);
    if (!init.success || !init.sheet) return [];
    var sheet = init.sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var n = (limit && limit > 0) ? Math.min(limit, lastRow - 1) : Math.min(20, lastRow - 1);
    var startRow = Math.max(2, lastRow - n + 1);
    var rawData = sheet.getRange(startRow, 1, lastRow, 11).getValues();
    var out = [];
    for (var i = rawData.length - 1; i >= 0; i--) {
      var r = rawData[i];
      out.push({
        id: r[0] || "",
        data: r[1] ? (r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[1])) : "",
        matricula: r[2] || "", motorista: r[3] || "", tipo: r[4] || "",
        kmAtuais: r[5] !== "" && !isNaN(parseFloat(r[5])) ? parseFloat(r[5]) : null,
        litros: r[6] !== "" && !isNaN(parseFloat(r[6])) ? parseFloat(r[6]) : null,
        custoTotal: r[7] !== "" && !isNaN(parseFloat(r[7])) ? parseFloat(r[7]) : null,
        consumoMedio: r[10] || ""
      });
    }
    return out;
  } catch (e) { return []; }
}

function _getVeiculoByMatricula(ss, matricula) {
  var sheet = ss.getSheetByName(FROTA_VEICULOS_TAB);
  if (!sheet || !matricula) return { combustivel: "", categoria: "" };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { combustivel: "", categoria: "" };
  var matNorm = String(matricula).trim().toUpperCase();
  var data = sheet.getRange(2, 1, lastRow, 11).getValues();
  for (var i = 0; i < data.length; i++) {
    var m = String(data[i][1] || "").trim().toUpperCase();
    if (m === matNorm) return { combustivel: String(data[i][4] || "").trim(), categoria: String(data[i][5] || "").trim() };
  }
  return { combustivel: "", categoria: "" };
}

function _calcPercDeducaoCIVA(combustivel, categoria) {
  var percDeducao = 0;
  var comb = (combustivel || "").toLowerCase();
  var cat = (categoria || "").toLowerCase();
  if (comb.indexOf("gasóleo") >= 0 || comb.indexOf("gasoleo") >= 0) percDeducao = (cat === "pesado") ? 1.0 : 0.50;
  else if (comb.indexOf("elétrico") >= 0 || comb.indexOf("eletrico") >= 0) percDeducao = 1.0;
  else if (comb.indexOf("gpl") >= 0) percDeducao = 0.50;
  else percDeducao = 0;
  return percDeducao;
}

function _parseDataDocCusto(dataStr) {
  if (!dataStr || typeof dataStr !== "string") return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var s = String(dataStr).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return String(m[1]).padStart(2, "0") + "/" + String(m[2]).padStart(2, "0") + "/" + m[3];
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function _saveFleetInvoiceToDrive(imageBase64, dataStr, matricula, fornecedor, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.folderId) return { success: false, error: "Folder do cliente não encontrado." };
    var folder = DriveApp.getFolderById(ctx.folderId);
    var subFolders = folder.getFoldersByName("Flowly_Fleet_Invoices");
    var targetFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder("Flowly_Fleet_Invoices");
    var dataNorm = (dataStr || "").replace(/\//g, "-").replace(/^(\d{2})-(\d{2})-(\d{4})$/, "$3-$2-$1");
    if (!dataNorm.match(/^\d{4}-\d{2}-\d{2}$/)) dataNorm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var matNorm = (matricula || "N/A").replace(/[\\/:*?"<>|]/g, "_").substring(0, 20);
    var fornNorm = (fornecedor || "N/A").replace(/[\\/:*?"<>|]/g, "_").substring(0, 30);
    var fileName = dataNorm + "_" + matNorm + "_" + fornNorm + ".jpg";
    var split = (imageBase64 || "").split(",");
    var mimeType = "image/jpeg";
    var cleanB64 = imageBase64;
    if (split.length >= 2) { var m = split[0].match(/:(.*?);/); if (m) mimeType = m[1]; cleanB64 = split[1]; }
    var ext = mimeType.indexOf("png") >= 0 ? "png" : "jpg";
    if (fileName.endsWith(".jpg") && ext === "png") fileName = fileName.replace(".jpg", ".png");
    var blob = Utilities.newBlob(Utilities.base64Decode(cleanB64), mimeType, fileName);
    var file = targetFolder.createFile(blob);
    return { success: true, url: file.getUrl() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveCustoFrota(payload, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ensureCCSheetWithHeaders(ss);
    var userEmail = Session.getActiveUser().getEmail();
    var docId = Utilities.getUuid();
    var now = new Date();
    var dataStr = _parseDataDocCusto(payload.dataDoc || "");
    var matricula = (payload.matricula || "").trim();
    var tipo = (payload.tipo || "").trim();
    var artigo = tipo || "Despesa Frota";
    if (tipo === "Abastecimento") artigo = "Gasóleo";
    else if (tipo === "Manutenção/Oficina") artigo = "Manutenção";
    else if (tipo === "Portagens") artigo = "Portagens";
    var valorTotal = parseFloat(payload.valorTotal) || 0;
    var kmAtuais = (payload.kmAtuais !== "" && payload.kmAtuais != null) ? parseFloat(payload.kmAtuais) : "";
    var litros = (payload.litros !== "" && payload.litros != null) ? parseFloat(payload.litros) : "";
    var fornecedor = (payload.fornecedor || "").trim();
    var observacoes = (payload.observacoes || "").trim();
    if (tipo === "Portagens" && !fornecedor) fornecedor = "VIA VERDE";

    var linkFoto = "";
    var imageBase64 = payload.imageBase64 || (payload.ocrData && payload.ocrData.imageBase64) || "";
    if (imageBase64 && imageBase64.length > 50) {
      var uploadRes = _saveFleetInvoiceToDrive(imageBase64, dataStr, matricula, fornecedor, impersonateTarget);
      if (uploadRes.success && uploadRes.url) linkFoto = uploadRes.url;
    }

    var valorIva = parseFloat(payload.valorIva) || 0;
    if ((isNaN(valorIva) || valorIva < 0 || valorIva === 0) && valorTotal > 0) valorIva = valorTotal - (valorTotal / 1.23);
    valorIva = Math.round(valorIva * 100) / 100;

    var combustivel = (_getVeiculoByMatricula(ss, matricula).combustivel || "").trim();
    var isDedutivel = "Não";
    if (combustivel.indexOf("Gasóleo") >= 0 || combustivel.indexOf("Gasoleo") >= 0 || combustivel.indexOf("Elétrico") >= 0 || combustivel.indexOf("Eletrico") >= 0 || combustivel.indexOf("GPL") >= 0 || tipo === "Manutenção/Oficina") {
      isDedutivel = "Sim";
    }

    var statusPagamento = (payload.statusPagamento || "").toString().trim();
    statusPagamento = (statusPagamento === "Aberto") ? "Aberto" : "Pago";
    var idEntidade = getEntityIdByName(fornecedor, "Despesa", impersonateTarget);
    var rowData = [
      dataStr, "Despesa", "App Motorista", fornecedor, artigo, 1, valorTotal, 0, "23%", valorIva, isDedutivel, "VALIDADO", observacoes, linkFoto,
      now, userEmail, statusPagamento, dataStr, valorTotal, "", docId, matricula, isNaN(kmAtuais) ? "" : kmAtuais, isNaN(litros) ? "" : litros, idEntidade, ""
    ];
    rowData[24] = idEntidade;
    sheet.appendRow(rowData);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _getKmAnteriorAbastecimento(allRows, matricula, currentIdx) {
  var matNorm = (matricula || "").toString().trim().toUpperCase();
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i].idx >= currentIdx) continue;
    var r = allRows[i].row;
    var mat = (r[21] || "").toString().trim().toUpperCase();
    var artigo = (r[4] || "").toString().trim();
    if (mat === matNorm && (artigo === "Gasóleo" || artigo === "Abastecimento")) {
      var km = parseFloat(r[22]);
      return (!isNaN(km) && km > 0) ? km : null;
    }
  }
  return null;
}

function _getKmAnteriorFromHistorico(historico, matricula, currentKm, currentIdx) {
  var matNorm = (matricula || "").toString().trim().toUpperCase();
  for (var i = currentIdx + 1; i < historico.length; i++) {
    var h = historico[i];
    var mat = (h.matricula || "").toString().trim().toUpperCase();
    if (mat === matNorm && (h.tipo === "Gasóleo" || h.tipo === "Abastecimento") && h.km != null && h.km > 0) {
      return h.km;
    }
  }
  return null;
}

function getCustosFrota(impersonateTarget, limit) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return (h || "").toString().trim(); });
    var idxMatricula = _indexOfHeader(headers, ["Matricula", "Matrícula"]);
    if (idxMatricula === -1) idxMatricula = 21;
    var idxKm = _indexOfHeader(headers, ["Km_Atuais", "Km Atuais", "Km"]);
    if (idxKm === -1) idxKm = 22;
    var idxLitros = _indexOfHeader(headers, ["Litros"]);
    if (idxLitros === -1) idxLitros = 23;
    var idxData = _indexOfHeader(headers, ["Data"]);
    if (idxData === -1) idxData = 0;
    var idxArtigo = _indexOfHeader(headers, ["Artigo"]);
    if (idxArtigo === -1) idxArtigo = 4;
    var idxValor = _indexOfHeader(headers, ["Valor Pago", "ValorPago", "Valor_Pago", "Custo_Total"]);
    if (idxValor === -1) idxValor = 18;
    var idxLinkFoto = _indexOfHeader(headers, ["Link_Foto", "Link", "LinkFoto"]);
    if (idxLinkFoto === -1) idxLinkFoto = 13;
    var historico = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var mat = idxMatricula !== -1 ? (row[idxMatricula] || "").toString().trim() : "";
      if (!mat) continue;
      var artigo = idxArtigo !== -1 ? (row[idxArtigo] || "").toString().trim() : "";
      var tipoExib = artigo || "Despesa";
      var valorRaw = idxValor !== -1 ? row[idxValor] : (row[18] || row[7] || 0);
      var valor = parseFloat(valorRaw);
      if (isNaN(valor)) valor = parseFloat(row[7]) || 0;
      var km = idxKm !== -1 && row[idxKm] !== "" && row[idxKm] != null ? parseFloat(row[idxKm]) : null;
      var litros = idxLitros !== -1 && row[idxLitros] !== "" && row[idxLitros] != null ? parseFloat(row[idxLitros]) : null;
      var dataStr = idxData !== -1 && row[idxData] ? (row[idxData] instanceof Date ? Utilities.formatDate(row[idxData], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[idxData])) : "";
      var linkFoto = idxLinkFoto !== -1 && row[idxLinkFoto] ? String(row[idxLinkFoto] || "").trim() : "";
      historico.push({ data: dataStr, matricula: mat, tipo: tipoExib, km: km, litros: litros, valor: isNaN(valor) ? 0 : valor, custoTotal: isNaN(valor) ? null : valor, consumoMedio: "N/A", linkFoto: linkFoto || null });
    }
    historico.reverse();
    var n = (limit && limit > 0) ? Math.min(limit, historico.length) : Math.min(20, historico.length);
    var out = historico.slice(0, n);
    for (var j = 0; j < out.length; j++) {
      var item = out[j];
      if ((item.tipo === "Gasóleo" || item.tipo === "Abastecimento") && item.litros > 0 && item.km > 0) {
        var kmAnteriores = _getKmAnteriorFromHistorico(historico, item.matricula, item.km, j);
        if (kmAnteriores !== null && item.km > kmAnteriores) item.consumoMedio = String(Math.round((item.litros / (item.km - kmAnteriores)) * 100 * 100) / 100);
      }
    }
    return out;
  } catch (e) { return []; }
}

function _categoriaGastoFrota(artigo, tipo, observacoes, fornecedor) {
  var s = (artigo || "").toString().toLowerCase() + " " + (tipo || "").toString().toLowerCase() + " " + (observacoes || "").toString().toLowerCase();
  var fornNorm = (fornecedor || "").toString().toUpperCase();
  if (/GALP|PETROGAL/.test(fornNorm)) { if (!/loja|cafetaria/.test(s)) return "combustivel"; }
  if (/gasoleo|gasóleo|gasolina|diesel|abastecimento|combustivel/.test(s)) return "combustivel";
  if (/portagem|via verde|portagens/.test(s)) return "portagens";
  var keywordsManutencao = /revis[aã]o|oficina|mec[aâ]nica|[oó]leo|filtro|pneus?|alinhamento|trav[oõ]es|cal[cç]os|embraiagem|correia|bateria|l[aâ]mpada|check-?up/;
  if (keywordsManutencao.test(s)) return "manutencao";
  return "outros";
}

function getDashboardFrotaData(impersonateTarget, batch) {
  var emptyFrota = { custoTotalMes: 0, custoPorKm: null, kmTotais: 0, consumoMedioGlobal: null, rankingConsumo: [], alertasCount: 0, veiculosComAlertas: [], proximaRevisao: [], racioGastos: {}, gastosPorCategoria: { combustivel: 0, manutencao: 0, portagens: 0, outros: 0 }, gastosPorFornecedor: {}, fleetSummaryText: "" };
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return emptyFrota;
    var data = [];
    var headers = [];
    if (batch && batch.dataNew && batch.dataNew.length > 0) { data = batch.dataNew; headers = batch.headersNew || []; } else {
      var ss = SpreadsheetApp.openById(ctx.sheetId);
      var sheet = ss.getSheetByName(SHEET_TAB_NAME);
      if (!sheet || sheet.getLastRow() < 2) return emptyFrota;
      var full = sheet.getDataRange().getValues();
      headers = full[0] ? full[0].map(function (h) { return (h || "").toString().trim(); }) : [];
      data = full.slice(1);
    }
    var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    var mesAtual = hoje.getMonth();
    var anoAtual = hoje.getFullYear();

    var custoTotalMes = 0, totalLitros = 0, totalKm = 0;
    var porMatricula = {};
    var gastosPorCategoria = { combustivel: 0, manutencao: 0, portagens: 0, outros: 0 };
    var gastosPorFornecedor = {};

    var idxMatricula = _indexOfHeader(headers, ["Matricula", "Matrícula"]); if (idxMatricula === -1) idxMatricula = 21;
    var idxKm = _indexOfHeader(headers, ["Km_Atuais", "Km Atuais", "Km"]); if (idxKm === -1) idxKm = 22;
    var idxLitros = _indexOfHeader(headers, ["Litros"]); if (idxLitros === -1) idxLitros = 23;
    var idxData = _indexOfHeader(headers, ["Data"]); if (idxData === -1) idxData = 0;
    var idxValor = _indexOfHeader(headers, ["Valor Pago", "ValorPago", "Valor_Pago", "Custo_Total"]); if (idxValor === -1) idxValor = 18;
    var idxQty = _indexOfHeader(headers, ["Quantidade"]); if (idxQty === -1) idxQty = 5;
    var idxArtigo = _indexOfHeader(headers, ["Artigo"]); if (idxArtigo === -1) idxArtigo = 4;
    var idxTipo = _indexOfHeader(headers, ["Tipo"]); if (idxTipo === -1) idxTipo = 1;
    var idxObservacoes = _indexOfHeader(headers, ["Observacoes", "Obs", "Observações"]); if (idxObservacoes === -1) idxObservacoes = 12;
    var idxFornecedor = _indexOfHeader(headers, ["Fornecedor", "fornecedor"]); if (idxFornecedor === -1) idxFornecedor = 3;

    var _extractKmFromObservacoes = function (obs) {
      if (!obs || typeof obs !== "string") return null;
      var m = obs.match(/(?:^|\s)(\d{1,7})\s*km\b/i) || obs.match(/\bkm[:\s]*(\d{1,7})\b/i);
      if (m && m[1]) { var n = parseFloat(m[1]); return !isNaN(n) && n > 0 ? n : null; }
      return null;
    };

    var dataFiltered = data.filter(function (r) {
      var d = r[idxData];
      if (typeof d === "string") { var p = (d || "").split(/[\/\-]/); if (p.length >= 3) d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10)); }
      return (d instanceof Date && !isNaN(d.getTime()) && d.getMonth() === mesAtual && d.getFullYear() === anoAtual);
    });

    var i, row, mat, d, valorRaw, valor, artigo, tipo, obs, cat, km, litros;
    for (i = 0; i < dataFiltered.length; i++) {
      row = dataFiltered[i];
      mat = idxMatricula !== -1 ? (row[idxMatricula] || "").toString().trim() : "";
      if (!mat) continue;
      d = row[idxData];
      if (typeof d === "string") { var p = (d || "").split(/[\/\-]/); if (p.length >= 3) d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10)); }
      if (!(d instanceof Date) || isNaN(d.getTime())) continue;

      valorRaw = idxValor !== -1 ? row[idxValor] : (row[18] || row[7] || 0);
      valor = parseFloat(valorRaw);
      if (isNaN(valor)) valor = (parseFloat(row[6]) || 0) * (parseFloat(row[idxQty]) || 1);
      custoTotalMes += valor;

      artigo = idxArtigo !== -1 ? (row[idxArtigo] || "").toString().trim() : "";
      tipo = idxTipo !== -1 ? (row[idxTipo] || "").toString().trim() : "";
      obs = idxObservacoes !== -1 ? (row[idxObservacoes] || "").toString().trim() : "";
      var forn = idxFornecedor !== -1 ? (row[idxFornecedor] || "").toString().trim() : "";
      forn = canonicalizeEntityForDisplay(forn) || forn;
      cat = _categoriaGastoFrota(artigo, tipo, obs, forn);
      gastosPorCategoria[cat] = (gastosPorCategoria[cat] || 0) + valor;
      if (forn) { gastosPorFornecedor[forn] = (gastosPorFornecedor[forn] || 0) + valor; }

      km = (idxKm !== -1 && row[idxKm] !== "" && row[idxKm] != null) ? parseFloat(row[idxKm]) : null;
      if (km == null || isNaN(km)) km = _extractKmFromObservacoes(obs);
      litros = idxLitros !== -1 && row[idxLitros] !== "" && row[idxLitros] != null ? parseFloat(row[idxLitros]) : null;

      if (!porMatricula[mat]) porMatricula[mat] = { litros: 0, km: [], valor: 0 };
      porMatricula[mat].valor += valor;
      if (litros && litros > 0) porMatricula[mat].litros += litros;
      if (km != null && !isNaN(km)) porMatricula[mat].km.push(km);
    }

    var veiculos = [];
    if (batch && batch.dataFrota && batch.dataFrota.length > 0) {
      for (var vi = 0; vi < batch.dataFrota.length; vi++) {
        var r = batch.dataFrota[vi];
        if (!r[0] && !r[1]) continue;
        veiculos.push({
          id: r[0] || "", matricula: r[1] || "", marcaModelo: r[2] || "", ano: r[3] || "", combustivel: r[4] || "",
          categoria: r[5] || "", lotacao: r[6] || "",
          dataInspecao: r[7] ? (r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[7])) : "",
          validadeSeguro: r[8] ? (r[8] instanceof Date ? Utilities.formatDate(r[8], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[8])) : "",
          proximaRevisao: r[9] ? (r[9] instanceof Date ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[9])) : "",
          status: r[10] || "Ativo",
          ultimaLavagem: r[11] ? (r[11] instanceof Date ? Utilities.formatDate(r[11], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[11]).trim()) : ""
        });
      }
    } else { veiculos = getVeiculos(impersonateTarget); }

    var kmPorMatricula = {};
    for (var m in porMatricula) {
      var kms = porMatricula[m].km || [];
      if (kms.length >= 2) {
        var kmMin = Math.min.apply(null, kms);
        var kmMax = Math.max.apply(null, kms);
        var delta = kmMax - kmMin;
        if (delta > 0) { kmPorMatricula[m] = delta; totalKm += delta; }
      }
      totalLitros += porMatricula[m].litros || 0;
    }

    var consumoMedioGlobal = (totalKm > 0 && totalLitros > 0) ? Math.round((totalLitros / totalKm) * 100 * 100) / 100 : null;

    var rankingConsumo = [];
    for (var m in porMatricula) {
      var lit = porMatricula[m].litros || 0;
      var kmDriven = kmPorMatricula[m] || 0;
      var consumo = (kmDriven > 0 && lit > 0) ? Math.round((lit / kmDriven) * 100 * 100) / 100 : null;
      rankingConsumo.push({ matricula: m, consumoL100km: consumo, litros: lit, km: kmDriven });
    }
    rankingConsumo.sort(function (a, b) {
      var ca = a.consumoL100km, cb = b.consumoL100km;
      if (ca == null && cb == null) return 0;
      if (ca == null) return 1;
      if (cb == null) return -1;
      return cb - ca;
    });

    var alertasCount = 0, veiculosComAlertas = [];
    var margem = MARGEM_ALERTA_IPO_SEGURO;
    for (var v = 0; v < veiculos.length; v++) {
      var vv = veiculos[v];
      var diasIPO = _diasAteData(vv.dataInspecao, hoje);
      var diasSeguro = _diasAteData(vv.validadeSeguro, hoje);
      var dataLavStr = vv.ultimaLavagem || "";
      var diasDesdeLavagem = dataLavStr ? -(_diasAteData(dataLavStr, hoje) || 0) : 999;
      var temAlerta = (diasIPO !== null && diasIPO <= margem) || (diasSeguro !== null && diasSeguro <= margem) || (diasDesdeLavagem === null || diasDesdeLavagem > 7);
      if (temAlerta) {
        alertasCount++;
        var statusSemaforo = "laranja";
        if ((diasIPO !== null && diasIPO <= 7) || (diasSeguro !== null && diasSeguro <= 7)) statusSemaforo = "vermelho";
        veiculosComAlertas.push({
          matricula: vv.matricula, marcaModelo: vv.marcaModelo, dataInspecao: vv.dataInspecao, validadeSeguro: vv.validadeSeguro,
          ultimaLavagem: vv.ultimaLavagem, diasIPO: diasIPO, diasSeguro: diasSeguro,
          alertaIPO: diasIPO !== null && diasIPO <= margem, alertaSeguro: diasSeguro !== null && diasSeguro <= margem,
          alertaLavagem: diasDesdeLavagem === null || diasDesdeLavagem > 7, statusSemaforo: statusSemaforo
        });
      }
    }

    var custoPorKm = (totalKm > 0 && custoTotalMes > 0) ? Math.round((custoTotalMes / totalKm) * 10000) / 10000 : null;
    var kmTotais = Math.round(totalKm);

    var proximaRevisao = [];
    for (var v2 = 0; v2 < veiculos.length; v2++) {
      var vv2 = veiculos[v2];
      var prStr = String(vv2.proximaRevisao || "").trim();
      if (/[\-\/]/.test(prStr)) continue;
      var prKm = parseFloat(prStr);
      if (isNaN(prKm) || prKm <= 0) continue;
      var kmAtual = null;
      if (porMatricula[vv2.matricula] && porMatricula[vv2.matricula].km && porMatricula[vv2.matricula].km.length > 0) {
        kmAtual = Math.max.apply(null, porMatricula[vv2.matricula].km);
      }
      if (kmAtual == null) continue;
      var diff = prKm - kmAtual;
      if (diff < 2000 && diff > 0) proximaRevisao.push({ matricula: vv2.matricula, marcaModelo: vv2.marcaModelo, kmAtuais: kmAtual, proximaRevisaoKm: prKm, kmRestantes: Math.round(diff) });
    }

    var totalGastos = gastosPorCategoria.combustivel + gastosPorCategoria.manutencao + gastosPorCategoria.portagens + (gastosPorCategoria.outros || 0);
    var racioGastos = { combustivel: 0, manutencao: 0, portagens: 0 };
    if (totalGastos > 0) {
      racioGastos.combustivel = Math.round((gastosPorCategoria.combustivel / totalGastos) * 1000) / 10;
      racioGastos.manutencao = Math.round((gastosPorCategoria.manutencao / totalGastos) * 1000) / 10;
      racioGastos.portagens = Math.round((gastosPorCategoria.portagens / totalGastos) * 1000) / 10;
    }

    var numVeiculos = veiculos.length;
    var maisIneficiente = rankingConsumo.length > 0 && rankingConsumo[0].consumoL100km != null ? rankingConsumo[0].matricula : null;
    var proximaRevisaoStr = proximaRevisao.length > 0 ? proximaRevisao.map(function (p) { return p.matricula + " (" + p.kmRestantes + " km)"; }).join(", ") : "";
    var racioStr = totalGastos > 0 ? " Rácio gastos: Combustível " + racioGastos.combustivel + "%, Manutenção " + racioGastos.manutencao + "%, Portagens " + racioGastos.portagens + "%." : "";
    var numVeiculosStr = "Número de veículos na frota: " + numVeiculos + ".";
    if (numVeiculos === 1) {
      var v1 = veiculos[0];
      var dIPO = _diasAteData(v1.dataInspecao, hoje);
      var dSeg = _diasAteData(v1.validadeSeguro, hoje);
      var diasIPOStr = dIPO !== null ? String(dIPO) : "N/A";
      var diasSegStr = dSeg !== null ? String(dSeg) : "N/A";
      numVeiculosStr += " Frota com apenas 1 veículo. INSTRUÇÕES PARA IA: 1) Ignora recomendações genéricas de condução eficiente ou otimização de rotas, a menos que haja um aumento súbito de consumo face ao histórico do próprio carro. 2) Prioriza alertas de calendário: Indica quantos dias faltam para o IPO e Seguro. Dias até IPO: " + diasIPOStr + ". Dias até Seguro: " + diasSegStr + ". 3) Foca na análise financeira: Detalha onde foi gasto o dinheiro (ex: '54% em manutenção é invulgar para este mês, verifica as faturas lidas').";
    }
    var funcionalidadesExistentes = " A App JÁ possui: Controlo de Custos, Alertas 30 dias (IPO/Seguro), Registo de Lavagens. Não sugiras estas funcionalidades como novas.";
    var fleetSummaryText = numVeiculosStr + " " + funcionalidadesExistentes + " Frota: custo total este mês " + custoTotalMes.toFixed(2) + "€. Custo por km " + (custoPorKm != null ? custoPorKm.toFixed(4) + "€/km" : "N/A") + ". Consumo médio global " + (consumoMedioGlobal != null ? consumoMedioGlobal + " L/100km" : "N/A") + ". KM totais: " + kmTotais + "." +
      (maisIneficiente ? " Veículo mais ineficiente: " + maisIneficiente + "." : "") +
      (proximaRevisaoStr ? " Revisão em breve (<2000 km): " + proximaRevisaoStr + "." : "") + racioStr +
      " Alertas ativos: " + alertasCount + " viatura(s) com IPO/Seguro em breve ou Lavagem em atraso. FOCO: Sugere melhorias na condução, rotas ou manutenção preventiva com base nos dados.";

    return {
      custoTotalMes: Math.round(custoTotalMes * 100) / 100, custoPorKm: custoPorKm, kmTotais: kmTotais, consumoMedioGlobal: consumoMedioGlobal,
      rankingConsumo: rankingConsumo, alertasCount: alertasCount, veiculosComAlertas: veiculosComAlertas, proximaRevisao: proximaRevisao,
      racioGastos: racioGastos, gastosPorCategoria: gastosPorCategoria, gastosPorFornecedor: gastosPorFornecedor, fleetSummaryText: fleetSummaryText
    };
  } catch (e) { return emptyFrota; }
}

function checkFrotaAlerts() {
  try {
    var ssMaster = SpreadsheetApp.openById(MASTER_DB_ID);
    var masterSheet = ssMaster.getSheets()[0];
    var data = masterSheet.getDataRange().getValues();
    var alertas = [];
    var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var sheetId = row[2];
      var status = String(row[4] || "").trim();
      var clientName = String(row[0] || "").trim();
      if (!sheetId || (status !== "Ativo" && status !== "")) continue;
      try {
        var ss = SpreadsheetApp.openById(sheetId);
        var frotaSheet = ss.getSheetByName(FROTA_VEICULOS_TAB);
        if (!frotaSheet || frotaSheet.getLastRow() < 2) continue;
        var numCols = Math.max(12, frotaSheet.getLastColumn());
        var rawData = frotaSheet.getRange(2, 1, frotaSheet.getLastRow(), numCols).getValues();
        for (var j = 0; j < rawData.length; j++) {
          var r = rawData[j];
          var matricula = (r[1] || "").toString().trim();
          var marcaModelo = (r[2] || "").toString().trim();
          if (!matricula) continue;
          var dataInspecao = r[7];
          var validadeSeguro = r[8];
          var ultimaLavagem = r[11];
          var dataInsStr = dataInspecao instanceof Date ? Utilities.formatDate(dataInspecao, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dataInspecao || "").trim();
          var dataSegStr = validadeSeguro instanceof Date ? Utilities.formatDate(validadeSeguro, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(validadeSeguro || "").trim();
          var dataLavStr = ultimaLavagem instanceof Date ? Utilities.formatDate(ultimaLavagem, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(ultimaLavagem || "").trim();
          var diasIns = _diasAteData(dataInsStr, hoje);
          var diasSeg = _diasAteData(dataSegStr, hoje);
          var diasDesdeLavagem = dataLavStr ? -(_diasAteData(dataLavStr, hoje) || 0) : 999;
          if (diasIns !== null && diasIns <= MARGEM_ALERTA_IPO_SEGURO) alertas.push({ cliente: clientName, matricula: matricula, veiculo: marcaModelo, tipo: "Inspeção (ITV)", data: dataInsStr, dias: diasIns });
          if (diasSeg !== null && diasSeg <= MARGEM_ALERTA_IPO_SEGURO) alertas.push({ cliente: clientName, matricula: matricula, veiculo: marcaModelo, tipo: "Seguro", data: dataSegStr, dias: diasSeg });
          if (diasDesdeLavagem > 7) alertas.push({ cliente: clientName, matricula: matricula, veiculo: marcaModelo, tipo: "Lavagem", data: dataLavStr || "Nunca", dias: diasDesdeLavagem });
        }
      } catch (e) { }
    }
    if (alertas.length > 0) {
      var dest = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
      if (dest) {
        var corpo = "Os seguintes veículos têm vencimentos próximos ou necessitam de atenção:\n\n";
        for (var k = 0; k < alertas.length; k++) {
          var a = alertas[k];
          if (a.tipo === "Lavagem") corpo += "⚠️ Viatura " + a.matricula + " não é lavada há mais de uma semana! (" + (a.data === "Nunca" ? "nunca lavada" : a.dias + " dias") + ")\n";
          else corpo += "• " + (a.cliente || "Cliente") + " | " + a.matricula + " | " + a.tipo + " | " + a.data + " (" + a.dias + " dias)\n";
        }
        MailApp.sendEmail(dest, "[Flowly Fleet] ⚠️ Alertas de Viaturas", corpo);
      }
    }
  } catch (e) { }
}

function createDailyFrotaTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkFrotaAlerts") ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("checkFrotaAlerts").timeBased().everyDays(1).atHour(8).create();
  return { success: true, message: "Trigger diário criado para checkFrotaAlerts às 08h00." };
}