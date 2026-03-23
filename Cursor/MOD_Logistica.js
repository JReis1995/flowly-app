// Ficheiro: MOD_Logistica.js
/// ==========================================
// 📦 MÓDULO DE LOGÍSTICA REVERSA E IMPORTAÇÕES
// ==========================================

function initVasilhameDB(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(LOGISTICA_VASILHAME_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(LOGISTICA_VASILHAME_TAB);
      sheet.appendRow(["Data", "Cliente", "Artigo", "Entregue_Qtd", "Devolvido_Qtd", "Observacoes", "Preco_Unitario_Doc"]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    } else if (sheet.getLastColumn() < 7) {
      sheet.getRange(1, 7).setValue("Preco_Unitario_Doc");
      sheet.getRange(1, 7).setFontWeight("bold");
    }
    return { success: true, sheet: sheet };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function initVasilhameConfig(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var sheet = ss.getSheetByName(CONFIG_VASILHAME_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG_VASILHAME_TAB);
      sheet.appendRow(["Nome_Vasilhame"]);
      sheet.getRange(1, 1).setFontWeight("bold");
    }
    if (sheet.getLastRow() < 2) {
      for (var i = 0; i < DEFAULT_VASILHAME_NAMES.length; i++) sheet.appendRow([DEFAULT_VASILHAME_NAMES[i]]);
    }
    return { success: true, sheet: sheet };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getVasilhameConfigs(impersonateTarget) {
  try {
    var init = initVasilhameConfig(impersonateTarget);
    if (!init.success || !init.sheet) return DEFAULT_VASILHAME_NAMES.slice();
    var sheet = init.sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return DEFAULT_VASILHAME_NAMES.slice();
    var vals = sheet.getRange(2, 1, lastRow, 1).getValues();
    var out = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] || "").trim();
      if (v && out.indexOf(v) === -1) out.push(v);
    }
    return out.length > 0 ? out : DEFAULT_VASILHAME_NAMES.slice();
  } catch (e) { return DEFAULT_VASILHAME_NAMES.slice(); }
}

function addVasilhameConfig(nome, impersonateTarget) {
  try {
    var n = (nome || "").toString().trim();
    if (!n) return { success: false, error: "Nome inválido ou vazio." };
    var init = initVasilhameConfig(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Config_Vasilhame." };
    var existing = getVasilhameConfigs(impersonateTarget);
    if (existing.indexOf(n) >= 0) return { success: true };
    init.sheet.appendRow([n]);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _findUltimoPrecoVasilhameNew(ss, cliente, artigo) {
  try {
    var sheet = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var numCols = Math.max(10, sheet.getLastColumn());
    var data = sheet.getRange(2, 1, sheet.getLastRow(), numCols).getValues();
    var clienteNorm = (cliente || "").toString().trim().toLowerCase();
    var artigoNorm = (artigo || "").toString().trim().toLowerCase();
    if (!clienteNorm || !artigoNorm) return null;
    for (var i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      var forn = (r[3] || "").toString().trim().toLowerCase();
      var art = (r[4] || "").toString().trim().toLowerCase();
      if (forn.indexOf(clienteNorm) >= 0 || clienteNorm.indexOf(forn) >= 0) {
        if (art.indexOf(artigoNorm) >= 0 || artigoNorm.indexOf(art) >= 0) {
          var precoVenda = parseFloat(r[7]) || 0;
          var precoCusto = parseFloat(r[6]) || 0;
          var preco = precoVenda > 0 ? precoVenda : precoCusto;
          if (preco > 0) return { precoUnitario: preco, precoVenda: precoVenda };
        }
      }
    }
    return null;
  } catch (e) { return null; }
}

function saveVasilhame(dados, impersonateTarget) {
  try {
    var init = initVasilhameDB(impersonateTarget);
    if (!init.success || !init.sheet) return { success: false, error: init.error || "Erro ao inicializar Vasilhame." };
    var sheet = init.sheet;
    var tipo = (dados.tipoMovimento || "").trim();
    var qtd = parseFloat(dados.quantidade) || 0;
    if (qtd <= 0) return { success: false, error: "Quantidade inválida." };
    var entregue = (tipo === "Entregámos ao Cliente") ? qtd : 0;
    var devolvido = (tipo === "Recolhemos do Cliente" || tipo === "Cliente Devolveu") ? qtd : 0;
    var hoje = new Date();
    var dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var precoUnitarioDoc = parseFloat(dados.precoUnitarioDoc) || 0;

    var row = [dataStr, (dados.cliente || "").trim(), (dados.artigo || "").trim(), entregue, devolvido, (dados.observacoes || "").trim(), precoUnitarioDoc > 0 ? precoUnitarioDoc : ""];
    sheet.appendRow(row);

    if (dados.gerarEstorno && (tipo === "Recolhemos do Cliente" || tipo === "Cliente Devolveu")) {
      var ctx = getClientContext(impersonateTarget || null);
      if (ctx && ctx.sheetId) {
        var ss = SpreadsheetApp.openById(ctx.sheetId);
        var precoInfo = precoUnitarioDoc > 0 ? { precoUnitario: precoUnitarioDoc, precoVenda: precoUnitarioDoc } : _findUltimoPrecoVasilhameNew(ss, dados.cliente, dados.artigo);
        var valorEstorno = 0;
        if (precoInfo && precoInfo.precoUnitario > 0) valorEstorno = precoInfo.precoUnitario * qtd;
        if (valorEstorno > 0) {
          var sheetNew = ensureCCSheetWithHeaders(ss);
          var userEmail = Session.getActiveUser().getEmail();
          var docId = Utilities.getUuid();
          var obsEstorno = "Estorno vasilhame: " + qtd + "x " + (dados.artigo || "") + " recolhido de " + (dados.cliente || "");
          var rowEstorno = [dataStr, "Ajuste", "Vasilhame Recolhido", (dados.cliente || "").trim(), "Estorno " + (dados.artigo || ""), qtd, 0, valorEstorno, "0%", 0, "Não", "VALIDADO", obsEstorno, "", hoje, userEmail, "Pago", dataStr, valorEstorno, "", docId, "", "", "", ""];
          sheetNew.appendRow(rowEstorno);
        }
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getSaldosVasilhame(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var allowedNames = getVasilhameConfigs(impersonateTarget);
    var map = {};

    var sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    if (sheetNew && sheetNew.getLastRow() >= 2) {
      var dataNew = sheetNew.getRange(2, 1, sheetNew.getLastRow(), Math.max(6, sheetNew.getLastColumn())).getValues();
      for (var i = 0; i < dataNew.length; i++) {
        var artigo = String(dataNew[i][4] || "").trim();
        if (!artigo || !_isVasilhameWithList(artigo, allowedNames)) continue;
        var cliente = String(dataNew[i][3] || "").trim();
        if (!cliente) continue;
        var tipo = normalizeTipo(dataNew[i][1]);
        var qtd = parseFloat(dataNew[i][5]) || 0;
        if (qtd <= 0) continue;
        var artigoNorm = _normalizeVasilhameWithList(artigo, allowedNames) || artigo;
        var key = cliente + "|" + artigoNorm;
        if (!map[key]) map[key] = { cliente: cliente, artigo: artigoNorm, entregue: 0, devolvido: 0 };
        if (tipo === "saida") map[key].entregue += qtd;
        else if (tipo === "entrada") map[key].devolvido += qtd;
      }
    }

    var init = initVasilhameDB(impersonateTarget);
    if (init.success && init.sheet) {
      var sheetV = init.sheet;
      if (sheetV.getLastRow() >= 2) {
        var dataV = sheetV.getRange(2, 1, sheetV.getLastRow(), Math.max(7, sheetV.getLastColumn())).getValues();
        for (var j = 0; j < dataV.length; j++) {
          var cliente = String(dataV[j][1] || "").trim();
          var artigo = String(dataV[j][2] || "").trim();
          if (!cliente || !artigo || !_isVasilhameWithList(artigo, allowedNames)) continue;
          var artigoNorm = _normalizeVasilhameWithList(artigo, allowedNames) || artigo;
          var key = cliente + "|" + artigoNorm;
          if (!map[key]) map[key] = { cliente: cliente, artigo: artigoNorm, entregue: 0, devolvido: 0 };
          map[key].entregue += parseFloat(dataV[j][3]) || 0;
          map[key].devolvido += parseFloat(dataV[j][4]) || 0;
        }
      }
    }

    var out = [];
    for (var k in map) {
      var m = map[k];
      var saldo = m.entregue - m.devolvido;
      if (saldo !== 0) out.push({ cliente: m.cliente, artigo: m.artigo, saldo: Math.round(saldo * 100) / 100 });
    }
    return out;
  } catch (e) { return []; }
}

function getSaldosVasilhameComValor(impersonateTarget, batch) {
  try {
    var saldos = [];
    if (batch && batch.dataNew && batch.dataVasilhame) {
      var allowedNames = getVasilhameConfigs(impersonateTarget);
      var map = {};
      for (var i = 0; i < batch.dataNew.length; i++) {
        var r = batch.dataNew[i];
        var artigo = String(r[4] || "").trim();
        if (!artigo || !_isVasilhameWithList(artigo, allowedNames)) continue;
        var cliente = String(r[3] || "").trim();
        if (!cliente) continue;
        var tipo = normalizeTipo(r[1]);
        var qtd = parseFloat(r[5]) || 0;
        if (qtd <= 0) continue;
        var artigoNorm = _normalizeVasilhameWithList(artigo, allowedNames) || artigo;
        var key = cliente + "|" + artigoNorm;
        if (!map[key]) map[key] = { cliente: cliente, artigo: artigoNorm, entregue: 0, devolvido: 0 };
        if (tipo === "saida") map[key].entregue += qtd;
        else if (tipo === "entrada") map[key].devolvido += qtd;
      }
      for (var j = 0; j < batch.dataVasilhame.length; j++) {
        var rv = batch.dataVasilhame[j];
        var cliente = String(rv[1] || "").trim();
        var artigo = String(rv[2] || "").trim();
        if (!cliente || !artigo || !_isVasilhameWithList(artigo, allowedNames)) continue;
        var artigoNorm = _normalizeVasilhameWithList(artigo, allowedNames) || artigo;
        var key = cliente + "|" + artigoNorm;
        if (!map[key]) map[key] = { cliente: cliente, artigo: artigoNorm, entregue: 0, devolvido: 0 };
        map[key].entregue += parseFloat(rv[3]) || 0;
        map[key].devolvido += parseFloat(rv[4]) || 0;
      }
      for (var k in map) {
        var m = map[k];
        var saldo = m.entregue - m.devolvido;
        if (saldo !== 0) saldos.push({ cliente: m.cliente, artigo: m.artigo, saldo: Math.round(saldo * 100) / 100 });
      }
    } else { saldos = getSaldosVasilhame(impersonateTarget); }

    var valorTotalNaRua = 0;
    if (batch && batch.dataNew && batch.dataNew.length > 0) {
      for (var si = 0; si < saldos.length; si++) {
        var s = saldos[si];
        if (s.saldo <= 0) continue;
        var cNorm = (s.cliente || "").toString().trim().toLowerCase();
        var aNorm = (s.artigo || "").toString().trim().toLowerCase();
        var preco = 0;
        for (var ni = batch.dataNew.length - 1; ni >= 0; ni--) {
          var nr = batch.dataNew[ni];
          var forn = (nr[3] || "").toString().trim().toLowerCase();
          var art = (nr[4] || "").toString().trim().toLowerCase();
          if (!forn || !art) continue;
          if (forn.indexOf(cNorm) < 0 && cNorm.indexOf(forn) < 0) continue;
          if (art.indexOf(aNorm) < 0 && aNorm.indexOf(art) < 0) continue;
          var tipoN = normalizeTipo(nr[1]);
          if (tipoN !== "saida" && tipoN !== "entrada") continue;
          var precoV = parseFloat(nr[7]) || 0;
          var precoC = parseFloat(nr[6]) || 0;
          preco = precoV > 0 ? precoV : precoC;
          if (preco > 0) break;
        }
        valorTotalNaRua += s.saldo * preco;
      }
    } else {
      for (var i = 0; i < saldos.length; i++) {
        var s = saldos[i];
        if (s.saldo <= 0) continue;
        var movs = getMovimentosVasilhameAuditoria(s.cliente, s.artigo, impersonateTarget);
        var preco = 0;
        if (movs && movs.length > 0) preco = parseFloat(movs[movs.length - 1].precoUnitario) || 0;
        valorTotalNaRua += s.saldo * preco;
      }
    }
    return { saldos: saldos, valorTotalNaRua: Math.round(valorTotalNaRua * 100) / 100 };
  } catch (e) { return { saldos: [], valorTotalNaRua: 0 }; }
}

function getClientesVasilhame(impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var set = {};

    var sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    if (sheetNew && sheetNew.getLastRow() >= 2) {
      var d = sheetNew.getRange(2, 1, sheetNew.getLastRow(), 5).getValues();
      for (var i = 0; i < d.length; i++) { var c = String(d[i][3] || "").trim(); if (c) set[c] = true; }
    }
    var init = initVasilhameDB(impersonateTarget);
    if (init.success && init.sheet && init.sheet.getLastRow() >= 2) {
      var dv = init.sheet.getRange(2, 1, init.sheet.getLastRow(), 2).getValues();
      for (var j = 0; j < dv.length; j++) { var cv = String(dv[j][1] || "").trim(); if (cv) set[cv] = true; }
    }
    return Object.keys(set).sort();
  } catch (e) { return []; }
}

function getMovimentosVasilhameAuditoria(cliente, artigo, impersonateTarget) {
  try {
    var ctx = getClientContext(impersonateTarget || null);
    if (!ctx || !ctx.sheetId) return [];
    var ss = SpreadsheetApp.openById(ctx.sheetId);
    var allowedNames = getVasilhameConfigs(impersonateTarget);
    var clienteNorm = (cliente || "").toString().trim().toLowerCase();
    var artigoNorm = (artigo || "").toString().trim().toLowerCase();
    if (!clienteNorm || !artigoNorm) return [];
    var movimentos = [];

    var sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    if (sheetNew && sheetNew.getLastRow() >= 2) {
      var numCols = Math.max(10, sheetNew.getLastColumn());
      var dataNew = sheetNew.getRange(2, 1, sheetNew.getLastRow(), numCols).getValues();
      for (var i = 0; i < dataNew.length; i++) {
        var r = dataNew[i];
        var forn = (r[3] || "").toString().trim().toLowerCase();
        var art = (r[4] || "").toString().trim().toLowerCase();
        if (!_isVasilhameWithList(art, allowedNames)) continue;
        if (forn.indexOf(clienteNorm) < 0 && clienteNorm.indexOf(forn) < 0) continue;
        if (art.indexOf(artigoNorm) < 0 && artigoNorm.indexOf(art) < 0) continue;
        var qtd = parseFloat(r[5]) || 0;
        if (qtd <= 0) continue;
        var precoVenda = parseFloat(r[7]) || 0;
        var precoCusto = parseFloat(r[6]) || 0;
        var preco = precoVenda > 0 ? precoVenda : precoCusto;
        var tipo = normalizeTipo(r[1]);
        var dataVal = r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[0] || "");
        var tipoMov = tipo === "saida" ? "Entrega" : (tipo === "entrada" ? "Recolha" : "");
        if (!tipoMov) continue;
        movimentos.push({ data: dataVal, tipo: tipoMov, quantidade: qtd, precoUnitario: preco, valorTotal: preco * qtd, fonte: "OCR" });
      }
    }

    var init = initVasilhameDB(impersonateTarget);
    if (init.success && init.sheet) {
      var sheetV = init.sheet;
      if (sheetV.getLastRow() >= 2) {
        var dataV = sheetV.getRange(2, 1, sheetV.getLastRow(), Math.max(7, sheetV.getLastColumn())).getValues();
        for (var j = 0; j < dataV.length; j++) {
          var rv = dataV[j];
          var c = (rv[1] || "").toString().trim().toLowerCase();
          var a = (rv[2] || "").toString().trim().toLowerCase();
          if (c.indexOf(clienteNorm) < 0 && clienteNorm.indexOf(c) < 0) continue;
          if (a.indexOf(artigoNorm) < 0 && artigoNorm.indexOf(a) < 0) continue;
          var entregue = parseFloat(rv[3]) || 0;
          var devolvido = parseFloat(rv[4]) || 0;
          var precoDoc = parseFloat(rv[6]) || 0;
          var dataVal = rv[0] instanceof Date ? Utilities.formatDate(rv[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rv[0] || "");
          if (entregue > 0) movimentos.push({ data: dataVal, tipo: "Entrega", quantidade: entregue, precoUnitario: precoDoc, valorTotal: precoDoc * entregue, fonte: "Manual" });
          if (devolvido > 0) movimentos.push({ data: dataVal, tipo: "Recolha", quantidade: devolvido, precoUnitario: precoDoc, valorTotal: precoDoc * devolvido, fonte: "Manual" });
        }
      }
    }

    movimentos.sort(function (a, b) {
      var da = a.data || "";
      var db = b.data || "";
      return da.localeCompare(db) || 0;
    });
    return movimentos;
  } catch (e) { return []; }
}

function importBulkStock(csvData, tipoMovimento, impersonateTarget) {
  try {
    if (!Array.isArray(csvData) || csvData.length === 0) return { success: false, error: "Dados vazios.", imported: 0 };
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada.", imported: 0 };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);

    const maxCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, maxCols).getValues()[0].map(function (h) { return (h || "").toString().trim(); });

    const getColIdx = function (name) { return headers.indexOf(name); };
    const getColIdxAny = function (names) {
      for (var i = 0; i < names.length; i++) { var idx = headers.indexOf(names[i]); if (idx >= 0) return idx; }
      return -1;
    };

    const idxData = getColIdx("Data"), idxTipo = getColIdx("Tipo"), idxMetodo = getColIdx("Metodo");
    const idxFornecedor = getColIdxAny(["Fornecedor", "fornecedor"]), idxArtigo = getColIdx("Artigo");
    const idxQuantidade = getColIdxAny(["Qtd", "Quantidade"]), idxPrecoCusto = getColIdxAny(["Preco", "Preco_Custo"]);
    const idxPrecoVenda = getColIdxAny(["Venda", "Preco_Venda"]), idxTaxaIva = getColIdxAny(["TaxaIva", "Taxa iva (%)"]);
    const idxValorIva = getColIdxAny(["ValorIva", "Valor iva (€)"]), idxObservacoes = getColIdxAny(["Obs", "Observacoes"]);
    const idxDedutivel = getColIdxAny(["Dedutivel", "Dedutível", "Dedutível (%)"]), idxValidado = getColIdx("Validado");
    const idxLink = getColIdx("Link"), idxTimestamp = getColIdx("Timestamp"), idxUser = getColIdx("User");
    const idxStatus = getColIdx("Status"), idxDataPag = getColIdx("DataPag"), idxValorPago = getColIdx("ValorPago");
    const idxContaStock = getColIdx("ContaStock"), idxDocID = getColIdx("DocID");
    const idxIDEntidade = getColIdxAny(["ID_Entidade", "ID Entidade"]), idxPrecoSugerido = getColIdxAny(["Preço Sugerido", "Preco Sugerido"]);

    const existingArtigos = {};
    if (sheet.getLastRow() > 1 && idxArtigo >= 0) {
      const artigos = sheet.getRange(2, idxArtigo + 1, sheet.getLastRow(), idxArtigo + 1).getValues();
      artigos.forEach(function (r) { const a = (r[0] || "").toString().trim(); if (a) existingArtigos[a.toLowerCase()] = true; });
    }

    const now = new Date();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const userEmail = Session.getActiveUser().getEmail();

    const linhasFormatadas = [];
    csvData.forEach(function (item) {
      const toPT = (val) => val ? val.toString().replace('.', ',').trim() : "";
      const artigo = (item.artigo || "").toString().trim();
      if (!artigo) return;
      if (existingArtigos[artigo.toLowerCase()]) return;
      existingArtigos[artigo.toLowerCase()] = true;

      const qty = (item.quantidade !== undefined && item.quantidade !== null && item.quantidade !== "") ? (parseFloat(item.quantidade) || 0) : 0;
      const preco = (item.preco_custo !== undefined && item.preco_custo !== null && item.preco_custo !== "") ? (parseFloat(item.preco_custo) || 0) : 0;
      const fornecedor = (item.fornecedor || "").toString().trim();
      const categoria = (item.categoria || "").toString().trim();
      var valorIvaCalc = qty * preco * 0.23;
      const obsBase = categoria ? "Importação inicial | Categoria: " + categoria : "Importação inicial";
      var tipoFinal = tipoMovimento || "Entrada";
      if (item.tipo !== undefined && item.tipo !== null && String(item.tipo).trim() !== "") {
        var norm = normalizeMovementType(item.tipo);
        tipoFinal = norm.charAt(0).toUpperCase() + norm.slice(1);
      } else if (tipoMovimento === "Multi-tipo") { tipoFinal = "Entrada"; }
      if (tipoFinal === "Saída") tipoFinal = "Saida";
      const valorSugerido = tipoFinal === "Entrada" ? calcularPrecoSugerido(preco) : "";
      if (!item.valor_iva && tipoFinal === "Saida") {
        var precoVenda = (item.preco_venda !== undefined && item.preco_venda !== null && item.preco_venda !== "") ? (parseFloat(item.preco_venda) || 0) : 0;
        valorIvaCalc = qty * precoVenda * 0.23;
      }

      var dataStr = todayStr;
      if (item.data !== undefined && item.data !== null && String(item.data).trim() !== "") {
        var parsed = toDBDate(item.data);
        if (parsed) dataStr = parsed;
      }

      const row = new Array(Math.max(maxCols, 26)).fill("");
      if (idxData >= 0) row[idxData] = dataStr;
      if (idxTipo !== -1) row[idxTipo] = tipoFinal;
      if (idxMetodo >= 0) row[idxMetodo] = (item.metodo !== undefined && item.metodo !== null && String(item.metodo).trim() !== "") ? String(item.metodo).trim() : "Importação";
      if (idxValidado >= 0) row[idxValidado] = (item.validado !== undefined && item.validado !== null && String(item.validado).trim() !== "") ? String(item.validado).trim() : "VALIDADO";
      if (idxDedutivel >= 0) {
        const dedutivelExplicito = (item.dedutivel !== undefined && item.dedutivel !== null && String(item.dedutivel).trim() !== "");
        if (dedutivelExplicito) row[idxDedutivel] = String(item.dedutivel).trim();
        else row[idxDedutivel] = ["Saida", "Quebra", "Oferta", "FechoCaixa"].indexOf(tipoFinal) >= 0 ? "Não" : "Sim";
      }
      if (idxContaStock >= 0) row[idxContaStock] = (item.conta_stock !== undefined && item.conta_stock !== null && String(item.conta_stock).trim() !== "") ? String(item.conta_stock).trim() : "Sim";
      if (idxStatus >= 0) row[idxStatus] = (item.status !== undefined && item.status !== null && String(item.status).trim() !== "") ? String(item.status).trim() : "Aberto";
      if (idxValorPago >= 0) row[idxValorPago] = (item.valor_pago !== undefined && item.valor_pago !== null && item.valor_pago !== "") ? (parseFloat(item.valor_pago) || 0) : 0;
      if (idxArtigo >= 0) row[idxArtigo] = artigo;
      if (idxQuantidade !== -1 && item.quantidade !== undefined && item.quantidade !== null && item.quantidade !== "") row[idxQuantidade] = item.quantidade;
      else if (idxQuantidade >= 0) row[idxQuantidade] = qty;
      if (idxPrecoCusto !== -1 && item.preco_custo) row[idxPrecoCusto] = toPT(item.preco_custo);
      else if (idxPrecoCusto >= 0) row[idxPrecoCusto] = toPT(preco);
      if (idxPrecoVenda !== -1 && item.preco_venda) row[idxPrecoVenda] = toPT(item.preco_venda);
      if (idxTaxaIva !== -1 && item.taxa_iva) row[idxTaxaIva] = toPT(item.taxa_iva);
      else if (idxTaxaIva >= 0) row[idxTaxaIva] = "23%";
      if (idxValorIva !== -1 && item.valor_iva) row[idxValorIva] = toPT(item.valor_iva);
      else if (idxValorIva >= 0) row[idxValorIva] = toPT(valorIvaCalc);
      if (idxFornecedor !== -1 && item.fornecedor !== undefined && item.fornecedor !== null && item.fornecedor !== "") row[idxFornecedor] = item.fornecedor;
      else if (idxFornecedor >= 0) row[idxFornecedor] = fornecedor;
      if (idxObservacoes !== -1 && item.observacoes !== undefined && item.observacoes !== null && item.observacoes !== "") row[idxObservacoes] = item.observacoes;
      else if (idxObservacoes >= 0) row[idxObservacoes] = obsBase;
      if (idxLink >= 0) row[idxLink] = "";
      if (idxTimestamp >= 0) row[idxTimestamp] = now;
      if (idxUser >= 0) row[idxUser] = userEmail;
      if (idxDataPag >= 0) row[idxDataPag] = (item.data_pag !== undefined && item.data_pag !== null && String(item.data_pag).trim() !== "") ? toDBDate(item.data_pag) || "" : "";
      if (idxDocID >= 0) row[idxDocID] = Utilities.getUuid();
      var idEnt = getEntityIdByName(fornecedor, tipoFinal, impersonateTarget) || "";
      if (idxIDEntidade >= 0) row[idxIDEntidade] = idEnt; else row[24] = idEnt;
      if (idxPrecoSugerido >= 0) row[idxPrecoSugerido] = valorSugerido; else row[25] = valorSugerido;
      linhasFormatadas.push(row);
    });

    if (linhasFormatadas.length === 0) return { success: true, imported: 0, message: "Todos os artigos já existem." };
    sheet.getRange(sheet.getLastRow() + 1, 1, linhasFormatadas.length, Math.max(maxCols, 26)).setValues(linhasFormatadas);
    
    if (tipoMovimento === "Entrada" || tipoMovimento === "Multi-tipo") {
       try { if (typeof invalidateFTCostCache === "function") invalidateFTCostCache(impersonateTarget); } catch(e){}
    }
    return { success: true, imported: linhasFormatadas.length };
  } catch (e) { return { success: false, error: e.toString(), imported: 0 }; }
}

function saveLogisticaConsolidated(dados) {
  try {
    const docId = Utilities.getUuid();
    const ctx = getClientContext(dados.impersonateEmail);
    if (!ctx || !ctx.sheetId) return { success: false, error: "ERRO: Infraestrutura não encontrada." };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ensureCCSheetWithHeaders(ss);
    const userEmail = Session.getActiveUser().getEmail();

    let formattedDate = "";
    const rawDate = (dados.data || "").toString().trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(rawDate)) { formattedDate = toDBDateFromInput(rawDate); }
    else if (/^\d{2}\/\d{2}\/\d{4}$/.test(rawDate)) { formattedDate = rawDate; }

    let tipoCol = "Entrada";
    const t = String(dados.tipo || "").toLowerCase().trim();
    if (t === "saída" || t === "saida") tipoCol = "Saida";
    else if (t === "quebra") tipoCol = "Quebra";
    else if (t === "oferta") tipoCol = "Oferta";
    else if (t === "consumo") tipoCol = "Consumo";
    else if (t === "fecho de caixa/relatório" || t === "fechocaixa" || t === "fecho caixa") tipoCol = "FechoCaixa";
    else if (t === "despesas") tipoCol = "Despesas";
    else if (dados.tipo && String(dados.tipo).trim()) tipoCol = String(dados.tipo).trim();

    const fornecedorRaw = (dados.fornecedor || "").trim();
    const isFechoCaixa = (tipoCol === "FechoCaixa" || String(tipoCol).toLowerCase().indexOf("fecho") >= 0);
    const fornecedor = isFechoCaixa ? "" : fornecedorRaw;
    const idEntidade = isFechoCaixa ? "" : getEntityIdByName(fornecedorRaw, dados.tipo || tipoCol, dados.impersonateEmail);
    const contaStock = (dados.stock != null && String(dados.stock).trim() !== "") ? String(dados.stock).trim() : "Sim";
    const jaPago = dados.jaPago === true;
    const dataPagamento = (dados.dataPagamento && String(dados.dataPagamento).trim()) ? String(dados.dataPagamento).trim() : "";
    const faturaPaga = jaPago ? "Sim" : "Não";
    const status = jaPago ? "Pago" : "Aberto";

    const isSaidaOuSaida = (tipoCol === "Saida" || tipoCol === "Saída" || tipoCol === "Quebra" || tipoCol === "Oferta" || tipoCol === "Consumo" || tipoCol === "Despesas");

    (dados.artigos || []).forEach(function (i) {
      const qtd = parseFloat(String(i.quantidade || "").replace(",", ".")) || 0;
      const custo = parseFloat(String(i.preco_custo || "").replace(",", ".")) || 0;
      const taxa = normalizeTaxaIva(i.taxa_iva);
      const valorBase = custo;
      const valIVA = (i.valor_iva != null && !isNaN(parseFloat(i.valor_iva))) ? parseFloat(String(i.valor_iva).replace(",", ".")) : (qtd * valorBase) * (taxa / 100);
      const colG = isSaidaOuSaida ? 0 : custo;
      const colH = isSaidaOuSaida ? custo : 0;
      const valorSugerido = !isSaidaOuSaida ? calcularPrecoSugerido(custo) : "";
      const totalLinha = qtd * valorBase;
      const valorPago = jaPago ? totalLinha : 0;

      const rowData = [
        formattedDate, tipoCol, "Multi-OCR", fornecedor, (i.artigo || "").toString(),
        qtd, colG, colH, taxa + "%", valIVA, faturaPaga, "VALIDADO", "", "", new Date(),
        userEmail, status, jaPago ? dataPagamento : "", valorPago, contaStock,
        docId, "", "", "", idEntidade, valorSugerido
      ];
      sheet.appendRow(rowData);
    });

    // INTEGRAÇÃO: Processar "Saídas" automáticas para Fichas Técnicas
    if (isSaidaOuSaida) {
       var interceptResponse = null;
       try {
          if(typeof processComposeStock === "function") {
              interceptResponse = processComposeStock(dados.artigos, dados.impersonateEmail);
          }
       } catch(warn) {}
       
       if(interceptResponse && interceptResponse.success && interceptResponse.hasComposed && interceptResponse.saidas) {
           interceptResponse.saidas.forEach(function (si) {
               // Append generated exit
               sheet.appendRow([
                  formattedDate, "Saida (FT)", "Baixa Auto FT", si.fornecedor, si.artigo,
                  si.quantidade, 0, si.preco_custo, "0%", 0, "Não", "VALIDADO", "Artigo decrementado via venda de Ficha Técnica", "", new Date(),
                  userEmail, "Pago", formattedDate, 0, "Sim", Utilities.getUuid(), "", "", "", "", ""
               ]);
           });
       }
    } else if (tipoCol === "Entrada") {
       try { if (typeof invalidateFTCostCache === "function") invalidateFTCostCache(dados.impersonateEmail); } catch(e){}
    }

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function isVasilhame(nomeArtigo) {
  if (!nomeArtigo || typeof nomeArtigo !== "string") return false;
  var s = String(nomeArtigo).toLowerCase();
  for (var i = 0; i < LISTA_VASILHAME.length; i++) {
    if (s.indexOf(LISTA_VASILHAME[i].toLowerCase()) >= 0) return true;
  }
  return false;
}

function _normalizeArtigoVasilhame(nomeArtigo) {
  if (!nomeArtigo || typeof nomeArtigo !== "string") return "";
  var s = String(nomeArtigo).toLowerCase().trim();
  var best = "";
  for (var i = 0; i < LISTA_VASILHAME.length; i++) {
    var term = LISTA_VASILHAME[i];
    if (s.indexOf(term.toLowerCase()) >= 0 && term.length > best.length) best = term;
  }
  return best || s;
}

function _isVasilhameWithList(artigo, allowedList) {
  if (!artigo || !allowedList || allowedList.length === 0) return false;
  var s = String(artigo).toLowerCase();
  for (var i = 0; i < allowedList.length; i++) {
    if (s.indexOf(String(allowedList[i]).toLowerCase()) >= 0) return true;
  }
  return false;
}

function _normalizeVasilhameWithList(artigo, allowedList) {
  if (!artigo || !allowedList || allowedList.length === 0) return "";
  var s = String(artigo).toLowerCase().trim();
  var best = "";
  for (var i = 0; i < allowedList.length; i++) {
    var term = String(allowedList[i]);
    if (s.indexOf(term.toLowerCase()) >= 0 && term.length > best.length) best = term;
  }
  return best || s;
}

function normalizeMovementType(rawType) {
  if (rawType == null || rawType === "") return "entrada";
  var t = String(rawType).toLowerCase().trim()
    .replace(/[àáâãä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íîï]/g, 'i')
    .replace(/[óôõö]/g, 'o').replace(/[úûü]/g, 'u').replace(/ç/g, 'c');
  if (["despesas", "despesa", "gasto", "gastos", "consumo", "fixo", "fixos", "expense", "renda"].indexOf(t) >= 0) return "despesas";
  if (["entrada", "compras", "compra", "stock", "lote", "purchase"].indexOf(t) >= 0) return "entrada";
  if (["saida", "venda", "vendas", "sale"].indexOf(t) >= 0) return "saida";
  return "entrada";
}
function uploadInvoiceToDrive(base64Data, fileName, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.folderId) {
      return { success: false, error: "ERRO: Infraestrutura não encontrada. Garante que o teu email de Super Admin está registado como cliente na Master DB ou entra em Modo Espião." };
    }
    const targetFolderId = ctx.folderId;

    // Entra na pasta do cliente
    const folder = DriveApp.getFolderById(targetFolderId);

    // Procura a sub-pasta de faturas ou cria se não existir
    const folders = folder.getFoldersByName("Faturas_Digitalizadas");
    const targetFolder = folders.hasNext() ? folders.next() : folder.createFolder("Faturas_Digitalizadas");

    // Limpa o base64
    const split = base64Data.split(',');
    const mimeType = split[0].match(/:(.*?);/)[1];
    const cleanB64 = split[1];

    // Cria e guarda o ficheiro no Google Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(cleanB64), mimeType, fileName);
    const file = targetFolder.createFile(blob);

    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}