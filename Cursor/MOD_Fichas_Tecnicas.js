// Ficheiro: MOD_Fichas_Tecnicas.js
/// ==========================================
// 📝 MÓDULO DE FICHAS TÉCNICAS E RECEITUÁRIO
// ==========================================

const FICHAS_DB_TAB = "Fichas_Tecnicas_DB";
const FICHAS_ITENS_TAB = "Fichas_Tecnicas_Itens";

/**
 * Helper interno para inicializar as abas de Fichas Técnicas.
 * Retorna objetos Sheet para uso exclusivo no servidor.
 */
function _initFichasTecnicasTabs(impersonateEmail) {
  try {
    const ctx = getClientContext(impersonateEmail);
    if (!ctx || !ctx.sheetId) {
      console.error("[ERROR] getClientContext não retornou sheetId para:", impersonateEmail);
      return { success: false, error: "Erro de Configuração: SheetID não resolvido." };
    }
    
    console.log("[DEBUG] _initFichasTecnicasTabs -> Abrindo Planilha ID:", ctx.sheetId, "| Contexto para:", impersonateEmail);
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(ctx.sheetId);
      console.log("[DEBUG] Planilha aberta com sucesso:", ss.getName());
    } catch (e) {
      console.error("[ERROR] Falha ao abrir SpreadsheetApp.openById:", ctx.sheetId, e.toString());
      return { success: false, error: "ID inválido ou sem permissão: " + e.toString() };
    }
    
    let dbSheet = ss.getSheetByName(FICHAS_DB_TAB);
    let itensSheet = ss.getSheetByName(FICHAS_ITENS_TAB);
    let created = false;

    // Tentar criar DB se não existe
    if (!dbSheet) {
      console.log("[DEBUG] Aba DB não encontrada. Criando:", FICHAS_DB_TAB);
      try {
        dbSheet = ss.insertSheet(FICHAS_DB_TAB);
        dbSheet.appendRow(["ID_FT", "Nome_Artigo", "Categoria", "Modo_Preparacao", "Preco_Custo_Total", "Preco_Sugerido", "Link_Logo_Cliente", "Timestamp", "User"]);
        dbSheet.getRange(1, 1, 1, 9).setFontWeight("bold");
        created = true;
      } catch (e) {
        return { success: false, error: "Erro ao criar aba DB: " + e.toString() };
      }
    }
    
    // Tentar criar Itens se não existe
    if (!itensSheet) {
      console.log("[DEBUG] Aba Itens não encontrada. Criando:", FICHAS_ITENS_TAB);
      try {
        itensSheet = ss.insertSheet(FICHAS_ITENS_TAB);
        itensSheet.appendRow(["ID_FT", "Artigo_Ingrediente", "Quantidade_Necessaria", "Unidade", "Custo_Unitario_Ref"]);
        itensSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
        created = true;
      } catch (e) {
        return { success: false, error: "Erro ao criar aba Itens: " + e.toString() };
      }
    }

    if (created) SpreadsheetApp.flush();
    
    return { success: true, dbSheet: dbSheet, itensSheet: itensSheet, ss: ss, created: created };
  } catch (e) {
    console.error("[CRITICAL] Erro em _initFichasTecnicasTabs:", e.toString());
    return { success: false, error: "Erro crítico em _initFichasTecnicasTabs: " + e.toString() };
  }
}

/**
 * Função pública para o frontend inicializar as abas.
 * NÃO retorna objetos complexos (Sheet/Spreadsheet), garantindo estabilidade na comunicação.
 */
function initFichasTecnicasTabs(impersonateEmail) {
  if (!impersonateEmail) return { success: false, error: "ImpersonateEmail é obrigatório para inicializar abas." };
  const res = _initFichasTecnicasTabs(impersonateEmail);
  if (!res.success) return res;
  return { success: true, created: res.created || false };
}

function calculateFTCost(itens, impersonateEmail) {
  try {
    const ctx = getClientContext(impersonateEmail);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    let dataNew = [];
    if (sheetNew && sheetNew.getLastRow() >= 2) {
      dataNew = sheetNew.getRange(2, 1, sheetNew.getLastRow() - 1, Math.max(10, sheetNew.getLastColumn())).getValues();
    }
    
    let custoTotal = 0;
    const itemResults = [];
    
    for (let j = 0; j < (itens || []).length; j++) {
      let item = itens[j];
      let nomeArtigo = (item.nome || "").toString().trim().toLowerCase().replace(/\s+/g, '');
      let qtdReq = parseFloat(item.quantidade) || 0;
      let unitCustoEncontrado = 0; 
      
      for (let i = dataNew.length - 1; i >= 0; i--) {
         let r = dataNew[i];
         let tipo = (r[1] || "").toString().toLowerCase().trim();
         let artRaw = (r[4] || "").toString().toLowerCase().trim();
         let art = artRaw.replace(/\s+/g, '');
         if (tipo === "entrada" || tipo === "compras") {
           if (art === nomeArtigo || art.includes(nomeArtigo) || nomeArtigo.includes(art)) {
             unitCustoEncontrado = parseFloat(r[6]) || 0;
             if (unitCustoEncontrado <= 0) unitCustoEncontrado = parseFloat(r[7]) || 0;
             break;
           }
         }
      }
      
      let custoPorLinha = unitCustoEncontrado * qtdReq;
      custoTotal += custoPorLinha;
      itemResults.push({
        nome: item.nome,
        quantidade: qtdReq,
        unidade: item.unidade,
        custo_unitario: unitCustoEncontrado,
      });
    }
    
    let precoSugerido = 0;
    
    if (typeof getMarginSettings === 'function') {
       try {
           const marginData = getMarginSettings(impersonateEmail) || {};
           const margemDesejada = parseFloat(marginData.margemDesejada) || 30;
           const ircEstimado = parseFloat(marginData.ircEstimado) || 20;
           const autoOpExRate = parseFloat(marginData.autoOpExRate) || 0;
           const totalTaxas = margemDesejada + ircEstimado + autoOpExRate;
           const metodoCalculo = marginData.metodoCalculo || "markup";

           if (metodoCalculo === "margem_real") {
               precoSugerido = (totalTaxas < 100) ? (custoTotal / (1 - (totalTaxas / 100))) : custoTotal;
           } else {
               precoSugerido = custoTotal * (1 + (totalTaxas / 100));
           }
       } catch(e) {
           precoSugerido = custoTotal * 1.5;
       }
    } else {
        precoSugerido = custoTotal * 1.5;
    }
    
    return {
      success: true,
      custoTotal: custoTotal,
      precoSugerido: precoSugerido,
      itensEvaluated: itemResults
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function saveFichaTecnica(payload) {
  try {
    const email = payload.impersonateEmail || null;
    const init = _initFichasTecnicasTabs(email);
    if (!init.success) return init;
    
    const dbSheet = init.dbSheet;
    const itensSheet = init.itensSheet;
    
    let idFt = payload.id;
    let isEdit = !!idFt;
    if (!idFt) idFt = "FT-" + Utilities.getUuid().substring(0,8).toUpperCase();
    
    if (isEdit) {
       let dbData = dbSheet.getDataRange().getValues();
       for (let i = dbData.length - 1; i >= 1; i--) {if (dbData[i][0] === idFt) {dbSheet.deleteRow(i + 1); break;}}
       let itData = itensSheet.getDataRange().getValues();
       for (let i = itData.length - 1; i >= 1; i--) {if (itData[i][0] === idFt) {itensSheet.deleteRow(i + 1);}}
    }
    
    const userEmail = Session.getActiveUser().getEmail();
    const rowHeader = [
       idFt,
       payload.nome || "Nova Ficha",
       payload.categoria || "Sem Categoria",
       payload.preparacao || "",
       payload.custoTotal || 0,
       payload.precoSugerido || 0,
       payload.logoBase64 || "",
       new Date(),
       userEmail
    ];
    dbSheet.appendRow(rowHeader);
    
    const rowsItens = [];
    let items = payload.itens || [];
    for(let i=0; i<items.length; i++){
       let it = items[i];
       rowsItens.push([
           idFt,
           it.nome,
           it.quantidade,
           it.unidade,
           it.custo_unitario || 0
       ]);
    }
    if(rowsItens.length > 0) {
       itensSheet.getRange(itensSheet.getLastRow() + 1, 1, rowsItens.length, 5).setValues(rowsItens);
    }
    
    if (payload.precoFinal !== undefined) {
        try {
            const valPv = parseFloat(payload.precoFinal) || 0;
            const sheetNew = init.ss.getSheetByName(SHEET_TAB_NAME);
            if (sheetNew && sheetNew.getLastRow() >= 2) {
                const dataNew = sheetNew.getRange(2, 5, sheetNew.getLastRow() - 1, 1).getValues();
                const fArt = (payload.nome || "").toString().trim().toLowerCase();
                if (fArt) {
                    for (let i = 0; i < dataNew.length; i++) {
                        let art = (dataNew[i][0] || "").toString().trim().toLowerCase();
                        if (art === fArt) {
                            sheetNew.getRange(i + 2, 26).setValue(valPv);
                        }
                    }
                }
            }
        } catch(ex) {
            console.log("Error updating prices in New tab: " + ex);
        }
    }
    
    return { success: true, data: { id: idFt, nome: payload.nome || "Nova Ficha" } };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getUniqueArticlesForDatalist(impersonateEmail) {
  try {
    const ctx = getClientContext(impersonateEmail);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    if (!ss) return { success: false, data: [] };
    const sheetNew = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheetNew || sheetNew.getLastRow() < 2) return { success: true, data: [] };
    
    const data = sheetNew.getRange(2, 1, sheetNew.getLastRow() - 1, Math.max(5, sheetNew.getLastColumn())).getValues();
    const articles = new Set();
    for (let i = 0; i < data.length; i++) {
      let tipo = (data[i][1] || "").toString().toLowerCase().trim();
      let art = (data[i][4] || "").toString().trim();
      if (art && (tipo === "entrada" || tipo === "compras")) {
        articles.add(art);
      }
    }
    return { success: true, data: Array.from(articles).sort() };
  } catch(e) {
    return { success: false, error: e.toString(), data: [] };
  }
}

function getFichasTecnicas(impersonateEmail) {
  console.log("[DEBUG] getFichasTecnicas iniciada para:", impersonateEmail);
  if (!impersonateEmail) {
    return JSON.stringify({ success: false, error: "Email de impersonação inválido ou não disponível.", data: [] });
  }
  try {
     const init = _initFichasTecnicasTabs(impersonateEmail);
     if (!init || !init.success) {
        console.warn("[DEBUG] getFichasTecnicas -> Falha no _initFichasTecnicasTabs:", init ? init.error : "null");
        return JSON.stringify({success:false, error: init ? init.error : "Erro na inicialização das tabelas de FT.", data: []});
     }
     const dbSheet = init.dbSheet;
     if(!dbSheet) {
        console.warn("[DEBUG] getFichasTecnicas -> dbSheet é null");
        return JSON.stringify({success:false, error:"Tabela Fichas_Tecnicas_DB inacessível.", data: []});
     }
     const lastRow = dbSheet.getLastRow();
     console.log("[DEBUG] getFichasTecnicas -> lastRow na DB:", lastRow);
     if(lastRow < 2) return JSON.stringify({ success: true, data: [] });
     
     const data = dbSheet.getRange(2, 1, lastRow - 1, 9).getValues();
     let result = [];
     for(let i=0; i<data.length; i++) {
        // Garantir que todos os campos são strings ou números simples para evitar erros de serialização
        result.push({
           id: String(data[i][0] || ""),
           nome: String(data[i][1] || ""),
           categoria: String(data[i][2] || ""),
           preparacao: String(data[i][3] || ""),
           custoTotal: parseFloat(data[i][4]) || 0,
           precoSugerido: parseFloat(data[i][5]) || 0,
           logo: data[i][6] ? String(data[i][6]) : "",
           timestamp: data[i][7] ? (typeof data[i][7].toISOString === 'function' ? data[i][7].toISOString() : Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'")) : ""
        });
     }
     console.log("[DEBUG] getFichasTecnicas -> Sucesso. Total processado:", result.length);
     return JSON.stringify({ success: true, data: result.reverse() });
  } catch(e) {
     const errorMsg = e.message || e.toString();
     console.error("[CRITICAL] Erro em getFichasTecnicas:", errorMsg);
     const userEmail = impersonateEmail || (typeof Session !== 'undefined' ? Session.getActiveUser().getEmail() : "system");
     try {
        logSystemError("MOD_Fichas_Tecnicas", e, userEmail);
     } catch(errLog) {
        console.error("Falha ao logar erro no sistema:", errLog);
     }
     return JSON.stringify({ success: false, error: errorMsg || "Exceção desconhecida em getFichasTecnicas.", data: [] });
  }
}

function getFichaTecnica(id, impersonateEmail) {
  try {
     const init = _initFichasTecnicasTabs(impersonateEmail);
     if (!init.success) return init;
     const dbSheet = init.dbSheet;
     const itensSheet = init.itensSheet;
     
     if(dbSheet.getLastRow() < 2) return {success:false, error:"Nenhuma ficha encontrada."};
     const dbData = dbSheet.getDataRange().getValues();
     let header = null;
     for(let i=1; i<dbData.length; i++){
        if(dbData[i][0] === id) {
           header = {
             id: dbData[i][0],
             nome: dbData[i][1],
             categoria: dbData[i][2],
             preparacao: dbData[i][3],
             custoTotal: parseFloat(dbData[i][4]) || 0,
             precoSugerido: parseFloat(dbData[i][5]) || 0,
             logo: dbData[i][6]
           };
           break;
        }
     }
     
     if(!header) return {success:false, error:"Ficha não encontrada."};
     
     let itens = [];
     if(itensSheet.getLastRow() >= 2) {
        let itemsData = itensSheet.getDataRange().getValues();
        for(let i=1; i<itemsData.length; i++){
           if(itemsData[i][0] === id) {
              itens.push({
                 nome: itemsData[i][1],
                 quantidade: parseFloat(itemsData[i][2]) || 0,
                 unidade: itemsData[i][3],
                 custo_unitario: parseFloat(itemsData[i][4]) || 0
              });
           }
        }
     }
     header.itens = itens;
     return { success: true, data: header };
  } catch(e) {
     return { success: false, error: e.toString() };
  }
}

function deleteFichaTecnica(id, impersonateEmail) {
    try {
        const init = _initFichasTecnicasTabs(impersonateEmail);
        if (!init.success) return init;
        
        let dbData = init.dbSheet.getDataRange().getValues();
        for (let i = dbData.length - 1; i >= 1; i--) {if (dbData[i][0] === id) {init.dbSheet.deleteRow(i + 1); break;}}
        let itData = init.itensSheet.getDataRange().getValues();
        for (let i = itData.length - 1; i >= 1; i--) {if (itData[i][0] === id) {init.itensSheet.deleteRow(i + 1);}}
        return { success: true };
    } catch(e) {
        return { success: false, error: e.toString() };
    }
}

function processComposeStock(venda_items, impersonateEmail) {
  try {
     const init = _initFichasTecnicasTabs(impersonateEmail);
     if(!init.success) return { success: false, hasComposed: false }; 
     if(init.dbSheet.getLastRow() < 2 || init.itensSheet.getLastRow() < 2) return { success: true, hasComposed: false };
     
     const dbData = init.dbSheet.getDataRange().getValues();
     const itemData = init.itensSheet.getDataRange().getValues();
     
     let generated_saidas = [];
     for(var v=0; v<venda_items.length; v++) {
         let soldArtigo = (venda_items[v].artigo || "").toString().trim().toLowerCase();
         let soldQtd = parseFloat(venda_items[v].quantidade) || 0;
         
         let ft_id = null;
         for(let i=1; i<dbData.length; i++){
             if((dbData[i][1] || "").toString().trim().toLowerCase() === soldArtigo) {
                 ft_id = dbData[i][0];
                 break;
             }
         }
         
         if(ft_id) {
             for(let j=1; j<itemData.length; j++){
                 if(itemData[j][0] === ft_id) {
                     generated_saidas.push({
                         artigo: itemData[j][1],
                         quantidade: (parseFloat(itemData[j][2]) || 0) * soldQtd,
                         preco_custo: parseFloat(itemData[j][4]) || 0,
                         fornecedor: "BAIXA AUTOMÁTICA (FT)",
                         taxa_iva: 0,
                         valor_iva: 0,
                         preco_venda: 0
                     });
                 }
             }
         }
     }
     if (generated_saidas.length > 0) {
         return { success:true, hasComposed: true, saidas: generated_saidas };
     }
     return { success:true, hasComposed: false };
  } catch(e) {
     return { success: false, error: e.toString() };
  }
}

function getFTCostMap(impersonateEmail) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "FT_COSTS_" + (impersonateEmail || "MASTER");
  const cached = cache.get(cacheKey);
  if (cached) {
     try {
       return JSON.parse(cached);
     } catch(e) {}
  }
  
  try {
     const init = _initFichasTecnicasTabs(impersonateEmail);
     if (!init || !init.success) return {};

     const dbSheet = init.dbSheet;
     if(!dbSheet || dbSheet.getLastRow() < 2) return {};

     const data = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 9).getValues();
     let costMap = {};
     for(let i=0; i<data.length; i++) {
        let nome = String(data[i][1] || "").trim();
        let custo = parseFloat(data[i][4]) || 0;
        if (nome) {
           costMap[nome.toLowerCase()] = custo;
        }
     }
     cache.put(cacheKey, JSON.stringify(costMap), 3600);
     return costMap;
  } catch(e) {
     return {};
  }
}

function invalidateFTCostCache(impersonateEmail) {
   const cache = CacheService.getScriptCache();
   const cacheKey = "FT_COSTS_" + (impersonateEmail || "MASTER");
   cache.remove(cacheKey);
}

