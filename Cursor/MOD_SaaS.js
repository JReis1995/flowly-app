// Ficheiro: MOD_SaaS.js
/// ==========================================
// 🏢 MÓDULO DE GESTÃO SAAS (SUPER ADMIN)
// ==========================================

function getSuperAdminDashboardData(callerEmail) {
  const activeUser = Session.getActiveUser().getEmail();
  const callerNorm = (callerEmail || "").toString().trim().toLowerCase();
  const effectiveSA = (activeUser === SUPER_ADMIN_EMAIL) || (activeUser === "" && callerNorm === SUPER_ADMIN_EMAIL.toLowerCase()) || (callerNorm && isFlowlyTeamMember(callerEmail));
  if (!effectiveSA) return { success: false };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'SA_DASHBOARD';
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed !== null && typeof parsed === 'object') return parsed;
      cache.remove(cacheKey);
    } catch (e) { cache.remove(cacheKey); }
  }

  ensureMasterDBColumns();
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const data = ss.getSheets()[0].getDataRange().getValues();
  const plans = getSaaSPlansList();
  const clients = [];
  let mrrTotal = 0;

  for (let i = 1; i < data.length; i++) {
    let modules = { rh: true, cc: true, logistica: true, dashboard: true };
    try { if (data[i][7]) modules = JSON.parse(data[i][7]); } catch (e) { }
    const planName = data[i][5] || "";
    const planObj = plans.find(p => p.name === planName) || { price: 0 };
    const avenca = parseFloat(planObj.price) || 0;
    const descontoRaw = parseFloat(data[i][MASTER_COL_DESCONTO - 1]) || 0;
    const mensalidadesOferta = parseInt(data[i][MASTER_COL_MENSALIDADES_OFERTA - 1]) || 0;
    const descontoTipo = String(data[i][MASTER_COL_DESCONTO_TIPO - 1] || "Permanente").trim();
    const descontoExpiraRaw = data[i][MASTER_COL_DESCONTO_EXPIRA - 1];
    const descontoExpira = parseDate(descontoExpiraRaw);
    let descontoEfetivo = descontoRaw;
    if (descontoTipo === "Temporário" && descontoExpira && !isNaN(descontoExpira.getTime())) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const expDate = new Date(descontoExpira);
      expDate.setHours(0, 0, 0, 0);
      if (today > expDate) descontoEfetivo = 0;
    }
    const valorFinal = Math.round(avenca * (1 - descontoEfetivo / 100) * 100) / 100;
    const contribuiMRR = (data[i][4] === "Ativo") && (mensalidadesOferta <= 0) ? valorFinal : 0;
    mrrTotal += contribuiMRR;

    clients.push({
      name: data[i][0], email: data[i][1], sheetId: data[i][2], status: data[i][4], plan: planName,
      planPrice: avenca, desconto: descontoRaw, descontoTipo: descontoTipo, descontoExpira: descontoExpira,
      mensalidadesOferta: mensalidadesOferta, valorFinal: valorFinal, modules: modules,
      saasPlan: data[i][MASTER_COL_SAAS_PLAN - 1] || 'FREE', aiCredits: data[i][MASTER_COL_AI_CREDITS - 1] || 0
    });
  }
  const result = { success: true, isMaster: true, clients: clients, plans: plans, mrrTotal: Math.round(mrrTotal * 100) / 100 };
  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { }
  return result;
}

function invalidateSACache() {
  try { CacheService.getScriptCache().remove('SA_DASHBOARD'); } catch (e) { }
}

function logFaturarSelecionados(emails) {
  Logger.log('Selecionados: ' + (Array.isArray(emails) ? emails.join(', ') : emails));
  return { success: true };
}

function insertMonthlySaaSInvoiceBatch(emails) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  const list = Array.isArray(emails) ? emails : [];
  const results = [];
  for (let i = 0; i < list.length; i++) {
    const email = String(list[i] || "").trim();
    if (!email || !email.includes("@")) continue;
    const r = insertMonthlySaaSInvoice(email);
    results.push({ email: email, success: r.success, valorFatura: r.valorFatura || 0, mensalidadesOferta: r.mensalidadesOferta, error: r.error });
  }
  invalidateSACache();
  return { success: true, processed: results.length, results: results };
}

function applyBulkSettings(emails, settings) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!settings || typeof settings !== "object") return { success: false, error: "Settings inválidos" };
  const list = Array.isArray(emails) ? emails : [];
  const results = [];
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    const data = ss.getDataRange().getValues();
    for (let i = 0; i < list.length; i++) {
      const email = String(list[i] || "").trim();
      if (!email || !email.includes("@")) continue;
      const row = data.findIndex(r => String(r[1]).toLowerCase().trim() === email.toLowerCase());
      if (row < 0) { results.push({ email: email, success: false, error: "Cliente não encontrado" }); continue; }
      const rowNum = row + 1;
      if (typeof settings.desconto === "number" || (settings.desconto !== "" && settings.desconto != null)) {
        ss.getRange(rowNum, MASTER_COL_DESCONTO).setValue(parseFloat(settings.desconto) || 0);
      }
      if (typeof settings.mensalidadesOferta === "number" || (settings.mensalidadesOferta !== "" && settings.mensalidadesOferta != null)) {
        ss.getRange(rowNum, MASTER_COL_MENSALIDADES_OFERTA).setValue(Math.max(0, parseInt(settings.mensalidadesOferta) || 0));
      }
      if (settings.descontoTipo != null && settings.descontoTipo !== undefined) {
        ss.getRange(rowNum, MASTER_COL_DESCONTO_TIPO).setValue(String(settings.descontoTipo) === "Temporário" ? "Temporário" : "Permanente");
      }
      if (settings.descontoExpira != null && settings.descontoExpira !== undefined) {
        ss.getRange(rowNum, MASTER_COL_DESCONTO_EXPIRA).setValue(settings.descontoExpira ? toDBDateFromInput(settings.descontoExpira) : "");
      }
      results.push({ email: email, success: true });
    }
    invalidateSACache();
    return { success: true, processed: results.length, results: results };
  } catch (e) { return { success: false, error: e.toString(), results: results }; }
}

function getClientValorFatura(clientEmail) {
  const cfg = getClientSaaSConfigFromMaster(clientEmail);
  if (!cfg) return { valorFatura: 0, valorFinal: 0, mensalidadesOferta: 0, custoProximoMes: 0 };
  const valorFinal = cfg.valorFinal;
  const mensalidadesOferta = cfg.mensalidadesOferta || 0;
  const custoProximoMes = mensalidadesOferta > 0 ? 0 : valorFinal;
  return { valorFatura: custoProximoMes, valorFinal, mensalidadesOferta, custoProximoMes };
}

function getClientSaaSConfigFromMaster(clientEmail) {
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const data = ss.getSheets()[0].getDataRange().getValues();
    const row = data.find(r => String(r[1]).toLowerCase().trim() === String(clientEmail).toLowerCase().trim());
    if (!row) return null;
    const plans = getSaaSPlansList();
    const plan = plans.find(p => p.name === (row[5] || "")) || { price: 0 };
    const avenca = parseFloat(plan.price) || 0;
    let desconto = parseFloat(row[MASTER_COL_DESCONTO - 1]) || 0;
    const mensalidadesOferta = parseInt(row[MASTER_COL_MENSALIDADES_OFERTA - 1]) || 0;
    const descontoTipo = String(row[MASTER_COL_DESCONTO_TIPO - 1] || "Permanente").trim();
    const descontoExpiraRaw = row[MASTER_COL_DESCONTO_EXPIRA - 1];
    const descontoExpira = parseDate(descontoExpiraRaw);
    if (descontoTipo === "Temporário" && descontoExpira && !isNaN(descontoExpira.getTime())) {
      const today = new Date(); today.setHours(0, 0, 0, 0);
      const expDate = new Date(descontoExpira); expDate.setHours(0, 0, 0, 0);
      if (today > expDate) desconto = 0;
    }
    const valorFinal = Math.round(avenca * (1 - desconto / 100) * 100) / 100;
    return { avenca, desconto, mensalidadesOferta, valorFinal, descontoTipo, descontoExpira, plan: row[5], name: row[0] };
  } catch (e) { return null; }
}

function applyFreeMonth(clientID) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => String(r[1]).toLowerCase().trim() === String(clientID).toLowerCase().trim());
    if (idx < 0) return { success: false, error: "Cliente não encontrado" };
    const rowNum = idx + 1;
    const currentVal = parseInt(sheet.getRange(rowNum, MASTER_COL_MENSALIDADES_OFERTA).getValue()) || 0;
    const newVal = Math.max(0, currentVal - 1);
    sheet.getRange(rowNum, MASTER_COL_MENSALIDADES_OFERTA).setValue(newVal);
    invalidateSACache();
    return { success: true, mensalidadesRestantes: newVal };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function createMonthlyInvoicesForAllClients() {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  invalidateSACache();
  const res = getSuperAdminDashboardData(null);
  if (!res.success || !res.clients) return { success: false, error: "Sem dados" };
  const ativos = res.clients.filter(c => c.status === "Ativo");
  const results = [];
  ativos.forEach(c => {
    const r = insertMonthlySaaSInvoice(c.email);
    results.push({ email: c.email, name: c.name, success: r.success, valorFatura: r.valorFatura || 0, mensalidadesOferta: r.mensalidadesOferta, error: r.error });
  });
  return { success: true, processed: results.length, results };
}

function insertMonthlySaaSInvoice(clientEmail) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  try {
    const cfg = getClientSaaSConfigFromMaster(clientEmail);
    if (!cfg) return { success: false, error: "Cliente não encontrado" };
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const data = ss.getSheets()[0].getDataRange().getValues();
    const row = data.find(r => String(r[1]).toLowerCase().trim() === String(clientEmail).toLowerCase().trim());
    if (!row) return { success: false, error: "Cliente não encontrado" };
    const sheetId = row[2];
    const clientName = row[0];
    const planName = row[5] || "Plano";
    const { valorFinal, mensalidadesOferta, custoProximoMes } = getClientValorFatura(clientEmail);
    const today = new Date();
    const ssClient = SpreadsheetApp.openById(sheetId);
    let logSheet = ssClient.getSheetByName(SHEET_TAB_NAME);
    if (!logSheet) logSheet = ssClient.insertSheet(SHEET_TAB_NAME);
    const valorFatura = custoProximoMes;
    const valorSugerido = calcularPrecoSugerido(valorFatura);
    logSheet.appendRow([
      today, "Entrada", "Sistema", "Flowly SaaS", `Mensalidade ${planName} ${clientName}`, "1", valorFatura, "", "23%", valorFatura * 0.23, "", "VALIDADO",
      mensalidadesOferta > 0 ? "Mês de Oferta (0€)" : "Débito Automático", "", today, "Sistema", "Aberto", "", 0, "Não", "", "", valorSugerido
    ]);
    if (mensalidadesOferta > 0) applyFreeMonth(clientEmail);
    return { success: true, valorFatura, mensalidadesOferta };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function createNewClient(clientName, clientEmail, planType, desconto, descontoTipo, descontoExpira, businessVertical, nifEmpresa) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  desconto = (typeof desconto === "number" ? desconto : parseFloat(desconto)) || 0;
  descontoTipo = (descontoTipo === "Temporário") ? "Temporário" : "Permanente";
  descontoExpira = (descontoExpira && String(descontoExpira).trim()) || "";
  businessVertical = businessVertical || "Standard";
  nifEmpresa = nifEmpresa || "";

  try {
    const plans = getSaaSPlansList();
    const plan = plans.find(p => p.name === planType) || { price: 0, modules: { rh: true, cc: true, logistica: true, dashboard: true } };
    const planPrice = plan.price || 0;

    const rootFolder = DriveApp.getFolderById(FLOWLY_ROOT_FOLDER_ID);
    const clientFolder = rootFolder.createFolder(`[${clientName}] - Flowly Data`);
    const docsFolder = clientFolder.createFolder("Faturas_Digitalizadas");

    const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    const newSheet = templateFile.makeCopy(`Flowly_DB_${clientName}`, clientFolder);
    const newSheetId = newSheet.getId();

    const ssClient = SpreadsheetApp.openById(newSheetId);
    let logSheet = ssClient.getSheetByName(SHEET_TAB_NAME);
    if (!logSheet) logSheet = ssClient.insertSheet(SHEET_TAB_NAME);

    const today = new Date();
    const valorAtivacao = Math.round(planPrice * (1 - desconto / 100) * 100) / 100;
    const valorSugerido = calcularPrecoSugerido(valorAtivacao);
    logSheet.appendRow([
      today, "Entrada", "Sistema", "Flowly SaaS", `Ativação Plano ${planType}`, "1", valorAtivacao, "", "23%", (valorAtivacao * 0.23), "", "VALIDADO", "Débito Automático", "", today, "Sistema", "Aberto", "", 0, "Não", "", "", valorSugerido
    ]);

    ensureMasterDBColumns();
    const ssMaster = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    ssMaster.appendRow([clientName, clientEmail, newSheetId, docsFolder.getId(), "Ativo", planType, new Date(), JSON.stringify(plan.modules), desconto, 0]);
    ssMaster.getRange(ssMaster.getLastRow(), MASTER_COL_DESCONTO_TIPO).setValue(descontoTipo);
    ssMaster.getRange(ssMaster.getLastRow(), MASTER_COL_DESCONTO_EXPIRA).setValue(descontoExpira ? toDBDateFromInput(descontoExpira) : "");
    ssMaster.getRange(ssMaster.getLastRow(), MASTER_COL_BUSINESS_VERTICAL).setValue(businessVertical);
    ssMaster.getRange(ssMaster.getLastRow(), MASTER_COL_NIF_EMPRESA).setValue(nifEmpresa);

    invalidateSACache();

    // 3. TENTATIVA DE PERMISSÕES DRIVE (seguro)
    safeAddEditor(clientFolder, clientEmail);
    safeAddEditor(docsFolder, clientEmail);
    safeAddEditor(newSheet, clientEmail);

    // 4. ENVIO DE CONVITES
    const token = Utilities.base64EncodeWebSafe(clientEmail + "_" + Date.now());
    sendInviteEmail(clientEmail, clientName, "Gestor", token);
    try {
      var folderHtml = `<p>Olá,</p><p>A sua área de armazenamento de documentos Flowly 360 foi criada e está disponível.</p><p><a href="${clientFolder.getUrl()}" style="background:#10B981;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;">Aceder à Minha Área</a></p><p style="margin-top:20px;font-size:12px;color:#64748B;">Atenciosamente,<br><strong>Equipa Flowly 360</strong><br><a href="mailto:geral@flowly.pt" style="color:#10B981;">geral@flowly.pt</a> &bull; www.flowly.pt<br><em>Onde o fluxo encontra a precisão.</em></p>`;
      // [NOTA TÉCNICA] O utilizador precisa de configurar o Alias (Send As) no Gmail da conta que executa o script para que o parâmetro 'from' funcione e minimize o impacto do email pessoal.
      var folderOptions = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: folderHtml };
      GmailApp.sendEmail(clientEmail, "A sua área Flowly 360 está pronta — " + clientName, "", folderOptions);
    } catch (eNotif) { }

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveClientSettings(email, newData) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false };
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    const data = ss.getDataRange().getValues();
    const row = data.findIndex(r => String(r[1]).toLowerCase().trim() === String(email).toLowerCase().trim());

    if (row > -1) {
      ss.getRange(row + 1, 1).setValue(newData.name);
      if (newData.status) ss.getRange(row + 1, 5).setValue(newData.status);
      if (newData.plan != null && newData.plan !== undefined) ss.getRange(row + 1, 6).setValue(newData.plan);
      ss.getRange(row + 1, 8).setValue(JSON.stringify(newData.modules || {}));
      if (typeof newData.desconto === "number" || (newData.desconto !== "" && newData.desconto != null)) {
        ss.getRange(row + 1, MASTER_COL_DESCONTO).setValue(parseFloat(newData.desconto) || 0);
      }
      if (typeof newData.mensalidadesOferta === "number" || (newData.mensalidadesOferta !== "" && newData.mensalidadesOferta != null)) {
        ss.getRange(row + 1, MASTER_COL_MENSALIDADES_OFERTA).setValue(Math.max(0, parseInt(newData.mensalidadesOferta) || 0));
      }
      if (newData.descontoTipo != null && newData.descontoTipo !== undefined) {
        ss.getRange(row + 1, MASTER_COL_DESCONTO_TIPO).setValue(String(newData.descontoTipo) === "Temporário" ? "Temporário" : "Permanente");
      }
      if (newData.descontoExpira != null && newData.descontoExpira !== undefined) {
        ss.getRange(row + 1, MASTER_COL_DESCONTO_EXPIRA).setValue(newData.descontoExpira ? toDBDateFromInput(newData.descontoExpira) : "");
      }
      if (newData.saasPlan != null && newData.saasPlan !== undefined) {
        ss.getRange(row + 1, MASTER_COL_SAAS_PLAN).setValue(String(newData.saasPlan).trim() || 'FREE');
      }
      if (newData.aiCredits != null && newData.aiCredits !== undefined && newData.aiCredits !== '') {
        ss.getRange(row + 1, MASTER_COL_AI_CREDITS).setValue(Math.max(0, parseInt(newData.aiCredits, 10) || 0));
      }
      invalidateSACache();
      return { success: true };
    }
    return { success: false, error: "Cliente não encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function massUpdateCredits(emails, newCreditLimit) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    const data = ss.getDataRange().getValues();
    const emailsSet = new Set((Array.isArray(emails) ? emails : []).map(function (e) { return String(e || "").trim().toLowerCase(); }).filter(Boolean));
    const props = PropertiesService.getScriptProperties();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][MASTER_COL_EMAIL - 1] || "").trim().toLowerCase();
      if (!rowEmail) continue;
      if (emailsSet.has(rowEmail)) {
        ss.getRange(i + 1, MASTER_COL_AI_CREDITS).setValue(newCreditLimit);
        props.setProperty("FLOWLY_AI_CREDITS_" + rowEmail, String(newCreditLimit));
      }
    }
    invalidateSACache();
    return { success: true, message: emails.length + " clientes atualizados com " + newCreditLimit + " créditos." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function massUpdatePlans(emails, newPlan) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    const data = ss.getDataRange().getValues();
    const emailsSet = new Set((Array.isArray(emails) ? emails : []).map(function (e) { return String(e || "").trim().toLowerCase(); }).filter(Boolean));
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][MASTER_COL_EMAIL - 1] || "").trim().toLowerCase();
      if (!rowEmail || !emailsSet.has(rowEmail)) continue;
      ss.getRange(i + 1, MASTER_COL_SAAS_PLAN).setValue(newPlan);
    }
    invalidateSACache();
    return { success: true, message: emails.length + " clientes movidos para o plano " + newPlan + "." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function massUpdateStatus(emails, newStatus) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  var statusVal = (newStatus || "").toString().trim();
  if (statusVal !== "Ativo" && statusVal !== "Suspenso") return { success: false, error: "Estado inválido. Use Ativo ou Suspenso." };
  try {
    var ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    var data = ss.getDataRange().getValues();
    var emailsSet = {};
    (Array.isArray(emails) ? emails : []).forEach(function (e) { var k = String(e || "").trim().toLowerCase(); if (k) emailsSet[k] = true; });
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][1] || "").trim().toLowerCase();
      if (rowEmail && emailsSet[rowEmail]) { ss.getRange(i + 1, 5).setValue(statusVal); count++; }
    }
    invalidateSACache();
    return { success: true, message: count + " cliente(s) atualizado(s) para " + statusVal + "." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function massDeleteClients(emails) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  try {
    var ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    var data = ss.getDataRange().getValues();
    var emailsSet = {};
    (Array.isArray(emails) ? emails : []).forEach(function (e) { var k = String(e || "").trim().toLowerCase(); if (k) emailsSet[k] = true; });
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][1] || "").trim().toLowerCase();
      if (rowEmail && emailsSet[rowEmail]) rowsToDelete.push(i + 1);
    }
    for (var j = rowsToDelete.length - 1; j >= 0; j--) ss.deleteRow(rowsToDelete[j]);
    invalidateSACache();
    return { success: true, message: rowsToDelete.length + " cliente(s) eliminado(s)." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getSaaSPlansList() {
  const p = PropertiesService.getScriptProperties().getProperty('SAAS_PLANS_DB');
  if (p) return JSON.parse(p);
  return [{ name: "Basic", price: 29 }, { name: "Pro", price: 59 }];
}

function saveSaaSPlansList(plans) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false };
  PropertiesService.getScriptProperties().setProperty('SAAS_PLANS_DB', JSON.stringify(plans));
  invalidateSACache();
  return { success: true };
}

function getPlanosData() {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, plans: [], error: "Sem permissão" };
  var plans = getSaaSPlansList();
  return { success: true, plans: plans };
}

function deletePlanData(planId) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false };
  try {
    let plans = getSaaSPlansList();
    plans = plans.filter(p => String(p.id) !== String(planId));
    PropertiesService.getScriptProperties().setProperty('SAAS_PLANS_DB', JSON.stringify(plans));
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function savePlanData(p) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false };
  try {
    let plans = getSaaSPlansList();
    if (p.id === 'NOVO' || !p.id) {
      p.id = 'PLN-' + Date.now();
      plans.push(p);
    } else {
      const idx = plans.findIndex(x => String(x.id) === String(p.id));
      if (idx > -1) plans[idx] = p; else plans.push(p);
    }
    PropertiesService.getScriptProperties().setProperty('SAAS_PLANS_DB', JSON.stringify(plans));
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function syncClientsWithMasterPlan(planName, newSettings) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!planName || !newSettings || typeof newSettings !== "object") return { success: false, error: "Parâmetros inválidos" };
  try {
    ensureMasterDBColumns();
    const ss = SpreadsheetApp.openById(MASTER_DB_ID).getSheets()[0];
    const data = ss.getDataRange().getValues();
    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const rowPlan = String(data[i][5] || "").trim();
      if (rowPlan !== String(planName).trim()) continue;
      const rowNum = i + 1;
      if (typeof newSettings.desconto === "number" || (newSettings.desconto !== "" && newSettings.desconto != null)) ss.getRange(rowNum, MASTER_COL_DESCONTO).setValue(parseFloat(newSettings.desconto) || 0);
      if (typeof newSettings.mensalidadesOferta === "number" || (newSettings.mensalidadesOferta !== "" && newSettings.mensalidadesOferta != null)) ss.getRange(rowNum, MASTER_COL_MENSALIDADES_OFERTA).setValue(Math.max(0, parseInt(newSettings.mensalidadesOferta) || 0));
      if (newSettings.descontoTipo != null && newSettings.descontoTipo !== undefined) ss.getRange(rowNum, MASTER_COL_DESCONTO_TIPO).setValue(String(newSettings.descontoTipo) === "Temporário" ? "Temporário" : "Permanente");
      if (newSettings.descontoExpira != null && newSettings.descontoExpira !== undefined) ss.getRange(rowNum, MASTER_COL_DESCONTO_EXPIRA).setValue(newSettings.descontoExpira ? toDBDateFromInput(newSettings.descontoExpira) : "");
      updated++;
    }
    invalidateSACache();
    return { success: true, updated: updated };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getPacotesIASheet() {
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  let sh = ss.getSheetByName('Pacotes_IA');
  if (!sh) {
    sh = ss.insertSheet('Pacotes_IA');
    sh.getRange(1, 1, 1, 8).setValues([['Data Criação', 'ID', 'Nome', 'Creditos', 'Preco', 'Link_Pagamento', 'Status', 'Ultima Modificação']]);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sh;
}

function getAIPacksConfig() {
  try {
    const sh = getPacotesIASheet();
    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return [];
    const packs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      packs.push({ id: row[1], name: row[2] || '', credits: row[3] != null ? row[3] : 0, price: row[4] != null ? row[4] : 0, link: row[5] || '', status: (row[6] || '').toString().trim() });
    }
    return packs;
  } catch (e) { return []; }
}

function getStorePackages() {
  try {
    const sh = getPacotesIASheet();
    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    const result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = (row[6] || '').toString().trim().toLowerCase();
      if (status !== 'ativo') continue;

      // CORREÇÃO: Usar as chaves em inglês (name, credits, price) para o frontend conseguir ler!
      result.push({
        id: row[1],
        name: row[2] || '',
        credits: row[3] != null ? row[3] : 0,
        price: row[4] != null ? row[4] : 0,
        link: row[5] || ''
      });
    }
    return result;
  } catch (e) { return []; }
}

function saveAIPackConfig(pack) {
  try {
    const sh = getPacotesIASheet();
    const data = sh.getDataRange().getValues();
    const now = new Date();
    const nowStr = now.toISOString ? now.toISOString().slice(0, 19).replace('T', ' ') : Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (!pack.id) {
      var newId = 'PACK_' + Date.now();
      sh.appendRow([nowStr, newId, pack.name || '', pack.credits != null ? pack.credits : 0, pack.price != null ? pack.price : 0, pack.link || '', (pack.status || 'Ativo').toString().trim() || 'Ativo', nowStr]);
    } else {
      var rowIdx = -1;
      for (var r = 1; r < data.length; r++) { if (data[r][1] === pack.id) { rowIdx = r + 1; break; } }
      if (rowIdx < 2) return { success: false, message: 'Pacote não encontrado.' };

      // CORREÇÃO: A fórmula do Google Sheets é getRange(linha_inicial, coluna_inicial, N_Linhas, N_Colunas)
      // Antes estava a tentar atualizar "rowIdx" linhas (ex: 2) com apenas 1 linha de dados.
      sh.getRange(rowIdx, 3, 1, 6).setValues([[
        pack.name || '',
        pack.credits != null ? pack.credits : 0,
        pack.price != null ? pack.price : 0,
        pack.link || '',
        (pack.status || 'Ativo').toString().trim() || 'Ativo',
        nowStr
      ]]);
    }
    return { success: true, message: 'Pacote IA guardado com sucesso!' };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteAIPackConfig(id) {
  try {
    const sh = getPacotesIASheet();
    const data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (data[r][1] === id) { sh.deleteRow(r + 1); return { success: true, message: 'Pacote apagado com sucesso.' }; }
    }
    return { success: false, message: 'Pacote não encontrado.' };
  } catch (e) { return { success: false, message: e.message }; }
}

function createStripeCheckout(packId) {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var secretKey = scriptProps.getProperty('STRIPE_SECRET_KEY');
    var webAppUrl = scriptProps.getProperty('WEBAPP_URL');
    if (!secretKey) throw new Error("Chave Stripe (STRIPE_SECRET_KEY) não configurada.");
    if (!webAppUrl) throw new Error("URL da App (WEBAPP_URL) não configurada.");

    var packs = getAIPacksConfig();
    var pack = packs.find(function (p) { return p.id === packId; });
    if (!pack) throw new Error("Pacote não encontrado.");

    var userEmail = (Session.getActiveUser() && Session.getActiveUser().getEmail()) ? Session.getActiveUser().getEmail() : '';
    if (!userEmail) userEmail = 'admin@flowly.pt';

    var price = parseFloat(pack.price);
    var finalPrice = Math.round(price * 100);
    var successUrl = webAppUrl + '?page=success&session_id={CHECKOUT_SESSION_ID}';

    var payload = {
      'mode': 'payment', 'success_url': successUrl, 'cancel_url': webAppUrl, 'line_items[0][price_data][currency]': 'eur',
      'line_items[0][price_data][unit_amount]': finalPrice.toString(), 'line_items[0][price_data][product_data][name]': 'Pack IA: ' + pack.name,
      'line_items[0][quantity]': '1', 'customer_email': userEmail, 'metadata[packId]': packId, 'metadata[credits]': pack.credits.toString(), 'metadata[userEmail]': userEmail
    };

    var options = { 'method': 'post', 'headers': { 'Authorization': 'Bearer ' + secretKey }, 'contentType': 'application/x-www-form-urlencoded', 'payload': payload, 'muteHttpExceptions': true };
    var response = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', options);
    var result = JSON.parse(response.getContentText());

    if (result.error) throw new Error(result.error.message);
    return { success: true, url: result.url };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function safeAddEditor(resource, email) {
  try {
    resource.addEditor(email);
  } catch (e) {
    console.error("Erro ao adicionar editor " + email + ": " + e);
    try {
      // Falha silenciosa para o cliente, mas notifica o Admin
      const AdminMail = "geral@flowly.pt"; // ou SUPER_ADMIN_EMAIL se disponível
      MailApp.sendEmail(AdminMail, "Aviso de Permissão Drive - Flowly", "Atenção: Não foi possível atribuir permissões automáticas no Google Drive para o email: " + email + ". O utilizador pode não ter uma conta Google. O cliente foi criado, mas verifique o acesso.");
    } catch(err) {}
  }
}