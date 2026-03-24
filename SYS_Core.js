// Ficheiro: SYS_Core.js
/// ==========================================
// 🚀 FLOWLY 360 - CORE INFRASTRUCTURE
// ==========================================

// --- CONFIGURAÇÃO GLOBAL ---
const SUPER_ADMIN_EMAIL = "josereis1995@gmail.com";
const DEFAULT_PLAN_CONFIG = { rh: true, dashRH: true, dashFinanceiro: true, dashStocks: true, dashFrota: true, caixaLivre: true, cc: true, logistica: true, admin: true, dashboard: true, ia: true, exportacao: true, crm: true, fichas_tecnicas: true };
const MASTER_DB_ID = "1N-RBi6Kbhnv2alCq3EZPIbAJgSFjYHsWUVr_bCvjdg0";
const FLOWLY_OFFICIAL_SHEET_ID = "1t0zLlGXf2zDvpQys7eGgZkDDZFN8h-A7Dorx3tZQUWw";
const SUPERADMIN_TEST_SHEET_ID = "1sfEwEk6dFotdPsm3owm4ImstHqQA5yPCyUeOzsNePD0"; // Sheet ID cedida para testes
const FLOWLY_ROOT_FOLDER_ID = "1VgT1fnnnEsOnHan5jPTOTsuc4TUXIDC1";
const TEMPLATE_SHEET_ID = "1EQVSSj6cgSubPxwGlNmG-gV-5I-BzEw1-sDHPgOKX4E";

// --- NOMES DAS ABAS (TABS) ---
const SHEET_TAB_NAME = "New";
const SHEET_STAFF_NAME = "Staff_DB";
const SHEET_STAFF_NAME_LEGACY = "Staff_DB";
const SHEET_USERS_NAME = "Users_DB";
const AI_HISTORY_TAB = "AI_History";
const FROTA_VEICULOS_TAB = "Frota_Veiculos";
const FROTA_CUSTOS_TAB = "Frota_Custos";
const LOGISTICA_VASILHAME_TAB = "Logistica_Vasilhame";
const CONFIG_VASILHAME_TAB = "Config_Vasilhame";
const CLIENTES_DB_TAB = "Clientes_DB";
const AUDIT_DB_TAB = "Audit_DB";
const FORNECEDORES_DB_TAB = "Fornecedores";
const CONFIG_MAPEAMENTO_TAB = "Config_Mapeamento";
const BUILDINGS_DB_TAB = "Buildings_DB";
const UNITS_DB_TAB = "Units_DB";

// --- CABEÇALHOS E LISTAS ---
const LISTA_VASILHAME = ["Palete Euro", "Palete Madeira", "Palete Plástico", "Epal", "Chep", "Caixa Plástico", "Caixa de Grade", "Grade de Cerveja", "Grade de Bebida", "Barril", "Keg", "Contentor IBC", "Big Bag", "Roll Container", "Grade de Leite", "Cesto Empilhável", "Garrafa Vidro", "Tara Perdida", "Vasilhame"];
const DEFAULT_VASILHAME_NAMES = ["Palete Euro", "Palete Madeira", "Caixa Plástico", "Roll Container"];
const MARGEM_ALERTA_IPO_SEGURO = 30;
const CC_CANONICAL_HEADERS_AT = ["Data", "Tipo", "Metodo", "Fornecedor", "Artigo", "Qtd", "Preco", "Venda", "TaxaIva", "ValorIva", "Dedutivel", "Validado", "Obs", "Link", "Timestamp", "User", "Status", "DataPag", "ValorPago", "ContaStock"];
const CC_COL_ID_ENTIDADE = 25;
const STAFF_HEADERS = ["ID", "Nome", "NIF", "Cargo", "Vencimento Base", "Sub. Alimentação (€/Dia)", "Seguro Acidentes (€/Mês)", "TSU Empresa (%)", "Estado", "Custo Mensal Real (€)", "Subs Férias", "Subs Natal", "Rescisão", "Formações", "Data Admissão", "Dias Contrato", "Prémios", "Email"];
const STAFF_COL_EMAIL = 18;
const CC_HEADERS = ["Data", "Tipo", "Metodo", "Fornecedor", "Artigo", "Qtd", "Preco", "Venda", "TaxaIva", "ValorIva", "Dedutivel", "Validado", "Obs", "Link", "Timestamp", "User", "Status", "DataPag", "ValorPago", "ContaStock", "DocID", "Matricula", "Km_Atuais", "Litros", "ID_Entidade", "Preço Sugerido"];
const CLIENTES_HEADERS = ["ID_Cliente", "Nome_Empresa", "NIF", "Email", "Telefone", "Morada", "Data_Consentimento_RGPD", "Status"];
const AUDIT_HEADERS = ["Timestamp", "Utilizador", "ID_Cliente_Acedido", "Acao", "Motivo_Justificacao"];
const FORNECEDORES_HEADERS = ["ID", "Nome_Empresa", "NIF", "Categoria_Fornecimento", "Email", "Telefone", "Condicoes_Pagamento"];
const MAPEAMENTO_HEADERS = ["ArtigoFlowly", "ContaSoftware"];
const BUILDINGS_HEADERS = ["ID_Predio", "Nome_Predio", "Morada", "NIF_Condominio", "IBAN", "Admin_Email", "Total_Fracoes", "Status", "Stripe_Key", "Notas", "Data_Criacao", "Tem_Elevador", "Data_Manutencao"];
const UNITS_HEADERS = ["ID_Fracao", "ID_Predio", "Designacao", "Proprietario", "NIF_Proprietario", "Email_Proprietario", "Telefone", "Quota_Mensal", "Quinhao_Perc", "Status", "Data_Registo"];

// --- ÍNDICES COLUNAS ---
const MASTER_COL_EMAIL = 2;
const MASTER_COL_LAST_AI_INSIGHT = 7;
const MASTER_COL_DESCONTO = 9;
const MASTER_COL_MENSALIDADES_OFERTA = 10;
const MASTER_COL_DESCONTO_TIPO = 14;
const MASTER_COL_DESCONTO_EXPIRA = 15;
const MASTER_COL_SAAS_PLAN = 16;
const MASTER_COL_AI_CREDITS = 17;
const MASTER_COL_BUSINESS_VERTICAL = 18;
const MASTER_COL_NIF_EMPRESA = 19;
const _BLDG_COL_ID = 0;
const _BLDG_COL_NOME = 1;
const _BLDG_COL_IBAN = 4;
const _BLDG_COL_ADMIN_EMAIL = 5;
const _BLDG_COL_STRIPE_KEY = 8;
const _UNIT_COL_ID = 0;
const _UNIT_COL_ID_PREDIO = 1;
const _UNIT_COL_DESIGNACAO = 2;
const _UNIT_COL_PROPRIETARIO = 3;
const _UNIT_COL_NIF_PROP = 4;
const _UNIT_COL_EMAIL_PROP = 5;
const _UNIT_COL_TELEFONE = 6;
const _UNIT_COL_QUOTA = 7;
const _UNIT_COL_QUINHAO = 8;
const _UNIT_COL_STATUS = 9;
const _UNIT_COL_DATA_REGISTO = 10;
const _UNIT_COL_ID_ENTIDADE = 11;

// --- FUNÇÕES CORE ---
function normalizePlanConfig(parsed) {
  if (!parsed || typeof parsed !== 'object') return {};
  const out = Object.assign({}, parsed);
  if (parsed.financeiro !== undefined) { out.dashFinanceiro = parsed.financeiro; }
  if (parsed.stocks !== undefined) { out.dashStocks = parsed.stocks; }
  if (parsed.rh !== undefined) { out.dashRH = parsed.rh; }
  return out;
}

function ensureMasterDBColumns() {
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = ss.getSheets()[0];
  const lastCol = sheet.getLastColumn();
  if (lastCol < MASTER_COL_LAST_AI_INSIGHT) sheet.getRange(1, MASTER_COL_LAST_AI_INSIGHT).setValue("Last_AI_Insight");
  if (lastCol < MASTER_COL_DESCONTO) sheet.getRange(1, MASTER_COL_DESCONTO).setValue("Desconto (%)");
  if (lastCol < MASTER_COL_MENSALIDADES_OFERTA) sheet.getRange(1, MASTER_COL_MENSALIDADES_OFERTA).setValue("Mensalidades de Oferta (Qtd)");
  if (lastCol < MASTER_COL_DESCONTO_TIPO) sheet.getRange(1, MASTER_COL_DESCONTO_TIPO).setValue("Desconto_Tipo");
  if (lastCol < MASTER_COL_DESCONTO_EXPIRA) sheet.getRange(1, MASTER_COL_DESCONTO_EXPIRA).setValue("Desconto_Expira");
  if (lastCol < MASTER_COL_SAAS_PLAN) sheet.getRange(1, MASTER_COL_SAAS_PLAN).setValue("SaaS_Plan");
  if (lastCol < MASTER_COL_AI_CREDITS) sheet.getRange(1, MASTER_COL_AI_CREDITS).setValue("AI_Credits");
  if (lastCol < MASTER_COL_BUSINESS_VERTICAL) sheet.getRange(1, MASTER_COL_BUSINESS_VERTICAL).setValue("Business_Vertical");
  if (lastCol < MASTER_COL_NIF_EMPRESA) sheet.getRange(1, MASTER_COL_NIF_EMPRESA).setValue("NIF_Empresa");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    for (let r = 2; r <= lastRow; r++) {
      const valDesconto = sheet.getRange(r, MASTER_COL_DESCONTO).getValue();
      if (valDesconto === "" || valDesconto == null) sheet.getRange(r, MASTER_COL_DESCONTO).setValue(0);
      const valOferta = sheet.getRange(r, MASTER_COL_MENSALIDADES_OFERTA).getValue();
      if (valOferta === "" || valOferta == null) sheet.getRange(r, MASTER_COL_MENSALIDADES_OFERTA).setValue(0);
      const valTipo = sheet.getRange(r, MASTER_COL_DESCONTO_TIPO).getValue();
      if (valTipo === "" || valTipo == null) sheet.getRange(r, MASTER_COL_DESCONTO_TIPO).setValue("Permanente");
      const valSaaSPlan = sheet.getRange(r, MASTER_COL_SAAS_PLAN).getValue();
      if (valSaaSPlan === "" || valSaaSPlan == null) sheet.getRange(r, MASTER_COL_SAAS_PLAN).setValue("FREE");
      const valAiCredits = sheet.getRange(r, MASTER_COL_AI_CREDITS).getValue();
      if (valAiCredits === "" || valAiCredits == null) sheet.getRange(r, MASTER_COL_AI_CREDITS).setValue(0);
      const valVertical = sheet.getRange(r, MASTER_COL_BUSINESS_VERTICAL).getValue();
      if (valVertical === "" || valVertical == null) sheet.getRange(r, MASTER_COL_BUSINESS_VERTICAL).setValue("Standard");
    }
  }
}

function getEquipaFlowlySheet() {
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  let sh = ss.getSheetByName('Equipa_Flowly');
  if (!sh) {
    sh = ss.insertSheet('Equipa_Flowly');
    sh.getRange(1, 1, 1, 7).setValues([['Email', 'Nome', 'Cargo', 'Status', 'Password', 'Reset_Token', 'Permissions']]);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sh;
}

function isFlowlyTeamMember(email) {
  if (!email || !email.includes("@")) return false;
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    const emailNorm = String(email).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm && String(data[i][3] || "").trim() === "Ativo") return true;
    }
  } catch (e) { }
  return false;
}

function getFlowlyTeamMemberData(email) {
  if (!email || !email.includes("@")) return null;
  try {
    const sh = getEquipaFlowlySheet();
    const lastCol = Math.max(sh.getLastColumn(), 7);
    const data = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
    const emailNorm = String(email).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm && String(data[i][3] || "").trim() === "Ativo") {
        var permsRaw = data[i][6];
        var permissions = {};
        if (permsRaw && typeof permsRaw === "string" && permsRaw.trim().startsWith("{")) {
          try { permissions = JSON.parse(permsRaw); } catch (e) { }
        } else if (permsRaw && typeof permsRaw === "object") { permissions = permsRaw; }
        var cargo = (data[i][2] || "Admin").toString().trim();
        if (cargo !== "Admin" && cargo !== "Developer" && cargo !== "Owner") cargo = "Admin";
        return { cargo: cargo, permissions: permissions };
      }
    }
  } catch (e) { }
  return null;
}

function getClientContext(targetEmail = null) {
  const activeUser = Session.getActiveUser().getEmail();
  const isSuperAdmin = (activeUser === SUPER_ADMIN_EMAIL);
  const ss = SpreadsheetApp.openById(MASTER_DB_ID);
  const data = ss.getSheets()[0].getDataRange().getValues();

  if (isSuperAdmin && !targetEmail) {
    const adminRow = data.find(r => r[1] === activeUser);
    if (adminRow) {
      let planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG);
      try { if (adminRow[7] && String(adminRow[7]).trim()) { const parsed = JSON.parse(adminRow[7]); if (parsed && typeof parsed === 'object') planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(parsed), parsed); } } catch (e) { }
      return {
        sheetId: adminRow[2] || FLOWLY_OFFICIAL_SHEET_ID,
        folderId: adminRow[3],
        clientName: adminRow[0],
        clientEmail: adminRow[1] || null,
        role: 'SuperAdmin',
        isMaster: true,
        isImpersonating: false,
        planConfig: planConfig,
        businessVertical: String(adminRow[MASTER_COL_BUSINESS_VERTICAL - 1] || "Standard").trim() || "Standard",
        nifEmpresa: String(adminRow[MASTER_COL_NIF_EMPRESA - 1] || "").trim()
      };
    }
    return {
      role: 'SuperAdmin',
      isMaster: true,
      sheetId: FLOWLY_OFFICIAL_SHEET_ID,
      clientName: 'Flowly Test DB',
      planConfig: Object.assign({}, DEFAULT_PLAN_CONFIG)
    };
  }

  if (targetEmail && !isSuperAdmin && activeUser !== "" && targetEmail.toLowerCase().trim() !== activeUser.toLowerCase().trim()) {
    throw new Error("Acesso negado: sem permissão para aceder a dados de outro utilizador.");
  }

  const effectiveIsSuperAdmin = isSuperAdmin || (activeUser === "" && targetEmail && targetEmail.toLowerCase().trim() === SUPER_ADMIN_EMAIL.toLowerCase());
  let emailToSearch = effectiveIsSuperAdmin ? (targetEmail || activeUser) : (targetEmail || activeUser);
  const emailNorm = (emailToSearch || "").toString().trim().toLowerCase();

  if (!emailNorm && !effectiveIsSuperAdmin) {
    throw new Error("Contexto de cliente inválido: é obrigatório fornecer o email do utilizador.");
  }

  if (!effectiveIsSuperAdmin && isFlowlyTeamMember(emailToSearch)) {
    var memberData = getFlowlyTeamMemberData(emailToSearch);
    var planConfig;
    if (memberData && memberData.cargo === "Owner") {
      planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, { admin: true, access: true, gastarCreditos: true });
    } else if (memberData && memberData.permissions && typeof memberData.permissions === "object" && Object.keys(memberData.permissions).length > 0) {
      planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(memberData.permissions), memberData.permissions);
    } else {
      planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, { admin: false, access: false, dashRH: false, dashFinanceiro: false, dashStocks: false, gastarCreditos: false });
    }
    return {
      sheetId: FLOWLY_OFFICIAL_SHEET_ID,
      folderId: null,
      clientName: "Flowly 360",
      clientEmail: emailToSearch,
      role: "FlowlyTeam",
      isMaster: true,
      isImpersonating: false,
      planConfig: planConfig,
      businessVertical: "Standard",
      nifEmpresa: ""
    };
  }

  let clientRow = data.find(r => String(r[1] || "").trim().toLowerCase() === emailNorm);

  if (!clientRow) {
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const sheetId = row[2];
      if (!sheetId || row[4] === "Bloqueado" || row[4] === "Desativado") continue;
      try {
        const ssClient = SpreadsheetApp.openById(sheetId);
        const userSheet = ssClient.getSheetByName(SHEET_USERS_NAME);
        if (!userSheet || userSheet.getLastRow() < 2) continue;
        const users = userSheet.getRange(2, 1, userSheet.getLastRow(), 6).getValues();
        const found = users.some(r => String(r[0] || "").trim().toLowerCase() === emailNorm);
        if (found) { clientRow = row; break; }
      } catch (e) { continue; }
    }
  }

  if (!clientRow) {
    console.error("[CRITICAL] getClientContext: Cliente não localizado para o email:", emailNorm);
    throw new Error("Conta não registada: O email '" + (emailToSearch || "vazio") + "' não foi encontrado na Base de Dados Mestre do Flowly 360. Verifique se o utilizador está corretamente configurado.");
  }

  var clientPlanConfig = Object.assign({}, DEFAULT_PLAN_CONFIG);
  try {
    if (clientRow[7] && String(clientRow[7]).trim()) {
      const parsed = JSON.parse(clientRow[7]);
      if (parsed && typeof parsed === 'object') clientPlanConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(parsed), parsed);
    }
  } catch (e) { }

  return {
    sheetId: clientRow[2],
    folderId: clientRow[3],
    clientName: clientRow[0],
    clientEmail: clientRow[1] || null,
    role: effectiveIsSuperAdmin ? 'SuperAdmin' : 'Gestor',
    isMaster: false,
    isImpersonating: (effectiveIsSuperAdmin && targetEmail),
    planConfig: clientPlanConfig,
    businessVertical: String(clientRow[MASTER_COL_BUSINESS_VERTICAL - 1] || "Standard").trim() || "Standard",
    nifEmpresa: String(clientRow[MASTER_COL_NIF_EMPRESA - 1] || "").trim()
  };
}

function handleLogin(email, password) {
  try {
    const emailTrim = (email || "").toString().trim();
    const passwordTrim = (password || "").toString().trim();

    if (emailTrim.toLowerCase() === SUPER_ADMIN_EMAIL) {
      return { success: true, role: "SuperAdmin", name: "José Reis", email: emailTrim, planConfig: Object.assign({}, DEFAULT_PLAN_CONFIG) };
    }

    const ssMaster = SpreadsheetApp.openById(MASTER_DB_ID);
    const dataMaster = ssMaster.getSheets()[0].getDataRange().getValues();

    const clientRow = dataMaster.find(r => String(r[1] || "").trim().toLowerCase() === emailTrim.toLowerCase());

    if (clientRow) {
      if (clientRow[4] === "Bloqueado" || clientRow[4] === "Desativado") return { success: false, error: "Conta Empresarial Bloqueada. Contacte o suporte." };
      let planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG);
      try { if (clientRow[7]) { const parsed = JSON.parse(clientRow[7]); if (parsed && typeof parsed === 'object') planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(parsed), parsed); } } catch (e) { }

      try {
        const ssClient = SpreadsheetApp.openById(clientRow[2]);
        const userSheet = ssClient.getSheetByName(SHEET_USERS_NAME);
        if (userSheet && userSheet.getLastRow() > 1) {
          const lastRow = userSheet.getLastRow();
          const users = userSheet.getRange(2, 1, lastRow, 6).getValues();
          const userRow = users.find(r => String(r[0] || "").trim().toLowerCase() === emailTrim.toLowerCase());
          if (userRow) {
            if (userRow[4] === "Suspenso") return { success: false, error: "Conta Suspensa. Contacte o administrador." };
            const storedHash = (userRow[1] || "").toString().trim();
            if (!verifyPassword(passwordTrim, storedHash)) return { success: false, error: "Senha incorreta." };
          }
        }
      } catch (e) { }

      return { success: true, role: "Gestor", name: clientRow[0], email: emailTrim, planConfig: planConfig };
    }

    for (let i = 0; i < dataMaster.length; i++) {
      const row = dataMaster[i];
      const sheetId = row[2];
      if (!sheetId) continue;
      if (row[4] === "Bloqueado" || row[4] === "Desativado") continue;
      try {
        const ssClient = SpreadsheetApp.openById(sheetId);
        const userSheet = ssClient.getSheetByName(SHEET_USERS_NAME);
        if (!userSheet || userSheet.getLastRow() < 2) continue;
        const users = userSheet.getRange(2, 1, userSheet.getLastRow(), 6).getValues();
        const userRow = users.find(r => String(r[0] || "").trim().toLowerCase() === emailTrim.toLowerCase());
        if (userRow) {
          if (userRow[4] === "Suspenso") return { success: false, error: "Conta Suspensa. Contacte o administrador." };
          const storedHash = (userRow[1] || "").toString().trim();
          if (!verifyPassword(passwordTrim, storedHash)) return { success: false, error: "Senha incorreta." };
          let planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, { admin: false });
          try { const perms = userRow[5] ? JSON.parse(userRow[5]) : {}; if (perms && typeof perms === 'object') planConfig = Object.assign(planConfig, normalizePlanConfig(perms), perms); } catch (e) { }
          return { success: true, role: "Operador", name: (userRow[2] || emailTrim), email: emailTrim, planConfig: planConfig };
        }
      } catch (e) { continue; }
    }

    try {
      const sh = getEquipaFlowlySheet();
      const teamData = sh.getDataRange().getValues();
      for (var t = 1; t < teamData.length; t++) {
        var r = teamData[t];
        if (String(r[0] || "").trim().toLowerCase() === emailTrim.toLowerCase()) {
          if (String(r[3] || "").trim() !== "Ativo") return { success: false, error: "Conta inativa. Contacte o administrador." };
          var storedHash = (r[4] || "").toString().trim();
          if (!storedHash) return { success: false, error: "Ainda não definiste a tua password. Usa o link enviado por email." };
          if (!verifyPassword(passwordTrim, storedHash)) return { success: false, error: "Senha incorreta." };
          var cargo = (r[2] || "Admin").toString().trim();
          if (cargo !== "Admin" && cargo !== "Developer" && cargo !== "Owner") cargo = "Admin";
          var permissions = {};
          if (r[6] && typeof r[6] === "string" && String(r[6]).trim().startsWith("{")) {
            try { permissions = JSON.parse(r[6]); } catch (e) { }
          } else if (r[6] && typeof r[6] === "object") { permissions = r[6]; }
          var planConfig;
          if (cargo === "Owner") {
            planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, { admin: true, access: true, gastarCreditos: true });
          } else if (Object.keys(permissions).length > 0) {
            planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(permissions), permissions);
          } else {
            planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, { admin: false, access: false, dashRH: false, dashFinanceiro: false, dashStocks: false, gastarCreditos: false });
          }
          return { success: true, role: "FlowlyTeam", name: (r[1] || emailTrim).toString().trim(), email: emailTrim, cargo: cargo, permissions: permissions, planConfig: planConfig };
        }
      }
    } catch (teamErr) { }

    return { success: false, error: "Utilizador não encontrado ou senha inválida." };
  } catch (e) { return { success: false, error: e.toString() }; }

}

function doPost(e) {
  try {
    // Verificar se é uma requisição de login via GitHub Pages
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    
    // Se for uma ação de login, tratar especificamente
    if (data && data.action === 'login') {
      var email = data.email;
      var password = data.password;
      
      // Chamar função de login existente e devolver resposta JSON
      var loginResult = handleLogin(email, password);
      
      // Garantir que o resultado é um objeto simples (sem referências a Sheets)
      var cleanResult = {
        success: loginResult.success || false,
        error: loginResult.error || null,
        role: loginResult.role || null,
        name: loginResult.name || null,
        email: loginResult.email || null,
        planConfig: loginResult.planConfig || null,
        cargo: loginResult.cargo || null,
        permissions: loginResult.permissions || null
      };
      
      return ContentService.createTextOutput(JSON.stringify(cleanResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Código original para captura de leads (mantido para compatibilidade)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Leads');

    if (!sheet) {
      throw new Error("Aba 'Leads' não encontrada. Cria uma aba com este nome no teu Sheets.");
    }

    // Gravar os dados na ordem: Data | Nome | Email | Empresa | Localidade | Setor | Mensagem
    sheet.appendRow([
      new Date(),
      data.name || "N/A",
      data.email || "N/A",
      data.company || "N/A",
      data.location || "N/A",
      data.sector || "N/A",
      data.message || "N/A"
    ]);

    // Resposta de sucesso para o site
    return ContentService.createTextOutput("Sucesso").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    Logger.log("Erro: " + error.toString());
    return ContentService.createTextOutput("Erro: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}