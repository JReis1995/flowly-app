// Ficheiro: MOD_Admin.js
/// ==========================================
// 🛡️ MÓDULO DE ADMINISTRAÇÃO E COMUNICAÇÕES
// ==========================================

function sendInviteEmail(targetEmail, name, role, token) {
  try {
    if (targetEmail && targetEmail.includes("@")) {
      const url = ScriptApp.getService().getUrl();
      const activationLink = url + '?action=register&email=' + encodeURIComponent(targetEmail) + '&token=' + encodeURIComponent(token);
      const subject = "Bem-vindo à Flowly 360 — Ative a sua conta";
      const innerBody = `<p style="color:#64748B;font-size:13px;margin:0 0 8px;">Olá,</p><h2 style="color:#020617;font-size:20px;font-weight:800;margin:0 0 20px;">${name}</h2><p style="color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;">A sua conta foi criada na plataforma <strong>Flowly 360</strong> com o perfil de <strong>${role}</strong>.</p><p style="color:#475569;font-size:14px;line-height:1.7;margin:0 0 32px;">Para definir a sua senha e ativar o acesso, clique no botão abaixo:</p><table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center"><a href="${activationLink}" target="_blank" style="display:inline-block;background:#10B981;color:#ffffff;font-size:13px;font-weight:800;text-decoration:none;padding:16px 40px;border-radius:50px;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 4px 16px rgba(16,185,129,0.4);">Definir Palavra-passe</a></td></tr></table><p style="color:#94A3B8;font-size:12px;margin:28px 0 0;line-height:1.6;">Se o botão não funcionar, clique no link abaixo:<br><a href="${activationLink}" style="color:#10B981;font-weight:700;text-decoration:underline;">Definir Palavra-passe</a></p>`;
      const htmlBody = _buildStandardEmailHTML("Flowly 360", innerBody);
      let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
      // [NOTA TÉCNICA] Para configurar o Alias 'noreply@flowly.pt' e "Flowly 360", associe este email como 'Send As' no Gmail que executa o script.
      var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
      GmailApp.sendEmail(targetEmail, subject, "", options);
      return true;
    }
    return false;
  } catch (e) { return false; }
}

function sendRecoveryEmail(email) {
  try {
    if (!email || !email.includes("@")) return { success: false, error: "Email inválido." };

    const url = ScriptApp.getService().getUrl();
    const rawToken = email + ':' + Date.now() + ':' + Math.random().toString(36);
    const token = Utilities.base64EncodeWebSafe(rawToken);

    const props = PropertiesService.getScriptProperties();
    const tokenKey = 'RECOVERY_' + hashPassword(email);
    props.setProperty(tokenKey, JSON.stringify({ token: token, ts: Date.now(), email: email }));

    const recoveryLink = url + '?action=recover&email=' + encodeURIComponent(email) + '&token=' + encodeURIComponent(token);
    const subject = "Recuperar Acesso — Flowly 360";

    const innerBody = `<p style="color:#64748B;font-size:13px;margin:0 0 8px;">Pedido de recuperação de acesso para:</p><h2 style="color:#020617;font-size:18px;font-weight:800;margin:0 0 20px;word-break:break-all;">${email}</h2><p style="color:#475569;font-size:14px;line-height:1.7;margin:0 0 32px;">Se foi você a solicitar a recuperação de acesso, clique no botão abaixo para redefinir a sua palavra-passe. O link é válido por <strong>30 minutos</strong>.</p><table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center"><a href="${recoveryLink}" target="_blank" style="display:inline-block;background:#020617;color:#ffffff;font-size:13px;font-weight:800;text-decoration:none;padding:16px 40px;border-radius:50px;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 4px 16px rgba(2,6,23,0.3);">Recuperar Acesso</a></td></tr></table><p style="color:#94A3B8;font-size:12px;margin:28px 0 0;line-height:1.6;">Se o botão não funcionar, clique no link abaixo:<br><a href="${recoveryLink}" style="color:#06B6D4;font-weight:700;text-decoration:underline;">Recuperar Acesso</a></p><p style="color:#CBD5E1;font-size:11px;margin:20px 0 0;line-height:1.6;">Se não solicitou esta recuperação, ignore este email. A sua conta permanece segura.</p>`;
    const htmlBody = _buildStandardEmailHTML("Flowly 360", innerBody);
    let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
    var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
    GmailApp.sendEmail(email, subject, "", options);
    return { success: true, message: "Link de recuperação enviado para " + email + ". Verifique a sua caixa de entrada." };
  } catch (e) {
    return { success: false, error: "Erro ao enviar email: " + e.toString() };
  }
}

function validateRecoveryToken(email, token) {
  try {
    if (!email || !token) return { valid: false, error: "Parâmetros em falta." };
    const props = PropertiesService.getScriptProperties();
    const tokenKey = 'RECOVERY_' + hashPassword(email);
    const stored = props.getProperty(tokenKey);
    if (!stored) return { valid: false, error: "Token não encontrado ou já utilizado." };
    const data = JSON.parse(stored);
    if (data.email !== email) return { valid: false, error: "Token inválido." };
    if (data.token !== token) return { valid: false, error: "Token inválido." };
    const thirtyMin = 30 * 60 * 1000;
    if ((Date.now() - data.ts) > thirtyMin) {
      props.deleteProperty(tokenKey);
      return { valid: false, error: "Token expirado. Solicite um novo link." };
    }
    props.deleteProperty(tokenKey);
    return { valid: true, email: email };
  } catch (e) { return { valid: false, error: e.toString() }; }
}

function resendInviteEmail(targetEmail) {
  try {
    if (!targetEmail || !targetEmail.includes("@")) return { success: false, error: "Email inválido." };
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    const data = ss.getSheets()[0].getDataRange().getValues();
    const clientRow = data.find(r => r[1] === targetEmail);
    if (!clientRow) return { success: false, error: "Utilizador não encontrado na Base de Dados." };
    const clientName = clientRow[0] || targetEmail;
    const role = "Gestor";
    const token = Utilities.base64EncodeWebSafe(targetEmail + "_" + Date.now());
    const sent = sendInviteEmail(targetEmail, clientName, role, token);
    if (sent) return { success: true };
    return { success: false, error: "Falha ao enviar email." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function sendStaffSetupEmail(targetEmail, name, setupLink) {
  try {
    if (!targetEmail || !targetEmail.includes("@")) return false;
    const subject = "Flowly 360 — Definir a sua palavra-passe";
    const innerBody = `<p style="color:#64748B;font-size:13px;margin:0 0 8px;">Olá,</p><h2 style="color:#020617;font-size:20px;font-weight:700;margin:0 0 20px;">${name}</h2><p style="color:#475569;font-size:14px;line-height:1.7;margin:0 0 32px;">Foi convidado para aceder à plataforma Flowly 360. Use o link abaixo para definir a sua palavra-passe e ativar o acesso.</p><table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center"><a href="${setupLink}" target="_blank" style="display:inline-block;background:#10B981;color:#ffffff;font-size:14px;font-weight:800;text-decoration:none;padding:16px 40px;border-radius:50px;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 4px 16px rgba(16,185,129,0.4);">Definir palavra-passe</a></td></tr></table><p style="color:#94A3B8;font-size:12px;margin:28px 0 0;line-height:1.6;">Se o botão não funcionar, copie e cole este link no browser: <a href="${setupLink}" style="color:#10B981;font-weight:600;text-decoration:underline;">${setupLink}</a></p>`;
    const htmlBody = _buildStandardEmailHTML("Flowly 360", innerBody);
    let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
    var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
    try {
      GmailApp.sendEmail(targetEmail, subject, "", options);
      return true;
    } catch (e1) {
      try {
        var optionsFallback = { name: "Flowly 360", from: "noreply@flowly.pt" };
        GmailApp.sendEmail(targetEmail, subject, "Defina a sua palavra-passe em: " + setupLink, optionsFallback);
        return true;
      } catch (e2) { return false; }
    }
  } catch (e) { return false; }
}

function inviteStaff(staffData, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada. Use Modo Espião." };
    const email = (staffData.email || "").trim();
    const nome = (staffData.nome || "").trim();
    if (!email || !email.includes("@")) return { success: false, error: "Email inválido." };

    const ss = SpreadsheetApp.openById(ctx.sheetId);
    let sheet = getStaffSheet(ss);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_STAFF_NAME);
      sheet.appendRow(STAFF_HEADERS.concat("Token"));
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 19) sheet.getRange(1, 19).setValue("Token");

    const token = Utilities.getUuid();
    const staffId = "STF-" + Date.now();
    const rowData = [staffId, nome || email, staffData.nif || "", staffData.cargo || "", 0, 0, 0, 23.75, "Pendente", 0, 0, 0, 0, 0, staffData.admissao || "", 0, 0, email, token];
    sheet.appendRow(rowData);

    const permsObj = (staffData.permissions && typeof staffData.permissions === "object") ? Object.assign({ dashboard: true, cc: false, logistica: true, ia: false, rh: false, admin: false }, staffData.permissions) : { dashboard: true, cc: false, logistica: true, ia: false, rh: false, admin: false };
    const permsJson = JSON.stringify(permsObj);
    let userSheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!userSheet) {
      userSheet = ss.insertSheet(SHEET_USERS_NAME);
      userSheet.appendRow(["Email", "Senha", "Nome", "Perfil", "Estado", "Permissoes"]);
    }
    const users = userSheet.getDataRange().getValues();
    if (users.findIndex(r => r[0] === email) === -1) {
      userSheet.appendRow([email, "", nome || email, "Operador", "Pendente", permsJson]);
    }

    const url = ScriptApp.getService().getUrl();
    const setupLink = url + "?page=setup-password&token=" + encodeURIComponent(token);
    const sent = sendStaffSetupEmail(email, nome || "Colaborador", setupLink);
    if (!sent) return { success: false, error: "Falha ao enviar email." };

    PropertiesService.getScriptProperties().setProperty("STAFF_SETUP_" + token, buildStaffSetupPayload(ctx.sheetId, email, staffId));
    return { success: true, message: "Convite enviado para " + email };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function buildStaffSetupPayload(sheetId, email, staffId) {
  const expiry = Date.now() + 172800000;
  return JSON.stringify({ email: email, sheetId: sheetId, staffId: staffId, expiry: expiry });
}

function validateStaffSetupToken(token) {
  const key = "STAFF_SETUP_" + (token || "");
  const stored = PropertiesService.getScriptProperties().getProperty(key);
  if (!stored) return { valid: false, error: "Este link de convite é inválido ou já expirou. Peça ao administrador para reenviar o convite." };
  try {
    const data = JSON.parse(stored);
    const expiry = Number(data.expiry);
    if (!expiry || Date.now() >= expiry) return { valid: false, error: "Este link de convite é inválido ou já expirou. Peça ao administrador para reenviar o convite." };
    return { valid: true, email: data.email, sheetId: data.sheetId, staffId: data.staffId };
  } catch (e) { return { valid: false, error: "Este link de convite é inválido ou já expirou. Peça ao administrador para reenviar o convite." }; }
}

function completeStaffSetupPassword(token, newPassword) {
  try {
    const val = validateStaffSetupToken(token);
    if (!val.valid) return { success: false, error: val.error };
    if (!newPassword || newPassword.length < 6) return { success: false, error: "A palavra-passe deve ter pelo menos 6 caracteres." };

    const ss = SpreadsheetApp.openById(val.sheetId);
    const staffSheet = getStaffSheet(ss);
    if (!staffSheet) return { success: false, error: "Folha de colaboradores não encontrada." };

    const ids = staffSheet.getRange(2, 1, staffSheet.getLastRow(), 1).getValues().flat();
    const idx = ids.indexOf(val.staffId);
    if (idx < 0) return { success: false, error: "Colaborador não encontrado." };
    const row = idx + 2;
    const nome = staffSheet.getRange(row, 2).getValue();

    let userSheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!userSheet) {
      userSheet = ss.insertSheet(SHEET_USERS_NAME);
      userSheet.appendRow(["Email", "Senha", "Nome", "Perfil", "Estado", "Permissoes"]);
    }
    const users = userSheet.getDataRange().getValues();
    const userIdx = users.findIndex(r => r[0] === val.email);
    const perms = JSON.stringify({ dashboard: true, cc: false, logistica: true, ia: false, rh: false, admin: false });
    try {
      const pwdHash = hashPassword(newPassword);
      if (userIdx > -1) {
        userSheet.getRange(userIdx + 1, 2).setValue(pwdHash);
        userSheet.getRange(userIdx + 1, 5).setValue("Ativo");
      } else {
        userSheet.appendRow([val.email, pwdHash, nome, "Operador", "Ativo", perms]);
      }
    } catch (eEnc) { return { success: false, error: "Erro ao gravar palavra-passe. Tente novamente." }; }

    const tokenCol = 19;
    staffSheet.getRange(row, 9).setValue("Ativo");
    try { if (staffSheet.getLastColumn() >= tokenCol) staffSheet.getRange(row, tokenCol).setValue(""); } catch (e) { }
    PropertiesService.getScriptProperties().deleteProperty("STAFF_SETUP_" + token);
    return { success: true, message: "Conta ativada com sucesso." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function resendStaffInvite(emailOrId, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    if (!ctx.sheetId) return { success: false, error: "Infraestrutura não encontrada." };
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = getStaffSheet(ss);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: "Nenhum colaborador encontrado." };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    let rowIdx = -1;
    if (emailOrId.includes("@")) {
      rowIdx = data.findIndex(r => r[17] === emailOrId);
    } else {
      rowIdx = data.findIndex(r => r[0] === emailOrId);
    }
    if (rowIdx < 0) return { success: false, error: "Colaborador não encontrado." };

    const row = data[rowIdx];
    const email = row[17];
    const nome = row[1] || "Colaborador";
    const staffId = row[0];
    const oldToken = (row[18] || "").toString();
    if (oldToken) PropertiesService.getScriptProperties().deleteProperty("STAFF_SETUP_" + oldToken);

    const token = Utilities.getUuid();
    sheet.getRange(rowIdx + 2, 19).setValue(token);
    sheet.getRange(rowIdx + 2, 9).setValue("Pendente");
    const url = ScriptApp.getService().getUrl();
    const setupLink = url + "?page=setup-password&token=" + encodeURIComponent(token);
    const sent = sendStaffSetupEmail(email, nome, setupLink);
    if (!sent) return { success: false, error: "Falha ao enviar email." };
    PropertiesService.getScriptProperties().setProperty("STAFF_SETUP_" + token, buildStaffSetupPayload(ctx.sheetId, email, staffId));
    return { success: true, message: "Convite reenviado para " + email };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getFlowlyTeam(callerEmail) {
  const activeUser = Session.getActiveUser().getEmail();
  const callerNorm = (callerEmail || activeUser || "").toString().trim().toLowerCase();
  const allowed = (activeUser === SUPER_ADMIN_EMAIL) || (callerNorm === SUPER_ADMIN_EMAIL.toLowerCase()) || isFlowlyTeamMember(callerEmail || activeUser);
  if (!allowed) return { success: false, error: "Sem permissão", team: [] };
  try {
    const sh = getEquipaFlowlySheet();
    const lastCol = Math.max(sh.getLastColumn(), 7);
    const data = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
    if (!data || data.length < 2) return { success: true, team: [] };
    const team = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var permsRaw = row[6];
      var permissions = {};
      if (permsRaw && typeof permsRaw === "string" && permsRaw.trim().startsWith("{")) {
        try { permissions = JSON.parse(permsRaw); } catch (e) { }
      } else if (permsRaw && typeof permsRaw === "object") { permissions = permsRaw; }
      team.push({ email: (row[0] || "").toString().trim(), nome: (row[1] || "").toString().trim(), cargo: (row[2] || "").toString().trim(), status: (row[3] || "").toString().trim(), permissions: permissions });
    }
    return { success: true, team: team };
  } catch (e) { return { success: false, error: e.toString(), team: [] }; }
}

function addFlowlyTeamMember(dados) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!dados || typeof dados !== "object") return { success: false, error: "Dados inválidos" };
  var email = (dados.email || "").toString().trim();
  if (!email || !email.includes("@")) return { success: false, error: "Email inválido" };
  var nome = (dados.nome || "").toString().trim();
  var cargo = (dados.cargo || "Admin").toString().trim();
  if (cargo !== "Admin" && cargo !== "Developer" && cargo !== "Owner") cargo = "Admin";
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    const emailNorm = email.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm) return { success: false, error: "Este email já está registado na equipa." };
    }
    var resetToken = Utilities.getUuid();
    var url = ScriptApp.getService().getUrl();
    var setupLink = url + "?action=setpassword&token=" + encodeURIComponent(resetToken);
    sh.appendRow([email, nome, cargo, "Ativo", "", resetToken]);
    var subject = "Bem-vindo à equipa Flowly - Define a tua password";
    var innerBody = "<p style=\"color:#64748B;font-size:13px;margin:0 0 8px;\">Olá " + (nome || "Membro") + ",</p><p style=\"color:#475569;font-size:14px;line-height:1.7;margin:0 0 16px;\">Foste convidado para a equipa Flowly (" + cargo + ").</p><p style=\"color:#475569;font-size:14px;line-height:1.7;margin:0 0 32px;\">Clica no link abaixo para definires a tua password e acederes ao Painel Admin:</p><table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr><td align=\"center\"><a href=\"" + setupLink + "\" target=\"_blank\" style=\"display:inline-block;background:#10B981;color:#ffffff;font-size:13px;font-weight:800;text-decoration:none;padding:16px 40px;border-radius:50px;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 4px 16px rgba(16,185,129,0.4);\">Definir Password</a></td></tr></table><p style=\"color:#94A3B8;font-size:12px;margin:28px 0 0;line-height:1.6;\">Se o botão não funcionar, copia e cola: <a href=\"" + setupLink + "\" style=\"color:#10B981;font-weight:700;\">" + setupLink + "</a></p>";
    var htmlBody = _buildStandardEmailHTML("Flowly 360", innerBody);
    let logoBlob = UrlFetchApp.fetch("https://i.postimg.cc/mrcDM13S/flowly-logo.jpg").getBlob().setName("flowlyLogo");
    var options = { name: "Flowly 360", from: "noreply@flowly.pt", htmlBody: htmlBody, inlineImages: { flowlyLogo: logoBlob } };
    try { GmailApp.sendEmail(email, subject, "", options); } catch (mailErr) { }
    return { success: true, message: "Membro adicionado. Email de convite enviado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveNewTeamPassword(token, novaSenha) {
  if (!token || !novaSenha) return { success: false, error: "Token ou senha inválidos" };
  if (novaSenha.length < 6) return { success: false, error: "A senha deve ter pelo menos 6 caracteres." };
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][5] || "").trim() === String(token).trim()) {
        var rowNum = i + 1;
        var hashed = hashPassword(novaSenha);
        sh.getRange(rowNum, 5).setValue(hashed);
        sh.getRange(rowNum, 6).setValue("");
        return { success: true, message: "Password definida com sucesso." };
      }
    }
    return { success: false, error: "Link inválido ou já expirado. Contacta o administrador." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateTeamMember(email, dados) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!email || !email.includes("@")) return { success: false, error: "Email inválido" };
  if (!dados || typeof dados !== "object") return { success: false, error: "Dados inválidos" };
  var nome = (dados.nome || "").toString().trim();
  var cargo = (dados.cargo || "").toString().trim();
  if (cargo !== "Admin" && cargo !== "Developer" && cargo !== "Owner") cargo = "Admin";
  var permissions = dados.permissions;
  var permsStr = "";
  if (permissions && typeof permissions === "object") { try { permsStr = JSON.stringify(permissions); } catch (e) { } }
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    const emailNorm = String(email).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm) {
        sh.getRange(i + 1, 2).setValue(nome);
        sh.getRange(i + 1, 3).setValue(cargo);
        if (permsStr) sh.getRange(i + 1, 7).setValue(permsStr);
        var finalPerms = permissions && typeof permissions === "object" ? permissions : {};
        try { if (data[i][6] && String(data[i][6]).trim().startsWith("{")) finalPerms = JSON.parse(data[i][6]); } catch (e2) { }
        if (permsStr) try { finalPerms = JSON.parse(permsStr); } catch (e3) { }
        var planConfig = Object.assign({}, DEFAULT_PLAN_CONFIG, normalizePlanConfig(finalPerms), finalPerms);
        try { CacheService.getScriptCache().remove("SA_DASHBOARD"); } catch (cErr) { }
        return { success: true, message: "Membro atualizado.", updatedUser: { email: email, nome: nome, cargo: cargo, permissions: finalPerms, planConfig: planConfig, role: "FlowlyTeam" } };
      }
    }
    return { success: false, error: "Membro não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function toggleTeamMemberStatus(email) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!email || !email.includes("@")) return { success: false, error: "Email inválido" };
  var emailNorm = String(email).trim().toLowerCase();
  if (emailNorm === SUPER_ADMIN_EMAIL.toLowerCase()) return { success: false, error: "Não é possível bloquear o Owner principal." };
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm) {
        var currentStatus = String(data[i][3] || "").trim();
        var newStatus = (currentStatus === "Ativo") ? "Inativo" : "Ativo";
        sh.getRange(i + 1, 4).setValue(newStatus);
        return { success: true, message: "Status alterado para " + newStatus + ".", newStatus: newStatus };
      }
    }
    return { success: false, error: "Membro não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteTeamMember(email) {
  if (Session.getActiveUser().getEmail() !== SUPER_ADMIN_EMAIL) return { success: false, error: "Sem permissão" };
  if (!email || !email.includes("@")) return { success: false, error: "Email inválido" };
  var emailNorm = String(email).trim().toLowerCase();
  if (emailNorm === SUPER_ADMIN_EMAIL.toLowerCase()) return { success: false, error: "Não é possível eliminar o Owner principal." };
  try {
    const sh = getEquipaFlowlySheet();
    const data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === emailNorm) {
        sh.deleteRow(i + 1);
        return { success: true, message: "Membro eliminado." };
      }
    }
    return { success: false, error: "Membro não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}
/**
 * 📝 FERRAMENTA INTERNA DE DEV: Regista alterações no Glossário / Changelog
 * Como usar: Na tua IDE, altera os valores dentro desta função e clica em "Executar".
 */
function logProjectChange(moduloModificado = "MOD_CC_Dashboard", descricaoAlteracao = "Corrigi o cálculo da margem de lucro líquido.") {
  try {
    // O ID do teu Google Sheets do Glossário
    const sheetId = "1C609Pg7bpphgmIlIYbSGD75nRl3P5tswC3trFeBfhjo";
    const ss = SpreadsheetApp.openById(sheetId);

    // Tenta abrir a aba Changelog
    let sheet = ss.getSheetByName("Changelog");
    if (!sheet) {
      sheet = ss.insertSheet("Changelog");
      sheet.appendRow(["Data", "Programador", "Módulo", "Alteração Feita"]);
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#E2E8F0");
    }

    // Quem está a fazer a alteração
    const devEmail = Session.getActiveUser().getEmail() || "Super Admin";
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    // Insere a alteração
    sheet.appendRow([dataAtual, devEmail, moduloModificado, descricaoAlteracao]);

    Logger.log("✅ Sucesso! A alteração foi registada na aba Changelog do teu Glossário.");
  } catch (e) {
    Logger.log("❌ Erro ao registar: " + e.toString());
  }
}