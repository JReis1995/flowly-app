// Ficheiro: MOD_RH.js
/// ==========================================
// 👥 MÓDULO DE RECURSOS HUMANOS (RH)
// ==========================================

function getStaffSheet(ss) {
  let s = ss.getSheetByName(SHEET_STAFF_NAME);
  if (!s) s = ss.getSheetByName(SHEET_STAFF_NAME_LEGACY);
  return s;
}

function saveStaffData(data, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    let sheet = getStaffSheet(ss);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_STAFF_NAME);
      sheet.appendRow(STAFF_HEADERS.concat("Token"));
    }

    let allIds = [];
    if (sheet.getLastRow() > 1) {
      allIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    }
    let r = -1;
    if (data.id && data.id !== "NOVO") r = allIds.indexOf(data.id);
    const row = r > -1 ? r + 2 : sheet.getLastRow() + 1;

    const base = parsePTFloat(data.vencimento) || 0;
    const premios = parsePTFloat(data.premios) || 0;
    const tsu = parsePTFloat(data.tsuPct) || 23.75;
    const seguro = parsePTFloat(data.seguro) || 0;
    const subAlim = parsePTFloat(data.subAlim) || 0;
    const diasContrato = parseInt(data.diasContrato, 10) || 0;

    const now = new Date();
    const diasUteis = getWorkingDaysInMonth(now.getMonth() + 1, now.getFullYear());
    const safeDiasUteis = diasUteis || getWorkingDaysInMonth(now.getMonth() + 1, now.getFullYear());

    const provFerias = base / 12;
    const provNatal = base / 12;
    const baseIncidenciaTSU = base + premios + provFerias + provNatal;
    const tsuValor = baseIncidenciaTSU * (tsu / 100);
    const custoMensalReal = base + premios + (subAlim * safeDiasUteis) + seguro + provFerias + provNatal + tsuValor;

    const provRescisao = parsePTFloat(data.provRescisao) || (base * 0.055);
    const provFormacao = parsePTFloat(data.provFormacao) || (base * 0.02);

    const isNew = (!data.id || data.id === "NOVO");
    const hasEmail = data.email && data.email.includes("@");
    const staffStatus = (isNew && hasEmail) ? "Pendente" : (data.status || "Ativo");

    const rowData = [
      (data.id && data.id !== "NOVO") ? data.id : ("STF-" + Date.now()),
      data.nome, data.nif || "", data.cargo || "", base, subAlim, seguro, tsu,
      staffStatus, custoMensalReal, provFerias, provNatal, provRescisao, provFormacao,
      data.admissao || "", diasContrato, premios, data.email || ""
    ];

    let prevStatus = "";
    if (!isNew && r > -1) {
      try { prevStatus = String(sheet.getRange(row, 9).getValue() || ""); } catch (e) { }
    }

    sheet.getRange(row, 1, 1, 18).setValues([rowData]);

    ensureColumn(sheet, 20, "Data Saída");
    if ((staffStatus === "Inativo" || staffStatus === "Bloqueado") && prevStatus === "Ativo") {
      sheet.getRange(row, 20).setValue(new Date());
    } else if (staffStatus === "Ativo" && (prevStatus === "Inativo" || prevStatus === "Bloqueado")) {
      sheet.getRange(row, 20).setValue("");
    }

    const syncData = Object.assign({}, data, { id: rowData[0], status: staffStatus });
    const syncResult = syncStaffToUsers(syncData, ctx.sheetId);
    return { success: true, emailSent: !!(syncResult && syncResult.emailSent), isNew: isNew };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function syncStaffToUsers(staffData, sheetId) {
  let emailSent = false;
  let permissionsObject = null;
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    let userSheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!userSheet) {
      userSheet = ss.insertSheet(SHEET_USERS_NAME);
      userSheet.appendRow(["Email", "Senha", "Nome", "Perfil", "Estado", "Permissoes"]);
    }

    const defPermsObj = { dashboard: true, cc: false, logistica: true, ia: false, rh: false, admin: false };
    const permsObj = (staffData.permissions && typeof staffData.permissions === "object") ? Object.assign({}, defPermsObj, staffData.permissions) : defPermsObj;
    const permsJson = JSON.stringify(permsObj);
    permissionsObject = permsObj;

    const users = userSheet.getDataRange().getValues();
    const userRowIndex = users.findIndex(r => r[0] === staffData.email);

    if (staffData.status === "Pendente") {
      if (staffData.email && staffData.email.includes("@") && userRowIndex === -1) {
        userSheet.appendRow([staffData.email, "", staffData.nome, "Operador", "Pendente", permsJson]);
        const staffId = staffData.id || "";
        const token = Utilities.getUuid();
        const url = ScriptApp.getService().getUrl();
        const setupLink = url + "?page=setup-password&token=" + encodeURIComponent(token);
        PropertiesService.getScriptProperties().setProperty("STAFF_SETUP_" + token, buildStaffSetupPayload(sheetId, staffData.email, staffId));
        emailSent = sendStaffSetupEmail(staffData.email, staffData.nome || "Colaborador", setupLink);
      }
      return { synced: true, emailSent: emailSent, permissionsObject: permissionsObject };
    }

    if (staffData.status === "Inativo" || staffData.status === "Bloqueado") {
      if (userRowIndex > -1) userSheet.getRange(userRowIndex + 1, 5).setValue("Suspenso");
    } else {
      if (staffData.email && staffData.email.includes("@")) {
        if (userRowIndex === -1) {
          userSheet.appendRow([staffData.email, "", staffData.nome, "Operador", "Pendente", permsJson]);
          const staffId = staffData.id || "";
          const token = Utilities.getUuid();
          const url = ScriptApp.getService().getUrl();
          const setupLink = url + "?page=setup-password&token=" + encodeURIComponent(token);
          PropertiesService.getScriptProperties().setProperty("STAFF_SETUP_" + token, buildStaffSetupPayload(sheetId, staffData.email, staffId));
          emailSent = sendStaffSetupEmail(staffData.email, staffData.nome || "Colaborador", setupLink);
        } else {
          const currentStatus = users[userRowIndex][4];
          if (currentStatus === "Suspenso" || currentStatus === "Pendente") userSheet.getRange(userRowIndex + 1, 5).setValue("Ativo");
          userSheet.getRange(userRowIndex + 1, 3).setValue(staffData.nome);
          userSheet.getRange(userRowIndex + 1, 6).setValue(permsJson);
        }
      }
    }
    return { synced: true, emailSent: emailSent, permissionsObject: permissionsObject };
  } catch (e) { return { synced: false, emailSent: false, error: e.toString() }; }
}

function deactivateStaffMember(id, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = getStaffSheet(ss);
    if (!sheet) return { success: false, error: "Folha de colaboradores não encontrada." };
    const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const idx = ids.indexOf(id);
    if (idx < 0) return { success: false, error: "Colaborador não encontrado." };
    const row = idx + 2;
    sheet.getRange(row, 9).setValue("Bloqueado");
    ensureColumn(sheet, 20, "Data Saída");
    sheet.getRange(row, 20).setValue(new Date());
    const email = sheet.getRange(row, STAFF_COL_EMAIL).getValue();
    if (email) {
      const userSheet = ss.getSheetByName(SHEET_USERS_NAME);
      if (userSheet) {
        const users = userSheet.getDataRange().getValues();
        const uIdx = users.findIndex(r => r[0] === email);
        if (uIdx > -1) userSheet.getRange(uIdx + 1, 5).setValue("Suspenso");
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteStaffMember(id, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = getStaffSheet(ss);
    if (!sheet) return { success: false, error: "Folha de colaboradores não encontrada." };
    const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const idx = ids.indexOf(id);
    if (idx > -1) {
      const email = sheet.getRange(idx + 2, STAFF_COL_EMAIL).getValue();
      sheet.deleteRow(idx + 2);
      if (email) deleteUser(email, impersonateTarget);
      return { success: true };
    }
    return { success: false, error: "ID não encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function anonymizeStaffRecord(staffId, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = getStaffSheet(ss);
    if (!sheet) return { success: false, error: "Folha de colaboradores não encontrada." };
    const ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
    const idx = ids.indexOf(staffId);
    if (idx < 0) return { success: false, error: "ID não encontrado." };
    const row = idx + 2;
    sheet.getRange(row, 2).setValue("[ANONIMIZADO]");
    sheet.getRange(row, 3).setValue("[ANONIMIZADO]");
    sheet.getRange(row, STAFF_COL_EMAIL).setValue("[ANONIMIZADO-" + new Date().toLocaleDateString('pt-PT') + "]");
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveUserPermissions(email, permissions, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_USERS_NAME);
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] === email);
    if (idx > -1) {
      sheet.getRange(idx + 1, 6).setValue(JSON.stringify(permissions));
      return { success: true };
    }
    return { success: false, error: "Utilizador não encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateUserPassword(newPassword, impersonateTarget) {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!sheet) return { success: false, error: "Folha de utilizadores não encontrada." };
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] === activeEmail);
    if (idx > -1) {
      sheet.getRange(idx + 1, 2).setValue(hashPassword(newPassword));
      return { success: true };
    }
    return { success: false, error: "Utilizador não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateUserEmail(newEmail, impersonateTarget) {
  try {
    if (!newEmail || !newEmail.includes("@")) return { success: false, error: "Email inválido." };
    const activeEmail = Session.getActiveUser().getEmail();
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!sheet) return { success: false, error: "Folha não encontrada." };
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] === activeEmail);
    if (idx > -1) {
      sheet.getRange(idx + 1, 1).setValue(newEmail);
      return { success: true };
    }
    return { success: false, error: "Utilizador não encontrado." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteUser(email, impersonateTarget) {
  try {
    const ctx = getClientContext(impersonateTarget);
    const ss = SpreadsheetApp.openById(ctx.sheetId);
    const sheet = ss.getSheetByName(SHEET_USERS_NAME);
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] === email);
    if (idx > -1) sheet.deleteRow(idx + 1);
    return { success: true };
  } catch (e) { return { success: false }; }
}