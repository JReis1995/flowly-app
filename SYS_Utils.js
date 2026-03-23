// Ficheiro: SYS_Utils.js
/// ==========================================
// 🛠️ FUNÇÕES UTILITÁRIAS
// ==========================================

function hashPassword(plainText) {
  if (!plainText) return "";
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plainText, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyPassword(plainText, storedHash) {
  if (!plainText || !storedHash) return false;
  if (storedHash.length < 50) return plainText === storedHash;
  return hashPassword(plainText) === storedHash;
}

function maskEmail(val) {
  if (!val || typeof val !== "string") return "***";
  var s = String(val).trim();
  if (s.length < 4) return "***";
  var at = s.indexOf("@");
  if (at < 0) return s.substring(0, 2) + "***";
  return s.substring(0, 2) + "***" + s.substring(at);
}

function isValidNIFPT(nif) {
  if (!nif) return false;
  var s = String(nif).replace(/\D/g, "");
  if (s.length !== 9) return false;
  var b = s.substring(0, 1), p = s.substring(0, 2);
  var v = ["1","2","3","5","6","8","9","45","70","71","72","74","75","77","78","79","90","91","98","99"];
  if (v.indexOf(b) === -1 && v.indexOf(p) === -1) return false;
  var sum = 0;
  for (var i = 0; i < 8; i++) { sum += parseInt(s.charAt(i), 10) * (9 - i); }
  var e = 11 - (sum % 11);
  if (e >= 10) e = 0;
  return e === parseInt(s.charAt(8), 10);
}

/** Algoritmo Módulo 11 para NIF (Servidor) */
function calcularCheckDigitNIF(nif8) {
  let s = String(nif8).padStart(8, '0');
  let soma = (parseInt(s[0])*9 + parseInt(s[1])*8 + parseInt(s[2])*7 + parseInt(s[3])*6 + parseInt(s[4])*5 + parseInt(s[5])*4 + parseInt(s[6])*3 + parseInt(s[7])*2);
  let resto = soma % 11;
  let nif9 = (resto < 2) ? 0 : 11 - resto;
  return s + String(nif9);
}

function maskNIF(val) {
  if (!val || typeof val !== "string") return "***";
  var s = String(val).trim().replace(/\D/g, "");
  if (s.length < 7) return "***";
  return s.substring(0, 3) + "***" + s.substring(s.length - 3);
}

function maskTelefone(val) {
  if (!val || typeof val !== "string") return "***";
  var s = String(val).trim();
  if (s.length < 4) return "***";
  var last3 = s.replace(/\D/g, "").slice(-3);
  return "+351 *** *** " + last3;
}

function _maskSensitiveField(value, prefixKeep, suffixKeep) {
  if (!value || typeof value !== "string" || value.trim() === "") return "";
  var s = value.trim();
  if (s.length <= (prefixKeep + suffixKeep)) return "****";
  return s.substring(0, prefixKeep) + "****" + s.substring(s.length - suffixKeep);
}

function parsePTFloat(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  const str = String(val).replace(/\s/g, '');
  const cleanStr = str.replace(/\./g, '').replace(',', '.');
  const num = parseFloat(cleanStr);
  return isNaN(num) ? 0 : num;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) return val;
  if (typeof val === 'number' && val > 20000) return new Date(Math.round((val - 25569) * 86400000));
  const s = String(val).trim();
  const pts = s.split("/");
  if (pts.length === 3) return new Date(pts[2], pts[1] - 1, pts[0]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function toDBDate(isoOrDate) {
  if (!isoOrDate) return "";
  const d = parseDate(isoOrDate);
  if (!d || isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return day + "/" + month + "/" + year;
}

function toDBDateFromInput(yyyyMmDd) {
  if (!yyyyMmDd || typeof yyyyMmDd !== "string") return "";
  const s = String(yyyyMmDd).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
  const parts = s.split("-");
  return parts[2] + "/" + parts[1] + "/" + parts[0];
}

function fromDBDate(dbStr) {
  const d = parseDate(dbStr);
  if (!d || isNaN(d.getTime())) return "";
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function getWorkingDaysInMonth(month, year) {
  let count = 0;
  const daysInMonth = new Date(year, month, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dayOfWeek = new Date(year, month - 1, d).getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) count++;
  }
  return count;
}

function getStaffCostFactor(status, admissao, dataSaida, refMonth, refYear) {
  if (status === "Pendente") return 0;
  if (status === "Ativo") {
    const admDate = parseDate(admissao);
    if (admDate && admDate.getMonth() + 1 === refMonth && admDate.getFullYear() === refYear) {
      const totalDias = getWorkingDaysInMonth(refMonth, refYear);
      let diasAtivos = 0;
      const daysInMonth = new Date(refYear, refMonth, 0).getDate();
      for (let d = admDate.getDate(); d <= daysInMonth; d++) {
        const dow = new Date(refYear, refMonth - 1, d).getDay();
        if (dow !== 0 && dow !== 6) diasAtivos++;
      }
      return totalDias > 0 ? diasAtivos / totalDias : 1;
    }
    return 1;
  }
  const saida = parseDate(dataSaida);
  if (!saida) return 0;
  const saidaMonth = saida.getMonth() + 1;
  const saidaYear = saida.getFullYear();
  if (saidaYear < refYear || (saidaYear === refYear && saidaMonth < refMonth)) return 0;
  if (saidaYear > refYear || (saidaYear === refYear && saidaMonth > refMonth)) return 1;
  const totalDias = getWorkingDaysInMonth(refMonth, refYear);
  let diasAtivos = 0;
  for (let d = 1; d <= saida.getDate(); d++) {
    const dow = new Date(refYear, refMonth - 1, d).getDay();
    if (dow !== 0 && dow !== 6) diasAtivos++;
  }
  return totalDias > 0 ? diasAtivos / totalDias : 0;
}

function ensureColumn(sheet, colNum, headerName) {
  if (sheet.getLastColumn() < colNum) {
    sheet.getRange(1, colNum).setValue(headerName);
  }
}

function round2(val) {
  if (val == null || isNaN(val)) return 0;
  return Math.round(parseFloat(val) * 100) / 100;
}

function isDateInRange(d, rangeStart, rangeEnd) {
  if (!(d instanceof Date) || isNaN(d)) return false;
  const dDay = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const sDay = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), rangeStart.getDate()).getTime();
  const eDay = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), rangeEnd.getDate()).getTime();
  return dDay >= sDay && dDay <= eDay;
}

function normalizeTipo(str) {
  if (str == null || str === "") return "";
  return String(str).toLowerCase().trim()
    .replace(/[àáâãä]/g, 'a').replace(/[éêë]/g, 'e').replace(/[íîï]/g, 'i')
    .replace(/[óôõö]/g, 'o').replace(/[úûü]/g, 'u').replace(/ç/g, 'c')
    .replace(/\s*\(ft\)/g, '');
}

function normalizeTaxaIva(val) {
  const v = parseInt(String(val || "").replace("%", "").trim()) || 23;
  if (v <= 6) return 6;
  if (v <= 13) return 13;
  return 23;
}

function _indexOfHeader(headers, variants) {
  for (var v = 0; v < variants.length; v++) {
    var idx = headers.indexOf(variants[v]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function _diasAteData(dataStr, refDate) {
  if (!dataStr || typeof dataStr !== "string") return null;
  var m = dataStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) {
    m = dataStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (m) { dataStr = m[3] + "-" + m[2] + "-" + m[1]; m = dataStr.match(/^(\d{4})-(\d{2})-(\d{2})/); }
    else return null;
  }
  if (!m) return null;
  var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  var ref = refDate || new Date();
  ref.setHours(0, 0, 0, 0);
  return Math.ceil((d - ref) / 86400000);
}

function _buildStandardEmailHTML(subjectTitle, innerBody) {
  return `<!DOCTYPE html><html lang="pt">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 8px; text-align: left; }
    ul { margin: 0; padding-left: 20px; }
  </style>
</head>
<body style="margin:0;padding:0;background:#F0F4F8;font-family:'Segoe UI',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F0F4F8;padding:40px 20px;">
    <tr>
      <td align="center">
        <table width="560" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:24px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08); max-width: 100%;">
          <tr>
            <td style="background:#020617;padding:36px 40px;text-align:center;">
              <a href="https://www.flowly.pt" target="_blank">
                <img src="cid:flowlyLogo" alt="Flowly 360" width="80" style="border-radius:16px;border:3px solid rgba(255,255,255,0.15);display:block;margin:0 auto 16px;">
              </a>
              <h1 style="color:#ffffff;font-size:22px;font-weight:800;margin:0;letter-spacing:-0.5px;">${subjectTitle}</h1>
              <p style="color:#06B6D4;font-size:11px;font-weight:700;margin:6px 0 0;letter-spacing:2px;text-transform:uppercase;">Onde o fluxo encontra a precisão</p>
            </td>
          </tr>
          <tr>
            <td style="padding:40px 40px 32px; font-size: 14px; color: #475569; line-height: 1.7; overflow-x: auto;">
              ${innerBody}
            </td>
          </tr>
          <tr>
            <td style="padding:0 40px;"><hr style="border:none;border-top:1px solid #E2E8F0;margin:0;"></td>
          </tr>
          <tr>
            <td style="padding:28px 40px;background:#F8FAFC;border-radius:0 0 24px 24px;">
              <p style="color:#020617;font-size:13px;font-weight:700;margin:0;">Atenciosamente,</p>
              <p style="color:#020617;font-size:13px;font-weight:700;margin:4px 0 0;"><strong>A Equipa Flowly 360</strong></p>
              <p style="color:#64748B;font-size:12px;margin:8px 0 0;line-height:1.6;">
                <a href="mailto:geral@flowly.pt" style="color:#10B981;font-weight:600;">geral@flowly.pt</a><br>
                www.flowly.pt<br>
                +351 927140717
              </p>
              <p style="color:#94A3B8;font-size:11px;margin:16px 0 0;font-style:italic;">Onde o fluxo encontra a precisão</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;
}

function logSystemError(module, error, userEmail) {
  try {
    const errorMsg = error instanceof Error ? error.message : String(error);
    const stackTrace = error instanceof Error ? error.stack || "" : "";
    const ss = SpreadsheetApp.openById(MASTER_DB_ID);
    let sheet = ss.getSheetByName("System_Logs");
    if (!sheet) {
      sheet = ss.insertSheet("System_Logs");
      sheet.appendRow(["Timestamp", "Módulo", "Erro", "Utilizador", "StackTrace"]);
      sheet.getRange("A1:E1").setFontWeight("bold");
    }
    sheet.appendRow([new Date(), module, errorMsg, userEmail || "Desconhecido", stackTrace]);
  } catch (e) {
    console.error("Erro fatal ao registar erro do sistema:", e);
  }
}
