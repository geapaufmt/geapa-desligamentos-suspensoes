/***************************************
 * 10_off_readers.gs
 * Leitura de dados (Sheets) com segurança.
 ***************************************/

/**
 * Abre a aba de respostas do Forms via Registry KEY.
 */
function off_getResponsesSheet_() {
  return GEAPA_CORE.coreGetSheetByKey(OFF_KEYS.RESPONSES);
}

/**
 * Abre a base de membros via Registry KEY.
 */
function off_getMembersSheet_() {
  return GEAPA_CORE.coreGetSheetByKey(OFF_KEYS.MEMBERS);
}

/**
 * Lê todas as respostas (linhas) e retorna:
 * { headerMap, values, startRow }
 */
function off_readResponsesTable_() {
  const sh = off_getResponsesSheet_();
  const headerRow = OFF_CFG.RESP.HEADER_ROW || 1;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < headerRow + 1) {
    return { sh, headerMap: {}, rows: [], startRow: headerRow + 1 };
  }

  const headerMap = GEAPA_CORE.coreHeaderMap(sh, headerRow);

  const rows = sh.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  return { sh, headerMap, rows, startRow: headerRow + 1 };
}

/**
 * Busca emails dos secretários (base de membros) pela coluna de cargo.
 */
function off_getSecretaryEmails_() {
  const boardSheet = GEAPA_CORE.coreGetSheetByKey("VIGENCIA_MEMBROS_DIRETORIAS");
  const currentSheet = GEAPA_CORE.coreGetSheetByKey("MEMBERS_ATUAIS");

  if (!boardSheet) {
    throw new Error("Não foi possível localizar VIGENCIA_MEMBROS_DIRETORIAS.");
  }
  if (!currentSheet) {
    throw new Error("Não foi possível localizar MEMBERS_ATUAIS.");
  }

  const secretaryRgas = off_getCurrentSecretaryRgas_(boardSheet);
  if (!secretaryRgas.length) return [];

  const emails = off_getEmailsByRgaFromCurrentMembers_(currentSheet, secretaryRgas);

  return [...new Set(emails.filter(Boolean))];
}

function off_getCurrentSecretaryRgas_(boardSheet) {
  const lastRow = boardSheet.getLastRow();
  const lastCol = boardSheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = boardSheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || "").trim());

  const rows = boardSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = offboard_getHeaderMap_(headers);

  const rgaIdx = off_pickHeaderIndex_(map, ["rga"]);
  const roleIdx = off_pickHeaderIndex_(map, ["cargo/função", "cargo/funcao", "cargo"]);
  const startIdx = off_pickHeaderIndex_(map, ["data_início", "data_inicio"]);
  const endIdx = off_pickHeaderIndex_(map, ["data_fim"]);
  const plannedEndIdx = off_pickHeaderIndex_(map, ["data_fim_previsto"]);

  if (rgaIdx == null || roleIdx == null) {
    throw new Error('Cabeçalhos "RGA" e/ou "Cargo/Função" não encontrados em VIGENCIA_MEMBROS_DIRETORIAS.');
  }

  const now = new Date();

  return rows
    .filter(row => {
      const role = off_normalizeSecretaryText_(row[roleIdx]);

      // aceita Secretário(a) Geral, Secretário(a) Executivo etc.
      if (!role.includes("SECRETAR")) return false;

      const start = startIdx != null ? off_toDate_(row[startIdx]) : null;
      const end = endIdx != null ? off_toDate_(row[endIdx]) : null;
      const plannedEnd = plannedEndIdx != null ? off_toDate_(row[plannedEndIdx]) : null;

      if (start && now < start) return false;
      if (end && now > end) return false;
      if (!end && plannedEnd && now > plannedEnd) return false;

      return true;
    })
    .map(row => off_normalizeRga_(row[rgaIdx]))
    .filter(Boolean);
}

function off_getEmailsByRgaFromCurrentMembers_(currentSheet, rgas) {
  const target = new Set(
    rgas.map(r => off_normalizeRga_(r)).filter(Boolean)
  );

  const lastRow = currentSheet.getLastRow();
  const lastCol = currentSheet.getLastColumn();
  if (lastRow < 2 || !target.size) return [];

  const headers = currentSheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || "").trim());

  const rows = currentSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = offboard_getHeaderMap_(headers);

  const rgaIdx = off_pickHeaderIndex_(map, ["rga"]);
  const emailIdx = off_pickHeaderIndex_(map, ["email"]);

  if (rgaIdx == null || emailIdx == null) {
    throw new Error('Cabeçalhos "RGA" e/ou "EMAIL" não encontrados em MEMBERS_ATUAIS.');
  }

  return rows
    .filter(row => target.has(off_normalizeRga_(row[rgaIdx])))
    .map(row => String(row[emailIdx] || "").trim().toLowerCase())
    .filter(Boolean);
}

function off_pickHeaderIndex_(map, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const key = String(candidates[i] || "").trim().toLowerCase();
    if (map[key] != null) return map[key];
  }
  return null;
}

function off_normalizeSecretaryText_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function off_normalizeRga_(value) {
  return String(value || "").replace(/\D/g, "").trim();
}

function off_toDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }
  const dt = new Date(value);
  return isNaN(dt) ? null : dt;
}