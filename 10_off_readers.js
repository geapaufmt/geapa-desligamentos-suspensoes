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
  const runId = GEAPA_CORE.coreRunId();
  const sh = off_getMembersSheet_();
  const headerMap = GEAPA_CORE.coreHeaderMap(sh, 1);

  const colRole = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.MEMBERS.COL_ROLE);
  const colEmail = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.MEMBERS.COL_EMAIL);

  if (!colRole || !colEmail) {
    GEAPA_CORE.coreLogError(runId, 'OFF: colunas de membros não encontradas', { colRole, colEmail });
    return [];
  }

  const rolesWanted = new Set((OFF_CFG.MEMBERS.SECRETARY_ROLES || []).map(s => String(s).trim()));
  const values = sh.getDataRange().getValues();
  const out = new Set();

  for (let r = 1; r < values.length; r++) {
    const role = String(values[r][colRole - 1] || '').trim();
    const email = String(values[r][colEmail - 1] || '').trim();
    if (!role || !email) continue;
    if (!rolesWanted.has(role)) continue;
    if (!GEAPA_CORE.coreIsValidEmail(email)) continue;
    out.add(email);
  }

  const list = Array.from(out);
  GEAPA_CORE.coreLogInfo(runId, 'OFF: secretários encontrados', { count: list.length, list });
  return list;
}