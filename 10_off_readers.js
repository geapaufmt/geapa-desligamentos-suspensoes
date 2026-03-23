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
 * Busca emails dos secretários usando a projeção institucional do core.
 */
function off_getSecretaryEmails_() {
  const emails = GEAPA_CORE.coreGetCurrentEmailsByEmailGroup('SECRETARIA');
  return [...new Set((emails || []).filter(Boolean))];
}
