/***************************************
 * 01_off_registry.gs
 * Acesso a planilhas, leitura por cabecalho e logs.
 ***************************************/

/**
 * Gera um runId consistente mesmo quando a library nao estiver disponivel.
 */
function off_newRunId_() {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreRunId === 'function') {
    return GEAPA_CORE.coreRunId();
  }
  return 'OFF-' + Utilities.formatDate(new Date(), OFF_CFG.TZ, 'yyyyMMdd-HHmmss');
}

/**
 * Normaliza texto para comparacoes de headers e enums.
 */
function off_normalizeTextKey_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase();
}

/**
 * Normaliza header usando o core quando possivel.
 */
function off_normalizeHeader_(header) {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreNormalizeHeader === 'function') {
    return GEAPA_CORE.coreNormalizeHeader(String(header || '').trim());
  }
  return off_normalizeTextKey_(header);
}

/**
 * Faz logging com fallback para console.
 */
function off_logInfo_(runId, message, details) {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreLogInfo === 'function') {
    GEAPA_CORE.coreLogInfo(runId, message, details || null);
    return;
  }
  console.log(message, details || '');
}

/**
 * Faz logging de alerta com fallback para console.
 */
function off_logWarn_(runId, message, details) {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreLogWarn === 'function') {
    GEAPA_CORE.coreLogWarn(runId, message, details || null);
    return;
  }
  console.warn(message, details || '');
}

/**
 * Faz logging de erro com fallback para console.
 */
function off_logError_(runId, message, details) {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreLogError === 'function') {
    GEAPA_CORE.coreLogError(runId, message, details || null);
    return;
  }
  console.error(message, details || '');
}

/**
 * Retorna o registry institucional quando disponivel.
 */
function off_getRegistry_() {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreGetRegistry === 'function') {
    return GEAPA_CORE.coreGetRegistry() || {};
  }
  return {};
}

/**
 * Retorna uma referencia especifica do registry.
 */
function off_getRegistryRef_(key, required) {
  const reg = off_getRegistry_();
  const ref = reg && reg[key];
  if (ref) {
    return {
      id: ref.id || '',
      sheet: ref.sheet || ref.sheetName || '',
      sheetName: ref.sheetName || ref.sheet || '',
      raw: ref,
    };
  }
  if (required) {
    throw new Error('Registry key nao encontrada: ' + key);
  }
  return null;
}

/**
 * Abre uma aba por key do registry.
 */
function off_getSheetByKey_(key, required) {
  if (typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE.coreGetSheetByKey === 'function') {
    try {
      const sheet = GEAPA_CORE.coreGetSheetByKey(key);
      if (sheet) {
        return sheet;
      }
    } catch (err) {
      if (required) {
        throw err;
      }
    }
  }

  const ref = off_getRegistryRef_(key, false);
  if (ref && ref.id && ref.sheetName) {
    return SpreadsheetApp.openById(ref.id).getSheetByName(ref.sheetName);
  }

  if (required) {
    throw new Error('Registry key nao encontrada: ' + key);
  }
  return null;
}

/**
 * Abre a aba bruta de respostas do formulario.
 */
function off_getResponsesSheet_() {
  return off_getSheetByKey_(OFF_KEYS.RESPONSES, true);
}

/**
 * Abre a base de membros quando configurada.
 */
function off_getMembersSheet_() {
  return off_getSheetByKey_(OFF_KEYS.MEMBERS, false);
}

/**
 * Abre planilhas opcionais de atividades.
 */
function off_getOptionalSheetByKey_(key) {
  return off_getSheetByKey_(key, false);
}

/**
 * Retorna a spreadsheet onde ficam a aba bruta e as filas operacionais.
 */
function off_getResponseSpreadsheet_() {
  const sh = off_getResponsesSheet_();
  return sh ? sh.getParent() : null;
}

/**
 * Abre uma spreadsheet por id.
 */
function off_openSpreadsheetById_(spreadsheetId) {
  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * Monta o indice de cabecalhos.
 */
function off_buildHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach(function(header, idx) {
    map[off_normalizeHeader_(header)] = idx + 1;
  });
  return map;
}

/**
 * Busca coluna por lista de aliases.
 */
function off_findColumn_(headerMap, candidates) {
  const list = Array.isArray(candidates) ? candidates : [candidates];
  for (let i = 0; i < list.length; i += 1) {
    const col = headerMap[off_normalizeHeader_(list[i])];
    if (col) {
      return col;
    }
  }
  return 0;
}

/**
 * Leitura completa de uma aba por cabecalho.
 */
function off_readSheetTable_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastCol) {
    return { sheet: sheet, headers: [], headerMap: {}, rows: [], startRow: OFF_CFG.HEADER_ROW + 1 };
  }

  const headers = sheet
    .getRange(OFF_CFG.HEADER_ROW, 1, 1, lastCol)
    .getValues()[0]
    .map(function(value) { return String(value || '').trim(); });

  const rows = lastRow > OFF_CFG.HEADER_ROW
    ? sheet.getRange(OFF_CFG.HEADER_ROW + 1, 1, lastRow - OFF_CFG.HEADER_ROW, lastCol).getValues()
    : [];

  return {
    sheet: sheet,
    headers: headers,
    headerMap: off_buildHeaderMap_(headers),
    rows: rows,
    startRow: OFF_CFG.HEADER_ROW + 1,
  };
}

/**
 * Leitura de uma linha com acesso por aliases.
 */
function off_readRowContext_(sheet, rowNumber) {
  const table = off_readSheetTable_(sheet);
  const row = sheet.getRange(rowNumber, 1, 1, table.headers.length).getValues()[0];
  return {
    sheet: sheet,
    rowNumber: rowNumber,
    headers: table.headers,
    headerMap: table.headerMap,
    row: row,
  };
}

/**
 * Le valor da linha por aliases.
 */
function off_getRowValue_(rowCtx, candidates, defaultValue) {
  const col = off_findColumn_(rowCtx.headerMap, candidates);
  if (!col) {
    return defaultValue;
  }
  const value = rowCtx.row[col - 1];
  return value === undefined ? defaultValue : value;
}

/**
 * Atualiza celulas de uma linha usando o nome do cabecalho.
 */
function off_setRowValues_(sheet, rowNumber, updates) {
  const table = off_readSheetTable_(sheet);
  Object.keys(updates || {}).forEach(function(header) {
    const col = off_findColumn_(table.headerMap, header);
    if (col) {
      sheet.getRange(rowNumber, col).setValue(updates[header]);
    }
  });
}

/**
 * Garante cabecalhos esperados na aba operacional sem apagar colunas extras.
 */
function off_ensureHeaders_(sheet, expectedHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  const currentHeaders = lastCol
    ? sheet.getRange(OFF_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(function(value) {
        return String(value || '').trim();
      })
    : [];

  if (!currentHeaders.filter(Boolean).length) {
    sheet.getRange(OFF_CFG.HEADER_ROW, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const headerMap = off_buildHeaderMap_(currentHeaders);
  const missing = expectedHeaders.filter(function(header) {
    return !off_findColumn_(headerMap, header);
  });

  if (!missing.length) {
    return;
  }

  const startCol = currentHeaders.filter(Boolean).length + 1;
  sheet.getRange(OFF_CFG.HEADER_ROW, startCol, 1, missing.length).setValues([missing]);
  sheet.setFrozenRows(1);
}

/**
 * Garante a existencia de uma aba operacional com os cabecalhos configurados.
 */
function off_ensureSheetByRegistryKey_(key, expectedHeaders) {
  const ref = off_getRegistryRef_(key, true);
  if (!ref.id || !ref.sheetName) {
    throw new Error('Registry ' + key + ' precisa informar id e sheet/sheetName.');
  }

  const ss = off_openSpreadsheetById_(ref.id);
  let sheet = ss.getSheetByName(ref.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(ref.sheetName);
  }
  off_ensureHeaders_(sheet, expectedHeaders);
  return sheet;
}

/**
 * Garante as duas filas operacionais e retorna as referencias.
 */
function off_ensureOperationalSheets_() {
  return {
    suspensions: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_SUSPENSIONS, OFF_CFG.SHEET_HEADERS.SUSPENSION),
    dismissals: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_DISMISSALS, OFF_CFG.SHEET_HEADERS.DISMISSAL),
  };
}

/**
 * Seleciona a fila operacional a partir do tipo de solicitacao.
 */
function off_getQueueSheetByType_(requestType) {
  const sheets = off_ensureOperationalSheets_();
  return requestType === OFF_TYPES.SUSPENSAO ? sheets.suspensions : sheets.dismissals;
}

/**
 * Verifica se uma sheet pertence a uma das filas operacionais configuradas no registry.
 */
function off_isOperationalQueueSheet_(sheet) {
  const queues = off_ensureOperationalSheets_();
  const queueIds = [
    queues.suspensions.getSheetId(),
    queues.dismissals.getSheetId(),
  ];
  return queueIds.indexOf(sheet.getSheetId()) >= 0;
}

/**
 * Retorna o conjunto de headers da fila conforme o tipo.
 */
function off_getQueueHeadersByType_(requestType) {
  return requestType === OFF_TYPES.SUSPENSAO
    ? OFF_CFG.SHEET_HEADERS.SUSPENSION
    : OFF_CFG.SHEET_HEADERS.DISMISSAL;
}

/**
 * Faz append de um objeto na fila mantendo a ordem dos cabecalhos.
 */
function off_appendObjectRow_(sheet, headers, data) {
  const row = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(data, header) ? data[header] : '';
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

/**
 * Busca uma solicitacao ja roteada a partir da linha de origem do formulario.
 */
function off_findExistingRequestBySourceRow_(sourceRow) {
  const sheets = off_ensureOperationalSheets_();
  const sheetList = [sheets.suspensions, sheets.dismissals];

  for (let i = 0; i < sheetList.length; i += 1) {
    const table = off_readSheetTable_(sheetList[i]);
    const sourceCol = off_findColumn_(table.headerMap, OFF_CFG.HEADERS.ORIGEM_LINHA_FORM);
    if (!sourceCol) {
      continue;
    }

    for (let rowIdx = 0; rowIdx < table.rows.length; rowIdx += 1) {
      if (String(table.rows[rowIdx][sourceCol - 1] || '') === String(sourceRow)) {
        return {
          sheet: sheetList[i],
          rowNumber: table.startRow + rowIdx,
        };
      }
    }
  }

  return null;
}

/**
 * Converte qualquer valor em string SIM/NAO/NAO_VERIFICADO.
 */
function off_toFlag_(value) {
  if (value === true) {
    return OFF_CFG.VALUES.YES;
  }
  if (value === false) {
    return OFF_CFG.VALUES.NO;
  }
  return OFF_CFG.VALUES.NAO_VERIFICADO;
}

/**
 * Faz parse de data com tolerancia para valores vindos de planilha.
 */
function off_parseDate_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Retorna apenas a data no timezone do modulo.
 */
function off_startOfDay_(value) {
  const date = off_parseDate_(value);
  if (!date) {
    return null;
  }
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Verifica se uma data venceu ate o dia atual.
 */
function off_isDueOnOrBeforeToday_(value, referenceDate) {
  const dueDate = off_startOfDay_(value);
  const today = off_startOfDay_(referenceDate || new Date());
  if (!dueDate || !today) {
    return false;
  }
  return dueDate.getTime() <= today.getTime();
}

/**
 * Junta mensagens evitando duplicacao literal.
 */
function off_joinMessage_(baseMessage, extraMessage) {
  const base = String(baseMessage || '').trim();
  const extra = String(extraMessage || '').trim();
  if (!extra) {
    return base;
  }
  if (!base) {
    return extra;
  }
  if (base.indexOf(extra) >= 0) {
    return base;
  }
  return base + ' | ' + extra;
}

/**
 * Gera ID deterministicamente a partir da linha do formulario.
 */
function off_generateRequestId_(requestType, executionMode, sourceRow, requestDate) {
  const datePart = Utilities.formatDate(requestDate || new Date(), OFF_CFG.TZ, 'yyyyMMdd');
  return [
    'OFF',
    requestType,
    executionMode,
    datePart,
    'R' + sourceRow,
  ].join('-');
}
