/***************************************
 * 05_off_execution_suspension.gs
 * Aplicacao de suspensoes e retornos.
 ***************************************/

/**
 * Executa a suspensao homologada, sempre registrando primeiro o evento institucional obrigatorio.
 */
function off_executeSuspensionRow_(sheet, rowNumber, runId, options) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const alreadyApplied = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.SUSPENSAO_APLICADA, ''));
  if (alreadyApplied === OFF_CFG.VALUES.YES) {
    return { skipped: true, message: 'Suspensao ja aplicada anteriormente.' };
  }

  try {
    const event = off_appendSuspensionEvent_({
      sourceKey: sheet.getName(),
      sourceRow: rowNumber,
      requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
      rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
      eventDate: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_INICIO_SUSPENSAO, new Date()),
      reason: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MOTIVO, ''),
      memberName: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
      memberEmail: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
      notes: 'Suspensao voluntaria homologada.',
    });
    const membersResult = off_tryApplySuspensionInMembers_({
      requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
      rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
      email: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
      startDate: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_INICIO_SUSPENSAO, ''),
      endDate: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_FIM_SUSPENSAO, ''),
    });

    off_setRowValues_(sheet, rowNumber, {
      STATUS_SOLICITACAO: OFF_STATUS.EXECUTADO,
      SUSPENSAO_APLICADA: OFF_CFG.VALUES.YES,
      DATA_EXECUCAO: new Date(),
      ID_EVENTO_SUSPENSAO: event.eventId || '',
      MENSAGEM_PROCESSAMENTO: off_joinMessage_(
        off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
        'Suspensao aplicada. ' + membersResult.message
      ),
      ERRO_EXECUCAO: '',
      RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    });

    return { event: event, membersResult: membersResult };
  } catch (err) {
    off_setRowValues_(sheet, rowNumber, {
      STATUS_SOLICITACAO: OFF_STATUS.ERRO,
      ERRO_EXECUCAO: String(err && err.message ? err.message : err),
      RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    });
    throw err;
  }
}

/**
 * Executa o retorno de uma suspensao ja aplicada quando a data final chega.
 */
function off_executeSuspensionReturnRow_(sheet, rowNumber, runId) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const alreadyReturned = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RETORNO_PROCESSADO, ''));
  if (alreadyReturned === OFF_CFG.VALUES.YES) {
    return { skipped: true, message: 'Retorno ja processado anteriormente.' };
  }

  try {
    const event = off_appendReturnEvent_({
      sourceKey: sheet.getName(),
      sourceRow: rowNumber,
      requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
      rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
      eventDate: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_FIM_SUSPENSAO, new Date()),
      reason: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MOTIVO, ''),
      memberName: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
      memberEmail: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
      notes: 'Retorno automatico ao final da suspensao voluntaria.',
    });
    const membersResult = off_tryFinishSuspensionInMembers_({
      requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
      rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
      email: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
      returnDate: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_FIM_SUSPENSAO, ''),
    });

    off_setRowValues_(sheet, rowNumber, {
      RETORNO_PROCESSADO: OFF_CFG.VALUES.YES,
      DATA_RETORNO_PROCESSADO: new Date(),
      ID_EVENTO_RETORNO: event.eventId || '',
      MENSAGEM_PROCESSAMENTO: off_joinMessage_(
        off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
        'Retorno processado. ' + membersResult.message
      ),
      ERRO_EXECUCAO: '',
      RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    });

    return { event: event, membersResult: membersResult };
  } catch (err) {
    off_setRowValues_(sheet, rowNumber, {
      STATUS_SOLICITACAO: OFF_STATUS.ERRO,
      ERRO_EXECUCAO: String(err && err.message ? err.message : err),
      RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    });
    throw err;
  }
}

/**
 * Processa todas as suspensoes deferidas cuja data de inicio ja chegou.
 */
function off_processDueSuspensionStarts_(runId) {
  const sheet = off_getQueueSheetByType_(OFF_TYPES.SUSPENSAO);
  const table = off_readSheetTable_(sheet);
  let processed = 0;

  for (let i = 0; i < table.rows.length; i += 1) {
    const rowNumber = table.startRow + i;
    const rowCtx = off_readRowContext_(sheet, rowNumber);
    const decision = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DECISAO_DIRETORIA, ''));
    const applied = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.SUSPENSAO_APLICADA, ''));
    const startDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_INICIO_SUSPENSAO, null));

    if (decision !== OFF_CFG.VALUES.DEFERIDO || applied === OFF_CFG.VALUES.YES) {
      continue;
    }
    if (!off_isDueOnOrBeforeToday_(startDate, new Date())) {
      continue;
    }

    off_executeSuspensionRow_(sheet, rowNumber, runId, { origin: 'JOB' });
    processed += 1;
  }

  return processed;
}

/**
 * Processa retornos de suspensao cuja data final ja chegou.
 */
function off_processDueSuspensionReturns_(runId) {
  const sheet = off_getQueueSheetByType_(OFF_TYPES.SUSPENSAO);
  const table = off_readSheetTable_(sheet);
  let processed = 0;

  for (let i = 0; i < table.rows.length; i += 1) {
    const rowNumber = table.startRow + i;
    const rowCtx = off_readRowContext_(sheet, rowNumber);
    const applied = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.SUSPENSAO_APLICADA, ''));
    const returned = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RETORNO_PROCESSADO, ''));
    const endDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_FIM_SUSPENSAO, null));

    if (applied !== OFF_CFG.VALUES.YES || returned === OFF_CFG.VALUES.YES) {
      continue;
    }
    if (!off_isDueOnOrBeforeToday_(endDate, new Date())) {
      continue;
    }

    off_executeSuspensionReturnRow_(sheet, rowNumber, runId);
    processed += 1;
  }

  return processed;
}
