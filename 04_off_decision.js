/***************************************
 * 04_off_decision.gs
 * Reage a edicoes de decisao nas filas operacionais.
 ***************************************/

/**
 * Trigger onEdit da spreadsheet: aplica a decisao da diretoria em qualquer fila operacional.
 */
function off_onEditDecision(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (!off_isOperationalQueueSheet_(sheet)) {
    return;
  }
  if (e.range.getRow() <= OFF_CFG.HEADER_ROW) {
    return;
  }

  const table = off_readSheetTable_(sheet);
  const decisionCol = off_findColumn_(table.headerMap, OFF_CFG.HEADERS.DECISAO_DIRETORIA);
  if (!decisionCol || e.range.getColumn() !== decisionCol) {
    return;
  }

  const decision = off_normalizeTextKey_(e.value || '');
  if (decision !== OFF_CFG.VALUES.DEFERIDO && decision !== OFF_CFG.VALUES.INDEFERIDO) {
    return;
  }

  const runId = off_newRunId_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    off_logWarn_(runId, 'off_onEditDecision: lock nao obtido', {
      sheetName: sheet.getName(),
      row: e.range.getRow(),
    });
    return;
  }

  try {
    off_processDecisionRow_(sheet, e.range.getRow(), runId);
  } catch (err) {
    off_logError_(runId, 'off_onEditDecision: erro', {
      err: String(err),
      stack: err && err.stack,
      sheetName: sheet.getName(),
      row: e.range.getRow(),
    });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processa uma linha apos decisao manual, preservando execucao imediata de desligamento.
 */
function off_processDecisionRow_(sheet, rowNumber, runId) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const decision = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DECISAO_DIRETORIA, ''));
  const requestType = String(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.TIPO_SOLICITACAO, '') || '').trim();
  const executionMode = String(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MODALIDADE_EXECUCAO, '') || '').trim();
  const currentStatus = String(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.STATUS_SOLICITACAO, '') || '').trim();

  if (currentStatus === OFF_STATUS.EXECUTADO && decision === OFF_CFG.VALUES.DEFERIDO) {
    off_logInfo_(runId, 'off_processDecisionRow_: linha ja executada, ignorando reprocessamento', {
      sheetName: sheet.getName(),
      rowNumber: rowNumber,
    });
    return;
  }

  const now = new Date();
  off_setRowValues_(sheet, rowNumber, {
    DATA_DECISAO: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_DECISAO, '') || now,
    STATUS_SOLICITACAO: decision === OFF_CFG.VALUES.INDEFERIDO ? OFF_STATUS.INDEFERIDO : OFF_STATUS.DEFERIDO,
    RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    ERRO_EXECUCAO: '',
  });

  if (decision === OFF_CFG.VALUES.INDEFERIDO) {
    off_maybeSendDecisionEmail_(sheet, rowNumber, false);
    off_setRowValues_(sheet, rowNumber, {
      MENSAGEM_PROCESSAMENTO: off_joinMessage_(
        off_getRowValue_(off_readRowContext_(sheet, rowNumber), OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
        'Solicitacao indeferida pela diretoria.'
      ),
    });
    return;
  }

  if (requestType === OFF_TYPES.SUSPENSAO) {
    off_maybeSendDecisionEmail_(sheet, rowNumber, true);
    off_handleApprovedSuspensionDecision_(sheet, rowNumber, runId);
    return;
  }

  if (requestType === OFF_TYPES.DESLIGAMENTO && executionMode === OFF_EXECUTION.IMEDIATO) {
    off_executeDismissalRow_(sheet, rowNumber, runId, {
      effectiveDate: now,
      origin: 'DECISION',
    });
    return;
  }

  if (requestType === OFF_TYPES.DESLIGAMENTO && executionMode === OFF_EXECUTION.FIM_SEMESTRE) {
    off_maybeSendDecisionEmail_(sheet, rowNumber, true);
    off_scheduleApprovedDismissal_(sheet, rowNumber, runId);
  }
}

/**
 * Trata deferimento de suspensao, aplicando agora ou agendando para o job diario.
 */
function off_handleApprovedSuspensionDecision_(sheet, rowNumber, runId) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const startDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_INICIO_SUSPENSAO, null));
  const endDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_FIM_SUSPENSAO, null));
  const days = off_calculateSuspensionDays_(startDate, endDate);

  if (!startDate || !endDate || days === null || days < OFF_CFG.MIN_SUSPENSION_DAYS) {
    off_setRowValues_(sheet, rowNumber, {
      STATUS_SOLICITACAO: OFF_STATUS.ERRO,
      ERRO_EXECUCAO: 'Suspensao deferida sem periodo valido de no minimo ' + OFF_CFG.MIN_SUSPENSION_DAYS + ' dias.',
      RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    });
    return;
  }

  if (off_isDueOnOrBeforeToday_(startDate, new Date())) {
    off_executeSuspensionRow_(sheet, rowNumber, runId, { origin: 'DECISION' });
    return;
  }

  off_setRowValues_(sheet, rowNumber, {
    STATUS_SOLICITACAO: OFF_STATUS.AGUARDANDO_EXECUCAO,
    MENSAGEM_PROCESSAMENTO: off_joinMessage_(
      off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
      'Suspensao deferida e aguardando chegada da data de inicio.'
    ),
    RUN_ID_ULTIMO_PROCESSAMENTO: runId,
  });
}

/**
 * Agenda desligamento deferido para execucao no fim do semestre.
 */
function off_scheduleApprovedDismissal_(sheet, rowNumber, runId) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const currentDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_EFETIVA_DESLIGAMENTO, null));
  const semesterReference = String(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.SEMESTRE_REFERENCIA, '') || '').trim();
  const effectiveDate = currentDate || off_computeSemesterEndDate_(semesterReference, new Date());
  const membersResult = off_tryMarkScheduledDismissalInMembers_({
    requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
    rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
    email: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
    effectiveDate: effectiveDate,
    semesterReference: semesterReference,
  });

  off_setRowValues_(sheet, rowNumber, {
    DATA_EFETIVA_DESLIGAMENTO: effectiveDate,
    STATUS_SOLICITACAO: OFF_STATUS.AGUARDANDO_EXECUCAO,
    MENSAGEM_INTEGRACAO_MEMBROS: membersResult.message,
    MENSAGEM_PROCESSAMENTO: off_joinMessage_(
      off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
      'Desligamento deferido e agendado para a data efetiva.'
    ),
    RUN_ID_ULTIMO_PROCESSAMENTO: runId,
  });
}

/**
 * Envia e marca o e-mail de decisao apenas uma vez por linha.
 */
function off_maybeSendDecisionEmail_(sheet, rowNumber, approved) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const alreadySent = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL_DECISAO_ENVIADO, ''));
  if (alreadySent === OFF_CFG.VALUES.YES) {
    return;
  }

  const sent = off_sendDecisionEmail_({
    name: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
    email: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
    requestType: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.TIPO_SOLICITACAO, ''),
    executionMode: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MODALIDADE_EXECUCAO, ''),
    approved: approved,
  });

  if (!sent.sent) {
    return;
  }

  off_setRowValues_(sheet, rowNumber, {
    EMAIL_DECISAO_ENVIADO: OFF_CFG.VALUES.YES,
    DATA_EMAIL_DECISAO: new Date(),
  });
}
