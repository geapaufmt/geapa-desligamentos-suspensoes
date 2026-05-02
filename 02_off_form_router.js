/***************************************
 * 02_off_form_router.gs
 * Entrada unica via formulario e roteamento para filas.
 ***************************************/

/**
 * Trigger de formulario: normaliza a submissao, evita duplicidade e roteia para a fila correta.
 */
function off_handleFormSubmit(e) {
  const runId = off_newRunId_();
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.RECEBIMENTO_FORMULARIO,
    OFF_OPS.CAPABILITIES.SYNC,
    {
      executionType: off_getExecutionTypeFromEvent_(e),
      runId: runId,
      origin: 'FORM_SUBMIT',
    },
    function(control) {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) {
        off_logWarn_(runId, 'off_handleFormSubmit: lock nao obtido');
        return null;
      }

      try {
        off_ensureOperationalSheets_();

        const responsesSheet = off_getResponsesSheet_();
        const sourceRow = off_extractSourceRow_(e, responsesSheet);
        if (!sourceRow || sourceRow <= OFF_CFG.HEADER_ROW) {
          off_logWarn_(runId, 'off_handleFormSubmit: linha de origem invalida', { sourceRow: sourceRow });
          return null;
        }

        const duplicated = off_findExistingRequestBySourceRow_(sourceRow);
        if (duplicated) {
          off_logInfo_(runId, 'off_handleFormSubmit: submissao ja roteada', duplicated);
          return duplicated;
        }

        const rawRow = off_readRowContext_(responsesSheet, sourceRow);
        const request = off_normalizeRawSubmission_(rawRow, sourceRow);
        const queueSheet = off_getQueueSheetByType_(request.requestType);
        const queueHeaders = off_getQueueHeadersByType_(request.requestType);
        const queuePayload = request.requestType === OFF_TYPES.SUSPENSAO
          ? off_buildSuspensionQueueRow_(request, runId)
          : off_buildDismissalQueueRow_(request, runId);

        if (control.dryRun) {
          off_logInfo_(runId, 'off_handleFormSubmit: DRY_RUN, roteamento validado sem escrever na fila oficial.', {
            requestId: request.requestId,
            queueSheet: queueSheet.getName(),
            requestType: request.requestType,
            executionMode: request.executionMode,
          });
          return {
            dryRun: true,
            requestId: request.requestId,
            queueSheet: queueSheet.getName(),
            queuePayload: queuePayload,
          };
        }

        const queueRow = off_appendObjectRow_(queueSheet, queueHeaders, queuePayload);
        const notification = off_notifySecretariesOfNewRequest_(Object.assign({}, request, {
          runId: runId,
          executionType: control.executionType,
        }));

        off_setRowValues_(queueSheet, queueRow, {
          NOTIFICACAO_SECRETARIA_ENVIADA: notification.sent ? OFF_CFG.VALUES.YES : OFF_CFG.VALUES.NO,
          DATA_NOTIFICACAO_SECRETARIA: notification.sent ? new Date() : '',
          MENSAGEM_PROCESSAMENTO: off_joinMessage_(queuePayload.MENSAGEM_PROCESSAMENTO, notification.message),
        });

        off_logInfo_(runId, 'off_handleFormSubmit: solicitacao roteada', {
          requestId: request.requestId,
          queueSheet: queueSheet.getName(),
          queueRow: queueRow,
          requestType: request.requestType,
          executionMode: request.executionMode,
        });

        return {
          sheet: queueSheet,
          rowNumber: queueRow,
          requestId: request.requestId,
        };
      } catch (err) {
        off_logError_(runId, 'off_handleFormSubmit: erro', {
          err: String(err),
          stack: err && err.stack,
        });
        throw err;
      } finally {
        lock.releaseLock();
      }
    }
  );
}

/**
 * Descobre a linha de origem do evento com fallback para ultima linha.
 */
function off_extractSourceRow_(e, responsesSheet) {
  if (e && e.range && e.range.getSheet().getSheetId() === responsesSheet.getSheetId()) {
    return e.range.getRow();
  }
  return responsesSheet.getLastRow();
}

/**
 * Normaliza uma linha da aba bruta para o contrato interno do modulo.
 */
function off_normalizeRawSubmission_(rowCtx, sourceRow) {
  const requestDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.TIMESTAMP, null)) || new Date();
  const requestType = off_normalizeRequestType_(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.REQUEST_TYPE, ''));
  if (!requestType) {
    throw new Error('Tipo de solicitacao nao reconhecido na linha ' + sourceRow + '.');
  }

  const executionMode = off_normalizeExecutionMode_(
    requestType,
    off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.DISMISSAL_TIMING, '')
  );

  const name = String(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.NAME, '') || '').trim();
  const rga = String(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.RGA, '') || '').trim();
  const email = String(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.EMAIL, '') || '').trim();
  const reason = String(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.REASON, '') || '').trim();
  const observations = String(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.OBSERVATIONS, '') || '').trim();
  const suspensionStart = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.SUSPENSION_START, null));
  const suspensionEnd = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.SUSPENSION_END, null));
  const requestedEffectiveDismissalDate = off_parseDate_(
    off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.DISMISSAL_EFFECTIVE_DATE, null)
  );
  const semesterReferenceInput = String(
    off_getRowValue_(rowCtx, OFF_CFG.RAW_FIELDS.SEMESTER_REFERENCE, '') || ''
  ).trim();
  const semesterReference = semesterReferenceInput || off_deriveSemesterReference_(requestDate);
  const effectiveDismissalDate = requestType === OFF_TYPES.DESLIGAMENTO
    ? off_resolveDismissalEffectiveDate_(executionMode, requestedEffectiveDismissalDate, semesterReference, requestDate)
    : null;
  const suspensionDays = requestType === OFF_TYPES.SUSPENSAO
    ? off_calculateSuspensionDays_(suspensionStart, suspensionEnd)
    : null;
  const requestId = off_generateRequestId_(requestType, executionMode, sourceRow, requestDate);
  const validation = off_validateNormalizedRequest_({
    requestType: requestType,
    executionMode: executionMode,
    requestDate: requestDate,
    name: name,
    rga: rga,
    email: email,
    reason: reason,
    observations: observations,
    suspensionStart: suspensionStart,
    suspensionEnd: suspensionEnd,
    suspensionDays: suspensionDays,
    effectiveDismissalDate: effectiveDismissalDate,
    semesterReference: semesterReference,
  });

  return {
    requestId: requestId,
    sourceRow: sourceRow,
    requestDate: requestDate,
    requestType: requestType,
    executionMode: executionMode,
    name: name,
    rga: rga,
    email: email,
    reason: reason,
    observations: observations,
    suspensionStart: suspensionStart,
    suspensionEnd: suspensionEnd,
    suspensionDays: suspensionDays,
    effectiveDismissalDate: effectiveDismissalDate,
    semesterReference: semesterReference,
    validation: validation,
  };
}

/**
 * Converte o valor do formulario para o enum interno de tipo.
 */
function off_normalizeRequestType_(rawType) {
  const key = off_normalizeTextKey_(rawType);
  if (key.indexOf('SUSPENSAO') >= 0) {
    return OFF_TYPES.SUSPENSAO;
  }
  if (key.indexOf('DESLIGAMENTO') >= 0) {
    return OFF_TYPES.DESLIGAMENTO;
  }
  return '';
}

/**
 * Converte a modalidade do formulario para o enum interno.
 */
function off_normalizeExecutionMode_(requestType, rawMode) {
  if (requestType === OFF_TYPES.SUSPENSAO) {
    return OFF_EXECUTION.TEMPORARIA;
  }

  const key = off_normalizeTextKey_(rawMode);
  if (key.indexOf('IMEDIAT') >= 0) {
    return OFF_EXECUTION.IMEDIATO;
  }
  if (key.indexOf('FIM_SEMESTRE') >= 0 || key.indexOf('FINAL_DO_SEMESTRE') >= 0) {
    return OFF_EXECUTION.FIM_SEMESTRE;
  }
  return OFF_EXECUTION.IMEDIATO;
}

/**
 * Resolve a data efetiva do desligamento imediato ou agendado.
 */
function off_resolveDismissalEffectiveDate_(executionMode, explicitDate, semesterReference, fallbackDate) {
  if (executionMode === OFF_EXECUTION.IMEDIATO) {
    return explicitDate || fallbackDate || new Date();
  }
  return explicitDate || off_computeSemesterEndDate_(semesterReference, fallbackDate);
}

/**
 * Cria a linha inicial da fila de suspensao.
 */
function off_buildSuspensionQueueRow_(request, runId) {
  return {
    ID_SOLICITACAO: request.requestId,
    ORIGEM_LINHA_FORM: request.sourceRow,
    DATA_PEDIDO: request.requestDate,
    NOME: request.name,
    RGA: request.rga,
    EMAIL: request.email,
    TIPO_SOLICITACAO: request.requestType,
    MODALIDADE_EXECUCAO: request.executionMode,
    MOTIVO: request.reason,
    OBSERVACOES_MEMBRO: request.observations,
    DATA_INICIO_SUSPENSAO: request.suspensionStart || '',
    DATA_FIM_SUSPENSAO: request.suspensionEnd || '',
    QTDE_DIAS_SUSPENSAO: request.suspensionDays || '',
    VALIDACAO_MEMBRO_ENCONTRADO: request.validation.memberFoundFlag,
    VALIDACAO_PERIODO_MINIMO: request.validation.minimumPeriodFlag,
    VALIDACAO_CONFLITO_APRESENTACAO: request.validation.presentationConflictFlag,
    VALIDACAO_PENDENCIA_ARQUIVO: request.validation.pendingFileFlag,
    VALIDACAO_OUTRAS_PENDENCIAS: request.validation.otherPendingFlag,
    VALIDACAO_GERAL: request.validation.generalStatus,
    DECISAO_DIRETORIA: '',
    DATA_DECISAO: '',
    OBS_DECISAO: '',
    STATUS_SOLICITACAO: OFF_STATUS.RECEBIDO,
    EMAIL_DECISAO_ENVIADO: OFF_CFG.VALUES.NO,
    DATA_EMAIL_DECISAO: '',
    SUSPENSAO_APLICADA: OFF_CFG.VALUES.NO,
    DATA_EXECUCAO: '',
    RETORNO_PROCESSADO: OFF_CFG.VALUES.NO,
    DATA_RETORNO_PROCESSADO: '',
    ID_EVENTO_SUSPENSAO: '',
    ID_EVENTO_RETORNO: '',
    MENSAGEM_PROCESSAMENTO: request.validation.message,
    ERRO_EXECUCAO: '',
    RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    NOTIFICACAO_SECRETARIA_ENVIADA: OFF_CFG.VALUES.NO,
    DATA_NOTIFICACAO_SECRETARIA: '',
  };
}

/**
 * Cria a linha inicial da fila de desligamento.
 */
function off_buildDismissalQueueRow_(request, runId) {
  return {
    ID_SOLICITACAO: request.requestId,
    ORIGEM_LINHA_FORM: request.sourceRow,
    DATA_PEDIDO: request.requestDate,
    NOME: request.name,
    RGA: request.rga,
    EMAIL: request.email,
    TIPO_SOLICITACAO: request.requestType,
    MODALIDADE_EXECUCAO: request.executionMode,
    MOTIVO: request.reason,
    OBSERVACOES_MEMBRO: request.observations,
    DATA_EFETIVA_DESLIGAMENTO: request.effectiveDismissalDate || '',
    SEMESTRE_REFERENCIA: request.semesterReference || '',
    VALIDACAO_MEMBRO_ENCONTRADO: request.validation.memberFoundFlag,
    VALIDACAO_CONFLITO_APRESENTACAO: request.validation.presentationConflictFlag,
    VALIDACAO_PENDENCIA_ARQUIVO: request.validation.pendingFileFlag,
    VALIDACAO_OUTRAS_PENDENCIAS: request.validation.otherPendingFlag,
    VALIDACAO_GERAL: request.validation.generalStatus,
    DECISAO_DIRETORIA: '',
    DATA_DECISAO: '',
    OBS_DECISAO: '',
    STATUS_SOLICITACAO: OFF_STATUS.RECEBIDO,
    EMAIL_DECISAO_ENVIADO: OFF_CFG.VALUES.NO,
    DATA_EMAIL_DECISAO: '',
    EMAIL_FINAL_ENVIADO: OFF_CFG.VALUES.NO,
    DATA_EMAIL_FINAL: '',
    INTEGRACAO_MEMBROS_PROCESSADA: '',
    DATA_INTEGRACAO_MEMBROS: '',
    MENSAGEM_INTEGRACAO_MEMBROS: '',
    DESLIGAMENTO_EXECUTADO: OFF_CFG.VALUES.NO,
    DATA_EXECUCAO: '',
    ID_EVENTO_DESLIGAMENTO: '',
    ERRO_EXECUCAO: '',
    RUN_ID_ULTIMO_PROCESSAMENTO: runId,
    NOTIFICACAO_SECRETARIA_ENVIADA: OFF_CFG.VALUES.NO,
    DATA_NOTIFICACAO_SECRETARIA: '',
  };
}
