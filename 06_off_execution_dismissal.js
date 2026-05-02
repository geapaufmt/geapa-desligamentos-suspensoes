/***************************************
 * 06_off_execution_dismissal.gs
 * Execucao de desligamentos imediatos e agendados.
 ***************************************/

/**
 * Executa o desligamento voluntario, preservando o contrato atual de integracao imediata com membros.
 */
function off_executeDismissalRow_(sheet, rowNumber, runId, options) {
  options = options || {};
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.EXECUCAO_DESLIGAMENTO,
    OFF_OPS.CAPABILITIES.SYNC,
    {
      executionType: options.executionType || OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: runId,
      origin: options.origin || 'MANUAL',
      sheetName: sheet.getName(),
      rowNumber: rowNumber,
    },
    function(control) {
      const rowCtx = off_readRowContext_(sheet, rowNumber);
      const alreadyExecuted = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DESLIGAMENTO_EXECUTADO, ''));
      if (alreadyExecuted === OFF_CFG.VALUES.YES) {
        return { skipped: true, message: 'Desligamento ja executado anteriormente.' };
      }

      if (control.dryRun) {
        off_logInfo_(runId, 'off_executeDismissalRow_: DRY_RUN, desligamento validado sem efeitos reais.', {
          sheetName: sheet.getName(),
          rowNumber: rowNumber,
        });
        return { dryRun: true, message: 'DRY_RUN: desligamento nao executado.' };
      }

      try {
        const emailResult = off_ensureFinalDismissalEmail_(sheet, rowNumber, {
          runId: runId,
          executionType: control.executionType,
          origin: options.origin || 'MANUAL',
        });
        if (emailResult && (emailResult.blocked || emailResult.dryRun)) {
          return off_buildBlockedResult_(control, emailResult.message);
        }
        if (emailResult && !emailResult.sent) {
          throw new Error(emailResult.message || 'Falha ao enviar e-mail final de desligamento.');
        }

        const integrationResult = off_processMandatoryImmediateDismissalInMembers_(
          off_buildMembersPayload_(sheet, rowNumber),
          {
            runId: runId,
            executionType: control.executionType,
            origin: options.origin || 'MANUAL',
            sheetName: sheet.getName(),
            rowNumber: rowNumber,
          }
        );
        if (integrationResult && (integrationResult.blocked || integrationResult.dryRun)) {
          return off_buildBlockedResult_(control, integrationResult.message);
        }

        const event = off_appendDismissalEvent_({
          sourceKey: sheet.getName(),
          sourceRow: rowNumber,
          requestId: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.ID_SOLICITACAO, ''),
          rga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
          eventDate: options.effectiveDate ? options.effectiveDate : new Date(),
          reason: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MOTIVO, ''),
          memberName: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
          memberEmail: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
          notes: 'Desligamento voluntario homologado.',
          runId: runId,
          executionType: control.executionType,
          origin: options.origin || 'MANUAL',
        });
        if (event && (event.blocked || event.dryRun)) {
          return off_buildBlockedResult_(control, event.message);
        }

        off_setRowValues_(sheet, rowNumber, {
          STATUS_SOLICITACAO: OFF_STATUS.EXECUTADO,
          DATA_EFETIVA_DESLIGAMENTO: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_EFETIVA_DESLIGAMENTO, '') || new Date(),
          INTEGRACAO_MEMBROS_PROCESSADA: OFF_CFG.VALUES.YES,
          DATA_INTEGRACAO_MEMBROS: integrationResult.processedAt || new Date(),
          MENSAGEM_INTEGRACAO_MEMBROS: integrationResult.message,
          DESLIGAMENTO_EXECUTADO: OFF_CFG.VALUES.YES,
          DATA_EXECUCAO: new Date(),
          ID_EVENTO_DESLIGAMENTO: event.eventId || '',
          MENSAGEM_PROCESSAMENTO: off_joinMessage_(
            off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MENSAGEM_PROCESSAMENTO, ''),
            'Desligamento executado com integracao em membros e evento institucional registrado.'
          ),
          ERRO_EXECUCAO: '',
          RUN_ID_ULTIMO_PROCESSAMENTO: runId,
        });

        return { integrationResult: integrationResult, event: event };
      } catch (err) {
        off_setRowValues_(sheet, rowNumber, {
          STATUS_SOLICITACAO: OFF_STATUS.ERRO,
          INTEGRACAO_MEMBROS_PROCESSADA: OFF_CFG.VALUES.NO,
          ERRO_EXECUCAO: String(err && err.message ? err.message : err),
          RUN_ID_ULTIMO_PROCESSAMENTO: runId,
        });
        throw err;
      }
    }
  );
}

/**
 * Garante o e-mail final antes da integracao obrigatoria de desligamento.
 */
function off_ensureFinalDismissalEmail_(sheet, rowNumber, options) {
  options = options || {};
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  const alreadySent = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL_FINAL_ENVIADO, ''));
  if (alreadySent === OFF_CFG.VALUES.YES) {
    return { sent: true, message: 'E-mail final ja havia sido enviado.' };
  }

  const sent = off_sendFinalDismissalEmail_({
    name: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
    email: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
    requestType: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.TIPO_SOLICITACAO, ''),
  }, options);

  if (!sent.sent) {
    return sent;
  }

  off_setRowValues_(sheet, rowNumber, {
    EMAIL_FINAL_ENVIADO: OFF_CFG.VALUES.YES,
    DATA_EMAIL_FINAL: new Date(),
    EMAIL_DECISAO_ENVIADO: OFF_CFG.VALUES.YES,
    DATA_EMAIL_DECISAO: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_EMAIL_DECISAO, '') || new Date(),
  });

  return sent;
}

/**
 * Monta o payload legado esperado pela library de membros para desligamento imediato aprovado.
 */
function off_buildMembersPayload_(sheet, rowNumber) {
  const rowCtx = off_readRowContext_(sheet, rowNumber);
  return {
    requestRowNumber: rowNumber,
    requestTimestamp: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_PEDIDO, ''),
    memberName: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.NOME, ''),
    memberEmail: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.EMAIL, ''),
    memberRga: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.RGA, ''),
    requestType: 'Desligamento',
    leaveTiming: 'Imediatamente',
    reason: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MOTIVO, ''),
    approvedAt: off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_DECISAO, new Date()),
    decision: 'DEFERIDO',
    finalEmailSent: OFF_CFG.VALUES.YES,
    sourceKey: sheet.getName(),
  };
}

/**
 * Processa desligamentos agendados cuja data efetiva ja chegou.
 */
function off_processDueScheduledDismissals_(runId, options) {
  options = options || {};
  const sheet = off_getQueueSheetByType_(OFF_TYPES.DESLIGAMENTO);
  const table = off_readSheetTable_(sheet);
  let processed = 0;

  for (let i = 0; i < table.rows.length; i += 1) {
    const rowNumber = table.startRow + i;
    const rowCtx = off_readRowContext_(sheet, rowNumber);
    const decision = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DECISAO_DIRETORIA, ''));
    const executionMode = String(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.MODALIDADE_EXECUCAO, '') || '').trim();
    const executed = off_normalizeTextKey_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DESLIGAMENTO_EXECUTADO, ''));
    const dueDate = off_parseDate_(off_getRowValue_(rowCtx, OFF_CFG.HEADERS.DATA_EFETIVA_DESLIGAMENTO, null));

    if (decision !== OFF_CFG.VALUES.DEFERIDO) {
      continue;
    }
    if (executionMode !== OFF_EXECUTION.FIM_SEMESTRE) {
      continue;
    }
    if (executed === OFF_CFG.VALUES.YES) {
      continue;
    }
    if (!off_isDueOnOrBeforeToday_(dueDate, new Date())) {
      continue;
    }

    const result = off_executeDismissalRow_(sheet, rowNumber, runId, {
      effectiveDate: dueDate || new Date(),
      origin: 'JOB',
      executionType: options.executionType || OFF_OPS.EXECUTION_TYPES.MANUAL,
    });
    if (result && !result.blocked && !result.dryRun && !result.skipped) {
      processed += 1;
    }
  }

  return processed;
}
