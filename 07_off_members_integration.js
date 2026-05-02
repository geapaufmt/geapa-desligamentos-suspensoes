/***************************************
 * 07_off_members_integration.gs
 * Wrappers defensivos para integracao com GEAPA_MEMBROS.
 ***************************************/

/**
 * Mantem obrigatoria a integracao ja existente de desligamento imediato homologado.
 */
function off_processMandatoryImmediateDismissalInMembers_(payload, options) {
  options = options || {};
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.INTEGRACAO_MEMBROS,
    OFF_OPS.CAPABILITIES.SYNC,
    {
      executionType: options.executionType || OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: options.runId,
      origin: options.origin || 'MANDATORY_DISMISSAL',
      sheetName: options.sheetName,
      rowNumber: options.rowNumber,
    },
    function(control) {
      if (control.dryRun) {
        return { dryRun: true, ok: false, message: 'DRY_RUN: integracao com membros nao executada.' };
      }

      if (
        typeof GEAPA_MEMBROS === 'undefined' ||
        !GEAPA_MEMBROS ||
        typeof GEAPA_MEMBROS.members_offboardApprovedImmediateExit !== 'function'
      ) {
        throw new Error('GEAPA_MEMBROS.members_offboardApprovedImmediateExit nao esta disponivel.');
      }

      const result = GEAPA_MEMBROS.members_offboardApprovedImmediateExit(payload) || {};
      const message = result.duplicatedHistory
        ? 'Registro ja existia em MEMBERS_HIST; membro removido de MEMBERS_ATUAIS.'
        : 'Membro movido com sucesso de MEMBERS_ATUAIS para MEMBERS_HIST.';

      return {
        ok: true,
        processedAt: new Date(),
        raw: result,
        message: message,
      };
    }
  );
}

/**
 * Tenta aplicar a suspensao na library de membros sem tornar isso fatal nesta etapa.
 */
function off_tryApplySuspensionInMembers_(payload, options) {
  return off_tryOptionalMembersCall_(
    ['members_applySuspensionEvent', 'members_applySuspension', 'members_applyApprovedSuspension'],
    payload,
    'Integracao de suspensao ainda nao implementada na library GEAPA_MEMBROS.',
    options
  );
}

/**
 * Tenta finalizar a suspensao na library de membros sem tornar isso fatal nesta etapa.
 */
function off_tryFinishSuspensionInMembers_(payload, options) {
  return off_tryOptionalMembersCall_(
    ['members_finishSuspensionEvent', 'members_finishSuspension', 'members_completeSuspension'],
    payload,
    'Integracao de retorno de suspensao ainda nao implementada na library GEAPA_MEMBROS.',
    options
  );
}

/**
 * Tenta registrar o agendamento do desligamento na library quando houver suporte futuro.
 */
function off_tryMarkScheduledDismissalInMembers_(payload, options) {
  return off_tryOptionalMembersCall_(
    ['members_markScheduledDismissal', 'members_scheduleDismissal', 'members_markApprovedScheduledDismissal'],
    payload,
    'Agendamento em GEAPA_MEMBROS ainda nao implementado na library.',
    options
  );
}

/**
 * Executa uma chamada opcional na library de membros sem quebrar o fluxo principal.
 */
function off_tryOptionalMembersCall_(candidateNames, payload, neutralMessage, options) {
  options = options || {};
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.INTEGRACAO_MEMBROS,
    OFF_OPS.CAPABILITIES.SYNC,
    {
      executionType: options.executionType || OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: options.runId,
      origin: options.origin || 'OPTIONAL_MEMBERS',
      sheetName: options.sheetName,
      rowNumber: options.rowNumber,
    },
    function(control) {
      if (control.dryRun) {
        return { dryRun: true, ok: false, implemented: false, message: 'DRY_RUN: integracao opcional com membros nao executada.' };
      }

      if (typeof GEAPA_MEMBROS === 'undefined' || !GEAPA_MEMBROS) {
        return { ok: false, implemented: false, message: neutralMessage };
      }

      for (let i = 0; i < candidateNames.length; i += 1) {
        const fnName = candidateNames[i];
        if (typeof GEAPA_MEMBROS[fnName] === 'function') {
          const raw = GEAPA_MEMBROS[fnName](payload) || {};
          return {
            ok: true,
            implemented: true,
            raw: raw,
            message: 'Integracao opcional executada por GEAPA_MEMBROS.' + (raw.message ? ' ' + raw.message : ''),
          };
        }
      }

      return { ok: false, implemented: false, message: neutralMessage };
    }
  );
}
