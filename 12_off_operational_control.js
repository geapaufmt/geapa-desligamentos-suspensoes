/***************************************
 * 12_off_operational_control.gs
 * Integracao conservadora com MODULOS_CONFIG e MODULOS_STATUS do GEAPA-CORE.
 ***************************************/

const OFF_OPS = Object.freeze({
  MODULE_NAME: 'DESLIGAMENTOS',
  FLOWS: Object.freeze({
    GERAL: 'GERAL',
    RECEBIMENTO_FORMULARIO: 'RECEBIMENTO_FORMULARIO',
    DECISAO_DIRETORIA: 'DECISAO_DIRETORIA',
    EXECUCAO_SUSPENSAO: 'EXECUCAO_SUSPENSAO',
    EXECUCAO_DESLIGAMENTO: 'EXECUCAO_DESLIGAMENTO',
    INTEGRACAO_MEMBROS: 'INTEGRACAO_MEMBROS',
    EMAIL_FINAL: 'EMAIL_FINAL',
    EMAIL_NOTIFICACAO: 'EMAIL_NOTIFICACAO',
    EVENTOS_VINCULO: 'EVENTOS_VINCULO',
    JOB_DIARIO_FILAS: 'JOB_DIARIO_FILAS',
  }),
  CAPABILITIES: Object.freeze({
    SYNC: 'SYNC',
    EMAIL: 'EMAIL',
  }),
  EXECUTION_TYPES: Object.freeze({
    TRIGGER: 'TRIGGER',
    MANUAL: 'MANUAL',
  }),
});

/**
 * Resolve se a chamada veio de trigger/evento ou de invocacao manual.
 */
function off_getExecutionTypeFromEvent_(e) {
  return e && typeof e === 'object'
    ? OFF_OPS.EXECUTION_TYPES.TRIGGER
    : OFF_OPS.EXECUTION_TYPES.MANUAL;
}

/**
 * Normaliza o tipo de execucao esperado pelo GEAPA-CORE.
 */
function off_normalizeExecutionType_(value) {
  const key = off_normalizeTextKey_(value || '');
  return key === OFF_OPS.EXECUTION_TYPES.TRIGGER
    ? OFF_OPS.EXECUTION_TYPES.TRIGGER
    : OFF_OPS.EXECUTION_TYPES.MANUAL;
}

/**
 * Executa um fluxo sob controle central do GEAPA-CORE, registrando status quando possivel.
 */
function off_runControlledFlow_(flowName, capability, options, runner) {
  const control = off_beginControlledFlow_(flowName, capability, options);
  if (!control.allowed) {
    return off_buildBlockedResult_(control, control.reason);
  }

  try {
    const result = runner(control);
    if (result && result.blocked) {
      off_markOperationalBlocked_(Object.freeze({
        moduleName: control.moduleName,
        flowName: control.flowName,
        capability: control.capability,
        modeRead: result.modeRead || control.modeRead,
        reason: result.message || control.reason,
        runId: control.runId,
        obs: off_joinMessage_(control.obs, 'Bloqueio propagado por subfluxo.'),
      }));
      return result;
    }
    off_markOperationalSuccess_(control, result && result.dryRun
      ? 'DRY_RUN concluido sem efeitos reais.'
      : ''
    );
    return result;
  } catch (err) {
    off_markOperationalError_(control, err);
    throw err;
  }
}

/**
 * Consulta MODULOS_CONFIG e marca inicio/bloqueio em MODULOS_STATUS.
 */
function off_beginControlledFlow_(flowName, capability, options) {
  options = options || {};

  const control = {
    moduleName: OFF_OPS.MODULE_NAME,
    flowName: flowName || OFF_OPS.FLOWS.GERAL,
    capability: capability || OFF_OPS.CAPABILITIES.SYNC,
    executionType: off_normalizeExecutionType_(options.executionType),
    runId: options.runId || off_newRunId_(),
    modeRead: 'NAO_LIDO',
    dryRun: false,
    allowed: false,
    blocked: false,
    reason: '',
    obs: off_buildOperationalObs_(options),
    config: null,
  };

  if (
    typeof GEAPA_CORE === 'undefined' ||
    !GEAPA_CORE ||
    typeof GEAPA_CORE.coreGetModuleConfig !== 'function' ||
    typeof GEAPA_CORE.coreAssertModuleExecutionAllowed !== 'function'
  ) {
    control.allowed = true;
    control.reason = 'GEAPA_CORE indisponivel; seguindo em modo compativel.';
    off_logWarn_(control.runId, 'Controle operacional do GEAPA-CORE indisponivel; seguindo em modo compativel.', {
      flowName: control.flowName,
      capability: control.capability,
    });
    return Object.freeze(control);
  }

  try {
    control.config = GEAPA_CORE.coreGetModuleConfig(control.moduleName, control.flowName, {
      executionType: control.executionType,
    });
    control.modeRead = control.config && control.config.mode ? control.config.mode : 'NAO_LIDO';

    const decision = GEAPA_CORE.coreAssertModuleExecutionAllowed(
      control.moduleName,
      control.flowName,
      control.capability,
      { executionType: control.executionType }
    );

    control.allowed = true;
    control.dryRun = decision && decision.dryRun === true;
    control.reason = decision && decision.reason ? decision.reason : 'PERMITIDO';

    off_markOperationalExecution_(control);
    if (control.dryRun) {
      off_logInfo_(control.runId, 'Fluxo liberado em DRY_RUN pelo GEAPA-CORE.', {
        flowName: control.flowName,
        capability: control.capability,
        modeRead: control.modeRead,
      });
    }
    return Object.freeze(control);
  } catch (err) {
    control.blocked = true;
    control.allowed = false;
    control.reason = err && err.message ? err.message : String(err);
    off_markOperationalBlocked_(control);
    off_logWarn_(control.runId, 'Fluxo bloqueado pelo GEAPA-CORE.', {
      flowName: control.flowName,
      capability: control.capability,
      modeRead: control.modeRead,
      reason: control.reason,
    });
    return Object.freeze(control);
  }
}

/**
 * Construi um resumo padrao para MODULOS_STATUS.OBS.
 */
function off_buildOperationalObs_(options) {
  options = options || {};
  const parts = [];

  if (options.runId) {
    parts.push('runId=' + options.runId);
  }
  if (options.origin) {
    parts.push('origin=' + options.origin);
  }
  if (options.sheetName) {
    parts.push('sheet=' + options.sheetName);
  }
  if (options.rowNumber) {
    parts.push('row=' + options.rowNumber);
  }
  if (options.obs) {
    parts.push(String(options.obs));
  }

  return parts.join(' | ');
}

/**
 * Marca ultima execucao do fluxo em MODULOS_STATUS.
 */
function off_markOperationalExecution_(control) {
  off_safeCoreStatusCall_(function() {
    GEAPA_CORE.coreModuleStatusMarkExecution(control.moduleName, control.flowName, control.capability, {
      modeRead: control.modeRead,
      obs: control.obs,
    });
  }, control, 'markExecution');
}

/**
 * Marca ultimo sucesso do fluxo em MODULOS_STATUS.
 */
function off_markOperationalSuccess_(control, extraObs) {
  off_safeCoreStatusCall_(function() {
    GEAPA_CORE.coreModuleStatusMarkSuccess(control.moduleName, control.flowName, control.capability, {
      modeRead: control.modeRead,
      obs: off_joinMessage_(control.obs, extraObs || ''),
    });
  }, control, 'markSuccess');
}

/**
 * Marca ultimo erro do fluxo em MODULOS_STATUS.
 */
function off_markOperationalError_(control, err) {
  off_safeCoreStatusCall_(function() {
    GEAPA_CORE.coreModuleStatusMarkError(control.moduleName, control.flowName, err, control.capability, {
      modeRead: control.modeRead,
      obs: control.obs,
    });
  }, control, 'markError');
}

/**
 * Marca bloqueio por configuracao em MODULOS_STATUS.
 */
function off_markOperationalBlocked_(control) {
  off_safeCoreStatusCall_(function() {
    GEAPA_CORE.coreModuleStatusMarkBlocked(
      control.moduleName,
      control.flowName,
      off_mapOperationalBlockReasonCode_(control.reason),
      control.reason,
      control.capability,
      control.modeRead,
      { obs: control.obs }
    );
  }, control, 'markBlocked');
}

/**
 * Padroniza codigos curtos de bloqueio para facilitar leitura operacional.
 */
function off_mapOperationalBlockReasonCode_(reason) {
  const text = String(reason || '');
  if (text.indexOf('ATIVO=NAO') >= 0) return 'ATIVO_NAO';
  if (text.indexOf('MODO=OFF') >= 0) return 'MODO_OFF';
  if (text.indexOf('MODO=MANUAL') >= 0) return 'MODO_MANUAL';
  if (text.indexOf('Capability bloqueada') >= 0) return 'CAPABILITY_BLOQUEADA';
  if (text.indexOf('MODULOS_CONFIG nao possui configuracao') >= 0) return 'CONFIG_INEXISTENTE';
  if (text.indexOf('Capability invalida') >= 0) return 'CAPABILITY_INVALIDA';
  return 'BLOQUEADO_CONFIG';
}

/**
 * Executa chamadas de status como melhor esforco, sem derrubar o fluxo principal.
 */
function off_safeCoreStatusCall_(fn, control, actionName) {
  try {
    if (
      typeof GEAPA_CORE === 'undefined' ||
      !GEAPA_CORE ||
      typeof fn !== 'function'
    ) {
      return;
    }
    fn();
  } catch (err) {
    off_logWarn_(control && control.runId ? control.runId : off_newRunId_(), 'Falha ao registrar MODULOS_STATUS.', {
      actionName: actionName,
      err: String(err),
      flowName: control && control.flowName ? control.flowName : '',
    });
  }
}

/**
 * Converte o bloqueio em objeto neutro para os chamadores tratarem sem excecao.
 */
function off_buildBlockedResult_(control, message) {
  return {
    blocked: true,
    dryRun: false,
    sent: false,
    ok: false,
    message: message || control.reason || 'Fluxo bloqueado por configuracao operacional.',
    modeRead: control.modeRead || 'NAO_LIDO',
    flowName: control.flowName,
    capability: control.capability,
  };
}
