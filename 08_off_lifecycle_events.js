/***************************************
 * 08_off_lifecycle_events.gs
 * Camada oficial de escrita na trilha MEMBER_EVENTOS_VINCULO.
 ***************************************/

/**
 * Helper base para append institucional de eventos de vinculo.
 */
function off_appendLifecycleEvent_(payload) {
  if (
    typeof GEAPA_CORE === 'undefined' ||
    !GEAPA_CORE ||
    typeof GEAPA_CORE.coreAppendMemberLifecycleEvent !== 'function'
  ) {
    throw new Error('GEAPA_CORE.coreAppendMemberLifecycleEvent nao esta disponivel.');
  }

  const result = GEAPA_CORE.coreAppendMemberLifecycleEvent({
    rga: payload.rga,
    eventType: payload.eventType,
    eventDate: payload.eventDate || new Date(),
    eventStatus: OFF_CFG.VALUES.HOMOLOGADO,
    reason: payload.reason || '',
    sourceModule: OFF_CFG.SOURCE_MODULE,
    sourceKey: payload.sourceKey || OFF_KEYS.RESPONSES,
    sourceRow: payload.sourceRow || '',
    memberName: payload.memberName || '',
    memberEmail: payload.memberEmail || '',
    notes: payload.notes || '',
  });

  if (!result) {
    throw new Error('Falha ao registrar evento em ' + OFF_KEYS.MEMBER_EVENTS + '.');
  }
  return result;
}

/**
 * Registra evento institucional de suspensao.
 */
function off_appendSuspensionEvent_(payload) {
  return off_appendLifecycleEvent_(Object.assign({}, payload, {
    eventType: 'SUSPENSAO',
  }));
}

/**
 * Registra evento institucional de retorno de suspensao.
 */
function off_appendReturnEvent_(payload) {
  return off_appendLifecycleEvent_(Object.assign({}, payload, {
    eventType: 'RETORNO',
  }));
}

/**
 * Registra evento institucional de desligamento voluntario.
 */
function off_appendDismissalEvent_(payload) {
  return off_appendLifecycleEvent_(Object.assign({}, payload, {
    eventType: 'DESLIGAMENTO_VOLUNTARIO',
  }));
}
