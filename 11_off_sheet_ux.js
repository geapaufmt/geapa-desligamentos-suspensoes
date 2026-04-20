/***************************************
 * 11_off_sheet_ux.gs
 * UX operacional das filas de suspensao e desligamento.
 ***************************************/

const OFF_SHEET_UX = Object.freeze({
  colors: Object.freeze({
    neutralText: '#1f1f1f',
    defaultHeader: '#f3f3f3',
  }),
  compactTextHeaders: Object.freeze([
    'MOTIVO',
    'OBSERVACOES_MEMBRO',
    'OBS_DECISAO',
    'MENSAGEM_PROCESSAMENTO',
    'MENSAGEM_INTEGRACAO_MEMBROS',
    'ERRO_EXECUCAO',
  ]),
  notes: Object.freeze({
    SUSPENSION: Object.freeze({
      ID_SOLICITACAO: 'Identificador imutavel da solicitacao roteada a partir da aba bruta do formulario.',
      ORIGEM_LINHA_FORM: 'Linha original da submissao na aba bruta do Forms. Usada para evitar duplicidade.',
      DATA_PEDIDO: 'Timestamp de entrada do pedido voluntario.',
      NOME: 'Nome do membro informado no formulario.',
      RGA: 'RGA usado como identidade principal entre modulos.',
      EMAIL: 'E-mail do membro informado no formulario.',
      TIPO_SOLICITACAO: 'Tipo canonico interno. Nesta fila deve permanecer SUSPENSAO.',
      MODALIDADE_EXECUCAO: 'Modalidade canonica interna. Nesta fila deve permanecer TEMPORARIA.',
      MOTIVO: 'Motivo informado pelo membro para a suspensao voluntaria.',
      OBSERVACOES_MEMBRO: 'Observacoes livres registradas pelo membro no formulario.',
      DATA_INICIO_SUSPENSAO: 'Data prevista para gerar o evento institucional SUSPENSAO.',
      DATA_FIM_SUSPENSAO: 'Data prevista para gerar o evento institucional RETORNO.',
      QTDE_DIAS_SUSPENSAO: 'Duracao calculada automaticamente de forma inclusiva.',
      VALIDACAO_MEMBRO_ENCONTRADO: 'Resultado da busca automatica do membro em MEMBERS_ATUAIS.',
      VALIDACAO_PERIODO_MINIMO: 'Indica se o periodo minimo institucional de 30 dias foi atendido.',
      VALIDACAO_CONFLITO_APRESENTACAO: 'Sinaliza conflito de apresentacao no periodo quando a integracao existir.',
      VALIDACAO_PENDENCIA_ARQUIVO: 'Sinaliza pendencia de arquivo relacionada a apresentacoes.',
      VALIDACAO_OUTRAS_PENDENCIAS: 'Espaco reservado para novas validacoes futuras sem travar o fluxo.',
      VALIDACAO_GERAL: 'Resumo agregado das validacoes automaticas.',
      DECISAO_DIRETORIA: 'Preenchimento manual da diretoria. Use apenas DEFERIDO ou INDEFERIDO.',
      DATA_DECISAO: 'Timestamp da decisao manual registrada na fila.',
      OBS_DECISAO: 'Observacoes da diretoria sobre o deferimento ou indeferimento.',
      STATUS_SOLICITACAO: 'Estado operacional do pedido ao longo do fluxo.',
      EMAIL_DECISAO_ENVIADO: 'Marca se o e-mail de decisao ja foi comunicado ao membro.',
      DATA_EMAIL_DECISAO: 'Timestamp do envio do e-mail de decisao.',
      SUSPENSAO_APLICADA: 'Marca se o evento institucional SUSPENSAO ja foi registrado e aplicado.',
      DATA_EXECUCAO: 'Momento em que a suspensao foi efetivamente aplicada.',
      RETORNO_PROCESSADO: 'Marca se o evento institucional RETORNO ja foi registrado.',
      DATA_RETORNO_PROCESSADO: 'Momento em que o retorno foi processado.',
      ID_EVENTO_SUSPENSAO: 'Id do evento SUSPENSAO em MEMBER_EVENTOS_VINCULO.',
      ID_EVENTO_RETORNO: 'Id do evento RETORNO em MEMBER_EVENTOS_VINCULO.',
      MENSAGEM_PROCESSAMENTO: 'Log funcional resumido do fluxo desta linha.',
      ERRO_EXECUCAO: 'Ultimo erro operacional encontrado para esta solicitacao.',
      RUN_ID_ULTIMO_PROCESSAMENTO: 'RunId da ultima execucao automatica ou manual.',
      NOTIFICACAO_SECRETARIA_ENVIADA: 'Indica se a secretaria foi notificada sobre o novo pedido.',
      DATA_NOTIFICACAO_SECRETARIA: 'Timestamp da notificacao inicial para a secretaria.',
    }),
    DISMISSAL: Object.freeze({
      ID_SOLICITACAO: 'Identificador imutavel da solicitacao roteada a partir da aba bruta do formulario.',
      ORIGEM_LINHA_FORM: 'Linha original da submissao na aba bruta do Forms. Usada para evitar duplicidade.',
      DATA_PEDIDO: 'Timestamp de entrada do pedido voluntario.',
      NOME: 'Nome do membro informado no formulario.',
      RGA: 'RGA usado como identidade principal entre modulos.',
      EMAIL: 'E-mail do membro informado no formulario.',
      TIPO_SOLICITACAO: 'Tipo canonico interno. Nesta fila deve permanecer DESLIGAMENTO.',
      MODALIDADE_EXECUCAO: 'Modalidade canonica interna: IMEDIATO ou FIM_SEMESTRE.',
      MOTIVO: 'Motivo informado pelo membro para o desligamento voluntario.',
      OBSERVACOES_MEMBRO: 'Observacoes livres registradas pelo membro no formulario.',
      DATA_EFETIVA_DESLIGAMENTO: 'Data em que o desligamento deve ser efetivado.',
      SEMESTRE_REFERENCIA: 'Semestre usado para calcular ou contextualizar o desligamento agendado.',
      VALIDACAO_MEMBRO_ENCONTRADO: 'Resultado da busca automatica do membro em MEMBERS_ATUAIS.',
      VALIDACAO_CONFLITO_APRESENTACAO: 'Sinaliza conflito de apresentacao no periodo quando a integracao existir.',
      VALIDACAO_PENDENCIA_ARQUIVO: 'Sinaliza pendencia de arquivo relacionada a apresentacoes.',
      VALIDACAO_OUTRAS_PENDENCIAS: 'Espaco reservado para novas validacoes futuras sem travar o fluxo.',
      VALIDACAO_GERAL: 'Resumo agregado das validacoes automaticas.',
      DECISAO_DIRETORIA: 'Preenchimento manual da diretoria. Use apenas DEFERIDO ou INDEFERIDO.',
      DATA_DECISAO: 'Timestamp da decisao manual registrada na fila.',
      OBS_DECISAO: 'Observacoes da diretoria sobre o deferimento ou indeferimento.',
      STATUS_SOLICITACAO: 'Estado operacional do pedido ao longo do fluxo.',
      EMAIL_DECISAO_ENVIADO: 'Marca se o e-mail de decisao ja foi comunicado ao membro.',
      DATA_EMAIL_DECISAO: 'Timestamp do envio do e-mail de decisao.',
      EMAIL_FINAL_ENVIADO: 'Marca se o e-mail final de desligamento efetivado ja foi enviado.',
      DATA_EMAIL_FINAL: 'Timestamp do envio do e-mail final.',
      INTEGRACAO_MEMBROS_PROCESSADA: 'Status da integracao com GEAPA_MEMBROS.',
      DATA_INTEGRACAO_MEMBROS: 'Momento da integracao com MEMBERS_ATUAIS/MEMBERS_HIST.',
      MENSAGEM_INTEGRACAO_MEMBROS: 'Resumo do resultado da integracao com GEAPA_MEMBROS.',
      DESLIGAMENTO_EXECUTADO: 'Marca se o desligamento foi efetivamente executado.',
      DATA_EXECUCAO: 'Momento da execucao efetiva do desligamento.',
      ID_EVENTO_DESLIGAMENTO: 'Id do evento DESLIGAMENTO_VOLUNTARIO em MEMBER_EVENTOS_VINCULO.',
      ERRO_EXECUCAO: 'Ultimo erro operacional encontrado para esta solicitacao.',
      RUN_ID_ULTIMO_PROCESSAMENTO: 'RunId da ultima execucao automatica ou manual.',
      NOTIFICACAO_SECRETARIA_ENVIADA: 'Indica se a secretaria foi notificada sobre o novo pedido.',
      DATA_NOTIFICACAO_SECRETARIA: 'Timestamp da notificacao inicial para a secretaria.',
    }),
  }),
  groups: Object.freeze({
    SUSPENSION: Object.freeze([
      { color: '#d9eaf7', headers: ['ID_SOLICITACAO', 'ORIGEM_LINHA_FORM', 'DATA_PEDIDO', 'NOME', 'RGA', 'EMAIL', 'TIPO_SOLICITACAO', 'MODALIDADE_EXECUCAO', 'MOTIVO', 'OBSERVACOES_MEMBRO'] },
      { color: '#e8f3dc', headers: ['DATA_INICIO_SUSPENSAO', 'DATA_FIM_SUSPENSAO', 'QTDE_DIAS_SUSPENSAO'] },
      { color: '#fff1cc', headers: ['VALIDACAO_MEMBRO_ENCONTRADO', 'VALIDACAO_PERIODO_MINIMO', 'VALIDACAO_CONFLITO_APRESENTACAO', 'VALIDACAO_PENDENCIA_ARQUIVO', 'VALIDACAO_OUTRAS_PENDENCIAS', 'VALIDACAO_GERAL'] },
      { color: '#f4d7f5', headers: ['DECISAO_DIRETORIA', 'DATA_DECISAO', 'OBS_DECISAO', 'STATUS_SOLICITACAO', 'EMAIL_DECISAO_ENVIADO', 'DATA_EMAIL_DECISAO'] },
      { color: '#f9ded7', headers: ['SUSPENSAO_APLICADA', 'DATA_EXECUCAO', 'RETORNO_PROCESSADO', 'DATA_RETORNO_PROCESSADO', 'ID_EVENTO_SUSPENSAO', 'ID_EVENTO_RETORNO', 'MENSAGEM_PROCESSAMENTO', 'ERRO_EXECUCAO'] },
      { color: '#e6e6e6', headers: ['RUN_ID_ULTIMO_PROCESSAMENTO', 'NOTIFICACAO_SECRETARIA_ENVIADA', 'DATA_NOTIFICACAO_SECRETARIA'] },
    ]),
    DISMISSAL: Object.freeze([
      { color: '#d9eaf7', headers: ['ID_SOLICITACAO', 'ORIGEM_LINHA_FORM', 'DATA_PEDIDO', 'NOME', 'RGA', 'EMAIL', 'TIPO_SOLICITACAO', 'MODALIDADE_EXECUCAO', 'MOTIVO', 'OBSERVACOES_MEMBRO'] },
      { color: '#e8f3dc', headers: ['DATA_EFETIVA_DESLIGAMENTO', 'SEMESTRE_REFERENCIA'] },
      { color: '#fff1cc', headers: ['VALIDACAO_MEMBRO_ENCONTRADO', 'VALIDACAO_CONFLITO_APRESENTACAO', 'VALIDACAO_PENDENCIA_ARQUIVO', 'VALIDACAO_OUTRAS_PENDENCIAS', 'VALIDACAO_GERAL'] },
      { color: '#f4d7f5', headers: ['DECISAO_DIRETORIA', 'DATA_DECISAO', 'OBS_DECISAO', 'STATUS_SOLICITACAO', 'EMAIL_DECISAO_ENVIADO', 'DATA_EMAIL_DECISAO'] },
      { color: '#f9ded7', headers: ['EMAIL_FINAL_ENVIADO', 'DATA_EMAIL_FINAL', 'INTEGRACAO_MEMBROS_PROCESSADA', 'DATA_INTEGRACAO_MEMBROS', 'MENSAGEM_INTEGRACAO_MEMBROS', 'DESLIGAMENTO_EXECUTADO', 'DATA_EXECUCAO', 'ID_EVENTO_DESLIGAMENTO', 'ERRO_EXECUCAO'] },
      { color: '#e6e6e6', headers: ['RUN_ID_ULTIMO_PROCESSAMENTO', 'NOTIFICACAO_SECRETARIA_ENVIADA', 'DATA_NOTIFICACAO_SECRETARIA'] },
    ]),
  }),
  dropdowns: Object.freeze({
    SUSPENSION: Object.freeze({
      TIPO_SOLICITACAO: off_buildOffDropdownRule_(['SUSPENSAO'], 'Campo canonico da fila de suspensao.'),
      MODALIDADE_EXECUCAO: off_buildOffDropdownRule_(['TEMPORARIA'], 'Campo canonico da fila de suspensao.'),
      VALIDACAO_MEMBRO_ENCONTRADO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Resultado da validacao automatica de membro encontrado.'),
      VALIDACAO_PERIODO_MINIMO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Resultado da validacao automatica do periodo minimo.'),
      VALIDACAO_CONFLITO_APRESENTACAO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de conflito de apresentacao.'),
      VALIDACAO_PENDENCIA_ARQUIVO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de pendencia de arquivo.'),
      VALIDACAO_OUTRAS_PENDENCIAS: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de outras pendencias.'),
      VALIDACAO_GERAL: off_buildOffDropdownRule_(['OK', 'ALERTA', 'NAO_VERIFICADO', 'ERRO'], 'Resumo agregado das validacoes automaticas.'),
      DECISAO_DIRETORIA: off_buildOffDropdownRule_(['DEFERIDO', 'INDEFERIDO'], 'Decisao manual da diretoria.'),
      STATUS_SOLICITACAO: off_buildOffDropdownRule_(['RECEBIDO', 'VALIDADO', 'EM_ANALISE', 'DEFERIDO', 'INDEFERIDO', 'AGUARDANDO_EXECUCAO', 'EXECUTADO', 'CANCELADO', 'ERRO'], 'Estado operacional da solicitacao.'),
      EMAIL_DECISAO_ENVIADO: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle do envio do e-mail de decisao.'),
      SUSPENSAO_APLICADA: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle da aplicacao da suspensao.'),
      RETORNO_PROCESSADO: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle do processamento do retorno.'),
      NOTIFICACAO_SECRETARIA_ENVIADA: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle da notificacao inicial para a secretaria.'),
    }),
    DISMISSAL: Object.freeze({
      TIPO_SOLICITACAO: off_buildOffDropdownRule_(['DESLIGAMENTO'], 'Campo canonico da fila de desligamento.'),
      MODALIDADE_EXECUCAO: off_buildOffDropdownRule_(['IMEDIATO', 'FIM_SEMESTRE'], 'Modalidade operacional do desligamento.'),
      VALIDACAO_MEMBRO_ENCONTRADO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Resultado da validacao automatica de membro encontrado.'),
      VALIDACAO_CONFLITO_APRESENTACAO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de conflito de apresentacao.'),
      VALIDACAO_PENDENCIA_ARQUIVO: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de pendencia de arquivo.'),
      VALIDACAO_OUTRAS_PENDENCIAS: off_buildOffDropdownRule_(['SIM', 'NAO', 'NAO_VERIFICADO'], 'Sinalizador de outras pendencias.'),
      VALIDACAO_GERAL: off_buildOffDropdownRule_(['OK', 'ALERTA', 'NAO_VERIFICADO', 'ERRO'], 'Resumo agregado das validacoes automaticas.'),
      DECISAO_DIRETORIA: off_buildOffDropdownRule_(['DEFERIDO', 'INDEFERIDO'], 'Decisao manual da diretoria.'),
      STATUS_SOLICITACAO: off_buildOffDropdownRule_(['RECEBIDO', 'VALIDADO', 'EM_ANALISE', 'DEFERIDO', 'INDEFERIDO', 'AGUARDANDO_EXECUCAO', 'EXECUTADO', 'CANCELADO', 'ERRO'], 'Estado operacional da solicitacao.'),
      EMAIL_DECISAO_ENVIADO: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle do envio do e-mail de decisao.'),
      EMAIL_FINAL_ENVIADO: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle do envio do e-mail final.'),
      INTEGRACAO_MEMBROS_PROCESSADA: off_buildOffDropdownRule_(['SIM', 'NAO', 'ERRO'], 'Resultado da integracao com GEAPA_MEMBROS.'),
      DESLIGAMENTO_EXECUTADO: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle da execucao do desligamento.'),
      NOTIFICACAO_SECRETARIA_ENVIADA: off_buildOffDropdownRule_(['SIM', 'NAO'], 'Controle da notificacao inicial para a secretaria.'),
    }),
  }),
});

function applyOffboardingSheetUx() {
  return off_applyOperationalSheetsUx_();
}

function reapplyOffboardingSheetUx() {
  return off_applyOperationalSheetsUx_();
}

function off_applyOperationalSheetsUx_() {
  const sheets = {
    suspensions: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_SUSPENSIONS, OFF_CFG.SHEET_HEADERS.SUSPENSION),
    dismissals: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_DISMISSALS, OFF_CFG.SHEET_HEADERS.DISMISSAL),
  };
  return off_tryApplyOperationalSheetsUx_(sheets);
}

function off_tryApplyOperationalSheetsUx_(sheets) {
  const resolved = sheets || {
    suspensions: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_SUSPENSIONS, OFF_CFG.SHEET_HEADERS.SUSPENSION),
    dismissals: off_ensureSheetByRegistryKey_(OFF_KEYS.QUEUE_DISMISSALS, OFF_CFG.SHEET_HEADERS.DISMISSAL),
  };
  return Object.freeze({
    suspensions: off_tryApplySheetUx_(resolved.suspensions, 'SUSPENSION'),
    dismissals: off_tryApplySheetUx_(resolved.dismissals, 'DISMISSAL'),
  });
}

function off_tryApplySheetUx_(sheet, profileName) {
  try {
    return off_applySheetUx_(sheet, profileName);
  } catch (err) {
    return Object.freeze({
      ok: false,
      sheetName: sheet ? sheet.getName() : '',
      profileName: profileName,
      error: err && err.message ? err.message : String(err),
    });
  }
}

function off_applySheetUx_(sheet, profileName) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const operations = [];
  const notesMap = OFF_SHEET_UX.notes[profileName] || {};
  const colorGroups = OFF_SHEET_UX.groups[profileName] || [];
  const validationRules = OFF_SHEET_UX.dropdowns[profileName] || {};

  operations.push(off_trySheetUxOperation_(sheet, 'freezeHeaderRow', function() {
    if (off_coreHas_('coreFreezeHeaderRow')) {
      GEAPA_CORE.coreFreezeHeaderRow(sheet, 1);
    } else {
      sheet.setFrozenRows(1);
    }
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'headerNotes', function() {
    if (off_coreHas_('coreApplyHeaderNotes')) {
      GEAPA_CORE.coreApplyHeaderNotes(sheet, notesMap, 1);
    } else {
      off_applyHeaderNotesFallback_(sheet, notesMap);
    }
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'headerStyles', function() {
    if (off_coreHas_('coreApplyHeaderColors')) {
      GEAPA_CORE.coreApplyHeaderColors(sheet, colorGroups, 1, {
        defaultColor: OFF_SHEET_UX.colors.defaultHeader,
        fontColor: OFF_SHEET_UX.colors.neutralText,
        fontWeight: 'bold',
        wrap: true,
      });
    } else {
      off_applyHeaderColorsFallback_(sheet, colorGroups);
    }
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'filter', function() {
    if (off_coreHas_('coreEnsureFilter')) {
      GEAPA_CORE.coreEnsureFilter(sheet, 1, { recreate: true });
    } else {
      off_ensureFilterFallback_(sheet);
    }
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'compactLongText', function() {
    off_applyCompactTextUx_(sheet, OFF_SHEET_UX.compactTextHeaders);
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'dataAlignment', function() {
    if (lastRow < 2 || lastColumn < 1) {
      return;
    }
    sheet.getRange(2, 1, lastRow - 1, lastColumn)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }));

  operations.push(off_trySheetUxOperation_(sheet, 'dropdownValidation', function() {
    if (!validationRules || !Object.keys(validationRules).length) {
      return;
    }
    if (off_coreHas_('coreApplyDropdownValidationByHeader')) {
      GEAPA_CORE.coreApplyDropdownValidationByHeader(sheet, validationRules, 1, {});
    } else {
      off_applyDropdownValidationFallback_(sheet, validationRules);
    }
  }));

  return Object.freeze({
    ok: true,
    sheetName: sheet.getName(),
    profileName: profileName,
    operations: Object.freeze(operations),
  });
}

function off_buildOffDropdownRule_(values, helpText) {
  return Object.freeze({
    values: values,
    allowInvalid: true,
    helpText: helpText,
  });
}

function off_coreHas_(fnName) {
  return typeof GEAPA_CORE !== 'undefined' && GEAPA_CORE && typeof GEAPA_CORE[fnName] === 'function';
}

function off_trySheetUxOperation_(sheet, operation, fn) {
  try {
    fn();
    return Object.freeze({ operation: operation, status: 'APPLIED' });
  } catch (err) {
    return Object.freeze({
      operation: operation,
      status: 'SKIPPED',
      reason: err && err.message ? err.message : String(err),
      sheetName: sheet.getName(),
    });
  }
}

function off_applyHeaderNotesFallback_(sheet, notesMap) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  headers.forEach(function(header, idx) {
    if (notesMap[header]) {
      sheet.getRange(1, idx + 1).setNote(notesMap[header]);
    }
  });
}

function off_applyHeaderColorsFallback_(sheet, groups) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  const headerRange = sheet.getRange(1, 1, 1, Math.max(headers.length, 1));
  headerRange
    .setBackground(OFF_SHEET_UX.colors.defaultHeader)
    .setFontColor(OFF_SHEET_UX.colors.neutralText)
    .setFontWeight('bold')
    .setWrap(true);

  (groups || []).forEach(function(group) {
    (group.headers || []).forEach(function(headerName) {
      const idx = headers.indexOf(headerName);
      if (idx >= 0) {
        sheet.getRange(1, idx + 1).setBackground(group.color || OFF_SHEET_UX.colors.defaultHeader);
      }
    });
  });
}

function off_ensureFilterFallback_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const currentFilter = sheet.getFilter();
  if (currentFilter) {
    currentFilter.remove();
  }
  sheet.getRange(1, 1, lastRow, lastColumn).createFilter();
}

function off_applyCompactTextUx_(sheet, headerNames) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 1);

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });

  (headerNames || []).forEach(function(headerName) {
    const idx = headers.indexOf(headerName);
    if (idx >= 0 && lastRow > 1) {
      sheet.getRange(2, idx + 1, lastRow - 1, 1)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setVerticalAlignment('middle');
    }
  });
}

function off_applyDropdownValidationFallback_(sheet, rulesByHeader) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  const lastRow = Math.max(sheet.getMaxRows(), 2);

  Object.keys(rulesByHeader || {}).forEach(function(headerName) {
    const idx = headers.indexOf(headerName);
    if (idx < 0) {
      return;
    }

    const rule = rulesByHeader[headerName] || {};
    const builder = SpreadsheetApp.newDataValidation()
      .requireValueInList(rule.values || [], true)
      .setAllowInvalid(rule.allowInvalid !== false);

    if (rule.helpText) {
      builder.setHelpText(rule.helpText);
      sheet.getRange(1, idx + 1).setNote(rule.helpText);
    }

    sheet.getRange(2, idx + 1, lastRow - 1, 1).setDataValidation(builder.build());
  });
}
