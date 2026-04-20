/***************************************
 * 03_off_validation.gs
 * Validacoes e contratos de integracao futura com atividades.
 ***************************************/

/**
 * Valida a solicitacao normalizada sem bloquear o fluxo quando a integracao futura ainda nao existir.
 */
function off_validateNormalizedRequest_(request) {
  const memberLookup = off_lookupMemberInMembersBase_(request.rga, request.email);
  const aggregate = off_checkPresentationAndFilePending_(request.rga, request.semesterReference, {
    requestType: request.requestType,
    executionMode: request.executionMode,
    requestDate: request.requestDate,
  });
  const otherPending = off_checkOtherPendingItems_(request.rga, {
    semesterReference: request.semesterReference,
    requestDate: request.requestDate,
  });
  const minimumPeriod = request.requestType === OFF_TYPES.SUSPENSAO
    ? off_validateSuspensionPeriod_(request.suspensionStart, request.suspensionEnd)
    : { status: OFF_CFG.VALUES.OK, flag: OFF_CFG.VALUES.NAO_VERIFICADO, message: '' };

  const issues = [];
  if (minimumPeriod.status === OFF_CFG.VALUES.ERRO) {
    issues.push(minimumPeriod.message);
  }
  if (aggregate.hasFuturePresentation || aggregate.hasPresentationInPeriod || aggregate.hasPendingFile) {
    issues.push('Ha sinalizacao de apresentacao/arquivo pendente para analise da diretoria.');
  }
  if (memberLookup.status === OFF_CFG.VALUES.ERRO) {
    issues.push(memberLookup.message);
  }

  let generalStatus = OFF_CFG.VALUES.OK;
  if (issues.length) {
    generalStatus = OFF_CFG.VALUES.ALERTA;
  }
  if (
    aggregate.status === OFF_CFG.VALUES.NAO_VERIFICADO ||
    otherPending.status === OFF_CFG.VALUES.NAO_VERIFICADO ||
    memberLookup.flag === OFF_CFG.VALUES.NAO_VERIFICADO
  ) {
    generalStatus = OFF_CFG.VALUES.NAO_VERIFICADO;
  }
  if (minimumPeriod.status === OFF_CFG.VALUES.ERRO) {
    generalStatus = OFF_CFG.VALUES.ERRO;
  }

  const messageParts = [
    memberLookup.message,
    minimumPeriod.message,
    aggregate.message,
    otherPending.message,
  ].filter(Boolean);

  return {
    memberFoundFlag: memberLookup.flag,
    minimumPeriodFlag: minimumPeriod.flag,
    presentationConflictFlag: off_toFlag_(aggregate.hasPresentationInPeriod || aggregate.hasFuturePresentation),
    pendingFileFlag: off_toFlag_(aggregate.hasPendingFile),
    otherPendingFlag: otherPending.flag,
    generalStatus: generalStatus,
    message: messageParts.join(' | ') || 'Validacao inicial registrada.',
    details: {
      memberLookup: memberLookup,
      aggregate: aggregate,
      otherPending: otherPending,
    },
  };
}

/**
 * Faz lookup defensivo na base de membros para registrar se o RGA/e-mail foi localizado.
 */
function off_lookupMemberInMembersBase_(rga, email) {
  const membersSheet = off_getMembersSheet_();
  if (!membersSheet) {
    return {
      status: OFF_CFG.VALUES.NAO_VERIFICADO,
      flag: OFF_CFG.VALUES.NAO_VERIFICADO,
      message: 'Base de membros nao conectada para validacao automatica.',
    };
  }

  const table = off_readSheetTable_(membersSheet);
  const rgaCol = off_findColumn_(table.headerMap, ['RGA']);
  const emailCol = off_findColumn_(table.headerMap, ['EMAIL', 'E_MAIL']);
  if (!rgaCol && !emailCol) {
    return {
      status: OFF_CFG.VALUES.NAO_VERIFICADO,
      flag: OFF_CFG.VALUES.NAO_VERIFICADO,
      message: 'Base de membros sem cabecalhos reconhecidos para busca automatica.',
    };
  }

  const found = table.rows.some(function(row) {
    const rowRga = rgaCol ? String(row[rgaCol - 1] || '').trim() : '';
    const rowEmail = emailCol ? String(row[emailCol - 1] || '').trim().toLowerCase() : '';
    return (rga && rowRga && rowRga === rga) || (email && rowEmail && rowEmail === String(email).toLowerCase());
  });

  return {
    status: OFF_CFG.VALUES.OK,
    flag: found ? OFF_CFG.VALUES.YES : OFF_CFG.VALUES.NO,
    message: found
      ? 'Membro localizado na base atual.'
      : 'Membro nao localizado automaticamente na base atual.',
  };
}

/**
 * Valida datas da suspensao e o periodo minimo de 30 dias.
 */
function off_validateSuspensionPeriod_(startDate, endDate) {
  const days = off_calculateSuspensionDays_(startDate, endDate);
  if (!startDate || !endDate) {
    return {
      status: OFF_CFG.VALUES.ERRO,
      flag: OFF_CFG.VALUES.NO,
      message: 'Suspensao exige DATA_INICIO_SUSPENSAO e DATA_FIM_SUSPENSAO.',
    };
  }
  if (days === null) {
    return {
      status: OFF_CFG.VALUES.ERRO,
      flag: OFF_CFG.VALUES.NO,
      message: 'Periodo de suspensao invalido.',
    };
  }
  if (days < OFF_CFG.MIN_SUSPENSION_DAYS) {
    return {
      status: OFF_CFG.VALUES.ERRO,
      flag: OFF_CFG.VALUES.NO,
      message: 'Periodo minimo de suspensao e de ' + OFF_CFG.MIN_SUSPENSION_DAYS + ' dias.',
    };
  }
  return {
    status: OFF_CFG.VALUES.OK,
    flag: OFF_CFG.VALUES.YES,
    message: 'Periodo de suspensao validado automaticamente.',
  };
}

/**
 * Calcula a quantidade de dias da suspensao de forma inclusiva.
 */
function off_calculateSuspensionDays_(startDate, endDate) {
  const start = off_startOfDay_(startDate);
  const end = off_startOfDay_(endDate);
  if (!start || !end || end.getTime() < start.getTime()) {
    return null;
  }
  const diffMs = end.getTime() - start.getTime();
  return Math.round(diffMs / 86400000) + 1;
}

/**
 * Stub de conflito de apresentacao; delega para o agregador.
 */
function off_checkPresentationConflict_(rga, context) {
  const aggregate = off_checkPresentationAndFilePending_(rga, context && context.periodoRef, context);
  return {
    hasConflict: aggregate.hasPresentationInPeriod || aggregate.hasFuturePresentation,
    status: aggregate.status,
    message: aggregate.message,
    details: aggregate.details,
  };
}

/**
 * Stub de pendencia de arquivo; delega para o agregador.
 */
function off_checkPendingFiles_(rga, context) {
  const aggregate = off_checkPresentationAndFilePending_(rga, context && context.periodoRef, context);
  return {
    hasPendingFile: aggregate.hasPendingFile,
    status: aggregate.status,
    message: aggregate.message,
    details: aggregate.details,
  };
}

/**
 * Stub para pendencias futuras ainda nao modeladas nesta fase.
 */
function off_checkOtherPendingItems_(rga, context) {
  return {
    flag: OFF_CFG.VALUES.NAO_VERIFICADO,
    status: OFF_CFG.VALUES.NAO_VERIFICADO,
    message: 'Outras pendencias ainda nao conectadas nesta versao.',
    details: {
      rga: rga,
      context: context || null,
    },
  };
}

/**
 * Contrato agregado com tentativa de filtro rapido por presencas e confirmacao canonica em Atividades_Apresentacoes.
 */
function off_checkPresentationAndFilePending_(rga, periodoRef, context) {
  const result = {
    hasPresentationInPeriod: false,
    hasFuturePresentation: false,
    hasPendingFile: false,
    presentationDate: null,
    presentationStatus: null,
    fileStatus: null,
    source: {
      presencasFlag: null,
      atividadeApresentacaoId: null,
    },
    status: OFF_CFG.VALUES.NAO_VERIFICADO,
    message: 'Integracao com atividades ainda nao conectada nesta versao.',
    details: null,
  };

  if (!rga) {
    result.message = 'RGA ausente; checagem de atividades nao executada.';
    return result;
  }

  const quickFlag = off_tryReadActivitiesPresenceFlag_(rga);
  result.source.presencasFlag = quickFlag.flag;

  if (quickFlag.flag === OFF_CFG.VALUES.NO) {
    result.status = OFF_CFG.VALUES.OK;
    result.message = 'Filtro rapido por presencas indicou ausencia de apresentacao no periodo.';
    result.details = { quickFlag: quickFlag, context: context || null };
    return result;
  }

  const canonical = off_tryReadCanonicalPresentation_(rga, periodoRef);
  if (!canonical.connected) {
    result.message = canonical.message || result.message;
    result.details = { quickFlag: quickFlag, canonical: canonical, context: context || null };
    return result;
  }

  result.hasPresentationInPeriod = canonical.hasPresentationInPeriod;
  result.hasFuturePresentation = canonical.hasFuturePresentation;
  result.hasPendingFile = canonical.hasPendingFile;
  result.presentationDate = canonical.presentationDate;
  result.presentationStatus = canonical.presentationStatus;
  result.fileStatus = canonical.fileStatus;
  result.source.atividadeApresentacaoId = canonical.atividadeApresentacaoId;
  result.status = OFF_CFG.VALUES.OK;
  result.message = canonical.message;
  result.details = { quickFlag: quickFlag, canonical: canonical, context: context || null };
  return result;
}

/**
 * Le o sinalizador derivado de presencas quando a base existir.
 */
function off_tryReadActivitiesPresenceFlag_(rga) {
  const sheet = off_getOptionalSheetByKey_(OFF_KEYS.ACTIVITIES_PRESENCAS) || off_findSheetByName_('Presencas');
  if (!sheet) {
    return { connected: false, flag: null, message: 'Aba de presencas nao localizada.' };
  }

  const table = off_readSheetTable_(sheet);
  const rgaCol = off_findColumn_(table.headerMap, ['RGA']);
  const flagCol = off_findColumn_(table.headerMap, ['PREVISAO_APRESENTACAO_NO_PERIODO']);
  if (!rgaCol || !flagCol) {
    return { connected: false, flag: null, message: 'Campos de presencas nao reconhecidos.' };
  }

  for (let i = 0; i < table.rows.length; i += 1) {
    if (String(table.rows[i][rgaCol - 1] || '').trim() !== String(rga)) {
      continue;
    }
    const rawFlag = off_normalizeTextKey_(table.rows[i][flagCol - 1]);
    if (rawFlag.indexOf('NAO') >= 0) {
      return { connected: true, flag: OFF_CFG.VALUES.NO, message: 'Filtro rapido encontrou PREVISAO_APRESENTACAO_NO_PERIODO=NAO.' };
    }
    if (rawFlag.indexOf('SIM') >= 0) {
      return { connected: true, flag: OFF_CFG.VALUES.YES, message: 'Filtro rapido encontrou PREVISAO_APRESENTACAO_NO_PERIODO=SIM.' };
    }
  }

  return { connected: true, flag: null, message: 'Filtro rapido sem correspondencia para o RGA.' };
}

/**
 * Consulta a base canonica de apresentacoes quando disponivel.
 */
function off_tryReadCanonicalPresentation_(rga, periodoRef) {
  const sheet = off_getOptionalSheetByKey_(OFF_KEYS.ACTIVITIES_PRESENTATIONS) || off_findSheetByName_('Atividades_Apresentacoes');
  if (!sheet) {
    return {
      connected: false,
      message: 'Aba Atividades_Apresentacoes nao localizada.',
    };
  }

  const table = off_readSheetTable_(sheet);
  const rgaCol = off_findColumn_(table.headerMap, ['RGA']);
  const semesterCol = off_findColumn_(table.headerMap, ['SEMESTRE_APRESENTACAO', 'PERIODO_REFERENCIA']);
  const dateCol = off_findColumn_(table.headerMap, ['DATA_ATIVIDADE']);
  const presentationStatusCol = off_findColumn_(table.headerMap, ['STATUS_APRESENTACAO']);
  const fileStatusCol = off_findColumn_(table.headerMap, ['STATUS_ENVIO_ARQUIVO']);
  const linkCol = off_findColumn_(table.headerMap, ['LINK_ARQUIVO_DRIVE']);
  if (!rgaCol) {
    return {
      connected: false,
      message: 'Aba Atividades_Apresentacoes sem coluna RGA reconhecida.',
    };
  }

  const today = off_startOfDay_(new Date());
  let canonical = null;

  for (let i = 0; i < table.rows.length; i += 1) {
    const row = table.rows[i];
    if (String(row[rgaCol - 1] || '').trim() !== String(rga)) {
      continue;
    }

    if (periodoRef && semesterCol) {
      const semesterValue = String(row[semesterCol - 1] || '').trim();
      if (semesterValue && off_normalizeTextKey_(semesterValue) !== off_normalizeTextKey_(periodoRef)) {
        continue;
      }
    }

    const presentationDate = off_startOfDay_(dateCol ? row[dateCol - 1] : null);
    const presentationStatus = presentationStatusCol ? String(row[presentationStatusCol - 1] || '').trim() : '';
    const fileStatus = fileStatusCol ? String(row[fileStatusCol - 1] || '').trim() : '';
    const fileLink = linkCol ? String(row[linkCol - 1] || '').trim() : '';

    canonical = {
      connected: true,
      atividadeApresentacaoId: sheet.getName() + ':' + (table.startRow + i),
      hasPresentationInPeriod: true,
      hasFuturePresentation: !!(presentationDate && presentationDate.getTime() > today.getTime() && !off_isConcludedPresentationStatus_(presentationStatus)),
      hasPendingFile: off_isPendingFileStatus_(fileStatus, presentationDate, fileLink),
      presentationDate: presentationDate,
      presentationStatus: presentationStatus || null,
      fileStatus: fileStatus || null,
      message: 'Checagem confirmada na base canonica Atividades_Apresentacoes.',
    };

    if (canonical.hasFuturePresentation || canonical.hasPendingFile) {
      break;
    }
  }

  if (!canonical) {
    return {
      connected: true,
      atividadeApresentacaoId: null,
      hasPresentationInPeriod: false,
      hasFuturePresentation: false,
      hasPendingFile: false,
      presentationDate: null,
      presentationStatus: null,
      fileStatus: null,
      message: 'Sem apresentacao localizada na base canonica para o periodo consultado.',
    };
  }

  return canonical;
}

/**
 * Interpreta status de apresentacao de forma centralizada e flexivel.
 */
function off_isConcludedPresentationStatus_(status) {
  const key = off_normalizeTextKey_(status);
  return key.indexOf('REALIZ') >= 0 || key.indexOf('CONCLUI') >= 0 || key.indexOf('FINALIZ') >= 0;
}

/**
 * Interpreta status de envio de arquivo de forma centralizada e flexivel.
 */
function off_isPendingFileStatus_(status, presentationDate, fileLink) {
  if (fileLink) {
    return false;
  }
  const key = off_normalizeTextKey_(status);
  if (key.indexOf('RECEB') >= 0 || key.indexOf('CONCLUI') >= 0 || key.indexOf('ENTREG') >= 0) {
    return false;
  }
  const today = off_startOfDay_(new Date());
  if (presentationDate && presentationDate.getTime() > today.getTime()) {
    return false;
  }
  return true;
}

/**
 * Localiza aba pelo nome na mesma spreadsheet do modulo.
 */
function off_findSheetByName_(sheetName) {
  const ss = off_getResponseSpreadsheet_();
  return ss ? ss.getSheetByName(sheetName) : null;
}

/**
 * Deriva semestre de referencia do calendario civil.
 */
function off_deriveSemesterReference_(referenceDate) {
  const date = off_parseDate_(referenceDate) || new Date();
  const semester = date.getMonth() <= 5 ? 1 : 2;
  return date.getFullYear() + '/' + semester;
}

/**
 * Calcula o ultimo dia do semestre a partir da referencia informada.
 */
function off_computeSemesterEndDate_(semesterReference, fallbackDate) {
  const ref = off_normalizeTextKey_(semesterReference || '');
  const fallback = off_parseDate_(fallbackDate) || new Date();
  const match = ref.match(/(\d{4})_(\d)/);
  if (match) {
    const year = Number(match[1]);
    const semester = Number(match[2]);
    return semester === 1 ? new Date(year, 5, 30) : new Date(year, 11, 31);
  }
  return fallback.getMonth() <= 5
    ? new Date(fallback.getFullYear(), 5, 30)
    : new Date(fallback.getFullYear(), 11, 31);
}
