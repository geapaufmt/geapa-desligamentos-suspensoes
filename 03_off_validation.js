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
  const hoursEligibility = request.requestType === OFF_TYPES.DESLIGAMENTO
    ? off_checkHoursEligibility_(request.rga, request.semesterReference, {
        requestType: request.requestType,
        executionMode: request.executionMode,
        requestDate: request.requestDate,
      })
    : null;

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
  if (hoursEligibility && hoursEligibility.sourceFlag === OFF_CFG.VALUES.PENDENTE) {
    issues.push('A elegibilidade a horas complementares ficou pendente para analise da diretoria.');
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
    hoursEligibility ? hoursEligibility.message : '',
  ].filter(Boolean);

  return {
    memberFoundFlag: memberLookup.flag,
    minimumPeriodFlag: minimumPeriod.flag,
    presentationConflictFlag: off_toFlag_(aggregate.hasPresentationInPeriod || aggregate.hasFuturePresentation),
    pendingFileFlag: off_toFlag_(aggregate.hasPendingFile),
    otherPendingFlag: otherPending.flag,
    hoursRightsFlag: hoursEligibility && hoursEligibility.sourceFlag ? hoursEligibility.sourceFlag : '',
    generalStatus: generalStatus,
    message: messageParts.join(' | ') || 'Validacao inicial registrada.',
    details: {
      memberLookup: memberLookup,
      aggregate: aggregate,
      otherPending: otherPending,
      hoursEligibility: hoursEligibility,
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
 * Apura elegibilidade a horas complementares usando presencas como fonte operacional e apresentacoes como confirmacao.
 */
function off_checkHoursEligibility_(rga, periodoRef, context) {
  const manualFlag = off_normalizeHoursRightsFlag_(context && context.manualFlag ? context.manualFlag : '');
  if (manualFlag) {
    return {
      eligible: manualFlag === OFF_CFG.VALUES.YES
        ? true
        : (manualFlag === OFF_CFG.VALUES.NO ? false : null),
      sourceType: 'MANUAL',
      sourceFlag: manualFlag,
      presentedInPeriod: manualFlag === OFF_CFG.VALUES.YES ? true : null,
      attendancePercent: null,
      disciplinaryStatus: null,
      status: manualFlag === OFF_CFG.VALUES.PENDENTE ? OFF_CFG.VALUES.ALERTA : OFF_CFG.VALUES.OK,
      message: 'Elegibilidade de horas complementares definida manualmente na fila.',
      basis: null,
      details: {
        context: context || null,
      },
    };
  }

  const presenceAssessment = off_tryAssessHoursEligibilityFromPresence_(rga, periodoRef);
  if (!presenceAssessment.connected) {
    return {
      eligible: null,
      sourceType: 'AUTOMATICO',
      sourceFlag: OFF_CFG.VALUES.PENDENTE,
      presentedInPeriod: null,
      attendancePercent: null,
      disciplinaryStatus: null,
      status: OFF_CFG.VALUES.NAO_VERIFICADO,
      message: presenceAssessment.message || 'Base de presencas indisponivel para apuracao automatica.',
      basis: null,
      details: {
        presenceAssessment: presenceAssessment,
        context: context || null,
      },
    };
  }

  const aggregate = off_checkPresentationAndFilePending_(rga, periodoRef, context);
  const result = {
    eligible: presenceAssessment.eligible,
    sourceType: 'AUTOMATICO',
    sourceFlag: presenceAssessment.eligible === true
      ? OFF_CFG.VALUES.YES
      : (presenceAssessment.eligible === false ? OFF_CFG.VALUES.NO : OFF_CFG.VALUES.PENDENTE),
    presentedInPeriod: presenceAssessment.presentedInPeriod,
    attendancePercent: presenceAssessment.attendancePercent,
    disciplinaryStatus: presenceAssessment.disciplinaryStatus,
    status: presenceAssessment.status,
    message: presenceAssessment.message,
    basis: presenceAssessment.basis,
    details: {
      presenceAssessment: presenceAssessment,
      aggregate: aggregate,
      context: context || null,
    },
  };

  if (
    result.presentedInPeriod !== true &&
    aggregate.status === OFF_CFG.VALUES.OK &&
    aggregate.hasPresentationInPeriod === true &&
    aggregate.hasFuturePresentation !== true
  ) {
    result.presentedInPeriod = true;
    if (result.eligible === null && off_isRegularDisciplinaryStatus_(result.disciplinaryStatus)) {
      result.eligible = true;
      result.sourceFlag = OFF_CFG.VALUES.YES;
      result.message = off_joinMessage_(result.message, 'Apresentação confirmada pela base canônica de atividades.');
      result.status = OFF_CFG.VALUES.OK;
    }
  }

  if (result.eligible === true && aggregate.hasPendingFile === true) {
    result.eligible = null;
    result.sourceFlag = OFF_CFG.VALUES.PENDENTE;
    result.status = OFF_CFG.VALUES.ALERTA;
    result.message = off_joinMessage_(result.message, 'Arquivo pendente identificado; elegibilidade mantida como PENDENTE.');
  }

  return result;
}

/**
 * Consulta a linha operacional de presencas e interpreta os principais indicadores do periodo.
 */
function off_tryAssessHoursEligibilityFromPresence_(rga, periodoRef) {
  const sheet = off_getOptionalSheetByKey_(OFF_KEYS.ACTIVITIES_PRESENCAS) || off_findSheetByName_('Presencas');
  if (!sheet) {
    return {
      connected: false,
      message: 'Aba de presencas nao localizada para apuracao de horas complementares.',
    };
  }

  const table = off_readSheetTable_(sheet);
  const rgaCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.PRESENCE_RGA);
  if (!rgaCol) {
    return {
      connected: false,
      message: 'Base de presencas sem coluna RGA reconhecida.',
    };
  }

  const semesterCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.PRESENCE_SEMESTER);
  const presentedCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.PRESENTED_IN_PERIOD);
  const forecastCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.PRESENTATION_FORECAST);
  const attendanceCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.ATTENDANCE_PERCENT);
  const totalActivitiesCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.TOTAL_ABSENCE_ACTIVITIES);
  const absenceLimitCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.ABSENCE_LIMIT);
  const netAbsencesCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.NET_ABSENCES);
  const usagePercentCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.LIMIT_USAGE_PERCENT);
  const disciplinaryCol = off_findColumn_(table.headerMap, OFF_CFG.ACTIVITIES_FIELDS.DISCIPLINARY_STATUS);

  for (let i = 0; i < table.rows.length; i += 1) {
    const row = table.rows[i];
    if (String(row[rgaCol - 1] || '').trim() !== String(rga || '').trim()) {
      continue;
    }

    if (periodoRef && semesterCol) {
      const semesterValue = String(row[semesterCol - 1] || '').trim();
      if (semesterValue && off_normalizeTextKey_(semesterValue) !== off_normalizeTextKey_(periodoRef)) {
        continue;
      }
    }

    const presentedFlag = presentedCol ? off_normalizeYesNoFlag_(row[presentedCol - 1]) : '';
    const forecastFlag = forecastCol ? off_normalizeYesNoFlag_(row[forecastCol - 1]) : '';
    const attendancePercent = attendanceCol ? off_parseNumber_(row[attendanceCol - 1]) : null;
    const totalActivities = totalActivitiesCol ? off_parseNumber_(row[totalActivitiesCol - 1]) : null;
    const absenceLimit = absenceLimitCol ? off_parseNumber_(row[absenceLimitCol - 1]) : null;
    const netAbsences = netAbsencesCol ? off_parseNumber_(row[netAbsencesCol - 1]) : null;
    const usagePercent = usagePercentCol ? off_parseNumber_(row[usagePercentCol - 1]) : null;
    const disciplinaryStatus = disciplinaryCol ? String(row[disciplinaryCol - 1] || '').trim() : '';

    const basis = {
      apresentouNoPeriodo: presentedFlag || null,
      previsaoApresentacaoNoPeriodo: forecastFlag || null,
      percentualFrequencia: attendancePercent,
      totalAtividadesQueContamFalta: totalActivities,
      limiteFaltasPeriodo: absenceLimit,
      faltasLiquidas: netAbsences,
      percentualUsoLimite: usagePercent,
      situacaoDisciplinar: disciplinaryStatus || null,
    };

    const presentedInPeriod = presentedFlag === OFF_CFG.VALUES.YES;
    const attendanceRegular = off_isAttendanceWithinLimit_(attendancePercent, netAbsences, absenceLimit, usagePercent);
    const disciplinaryRegular = off_isRegularDisciplinaryStatus_(disciplinaryStatus);

    if (presentedFlag === OFF_CFG.VALUES.NO && forecastFlag === OFF_CFG.VALUES.NO) {
      return {
        connected: true,
        eligible: false,
        presentedInPeriod: false,
        attendancePercent: attendancePercent,
        disciplinaryStatus: disciplinaryStatus || null,
        status: OFF_CFG.VALUES.OK,
        message: 'Base de presenças indica que o membro não apresentou no período.',
        basis: basis,
      };
    }

    if (presentedFlag === OFF_CFG.VALUES.YES && attendanceRegular === true && disciplinaryRegular === true) {
      return {
        connected: true,
        eligible: true,
        presentedInPeriod: true,
        attendancePercent: attendancePercent,
        disciplinaryStatus: disciplinaryStatus || null,
        status: OFF_CFG.VALUES.OK,
        message: 'Base de presenças indica apresentação no período e situação regular para horas complementares.',
        basis: basis,
      };
    }

    if (presentedFlag === OFF_CFG.VALUES.YES && (attendanceRegular === false || disciplinaryRegular === false)) {
      return {
        connected: true,
        eligible: false,
        presentedInPeriod: true,
        attendancePercent: attendancePercent,
        disciplinaryStatus: disciplinaryStatus || null,
        status: OFF_CFG.VALUES.OK,
        message: 'Base de presenças indica apresentação, mas sem regularidade suficiente para horas complementares.',
        basis: basis,
      };
    }

    if (presentedFlag === OFF_CFG.VALUES.YES && off_isPendingDisciplinaryStatus_(disciplinaryStatus)) {
      return {
        connected: true,
        eligible: null,
        presentedInPeriod: true,
        attendancePercent: attendancePercent,
        disciplinaryStatus: disciplinaryStatus || null,
        status: OFF_CFG.VALUES.ALERTA,
        message: 'Base de presenças indica apresentação no período, mas há pendência a concluir.',
        basis: basis,
      };
    }

    return {
      connected: true,
      eligible: null,
      presentedInPeriod: presentedInPeriod ? true : null,
      attendancePercent: attendancePercent,
      disciplinaryStatus: disciplinaryStatus || null,
      status: OFF_CFG.VALUES.ALERTA,
      message: 'Base de presenças localizada, mas sem dados suficientes para apuração automática conclusiva.',
      basis: basis,
    };
  }

  return {
    connected: true,
    eligible: null,
    presentedInPeriod: null,
    attendancePercent: null,
    disciplinaryStatus: null,
    status: OFF_CFG.VALUES.ALERTA,
    message: 'RGA não localizado na base de presenças para apuração de horas complementares.',
    basis: null,
  };
}

/**
 * Interpreta flags SIM/NAO de forma tolerante.
 */
function off_normalizeYesNoFlag_(value) {
  const key = off_normalizeTextKey_(value || '');
  if (key.indexOf('SIM') >= 0) {
    return OFF_CFG.VALUES.YES;
  }
  if (key.indexOf('NAO') >= 0) {
    return OFF_CFG.VALUES.NO;
  }
  return '';
}

/**
 * Faz parse tolerante de numeros vindos de planilha.
 */
function off_parseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }
  const normalized = String(value)
    .replace(/\s+/g, '')
    .replace('%', '')
    .replace(/\./g, '')
    .replace(',', '.');
  const parsed = Number(normalized);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Interpreta a situacao disciplinar com flexibilidade de nomenclatura.
 */
function off_isRegularDisciplinaryStatus_(status) {
  const key = off_normalizeTextKey_(status || '');
  if (!key) {
    return null;
  }
  if (
    key.indexOf('IRREGULAR') >= 0 ||
    key.indexOf('INAPTO') >= 0 ||
    key.indexOf('BLOQUE') >= 0 ||
    key.indexOf('REPROV') >= 0
  ) {
    return false;
  }
  if (
    key.indexOf('REGULAR') >= 0 ||
    key.indexOf('APTO') >= 0 ||
    key.indexOf('OK') >= 0 ||
    key.indexOf('ADIMPL') >= 0
  ) {
    return true;
  }
  return null;
}

/**
 * Identifica situações disciplinares ou administrativas que ainda dependem de conclusão posterior.
 */
function off_isPendingDisciplinaryStatus_(status) {
  const key = off_normalizeTextKey_(status || '');
  return key.indexOf('PEND') >= 0 || key.indexOf('ANALISE') >= 0 || key.indexOf('AGUARD') >= 0;
}

/**
 * Interpreta frequencia/faltas usando a melhor combinacao de colunas disponivel.
 */
function off_isAttendanceWithinLimit_(attendancePercent, netAbsences, absenceLimit, usagePercent) {
  if (netAbsences !== null && absenceLimit !== null) {
    return netAbsences <= absenceLimit;
  }
  if (usagePercent !== null) {
    return usagePercent <= 100;
  }
  if (attendancePercent !== null) {
    return attendancePercent > 0;
  }
  return null;
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
