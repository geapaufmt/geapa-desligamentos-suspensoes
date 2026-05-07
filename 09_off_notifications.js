/***************************************
 * 09_off_notifications.gs
 * Templates e envios de e-mail do modulo.
 ***************************************/

/**
 * Notifica a secretaria quando um novo pedido foi roteado.
 */
function off_notifySecretariesOfNewRequest_(request) {
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.EMAIL_NOTIFICACAO,
    OFF_OPS.CAPABILITIES.EMAIL,
    {
      executionType: request && request.executionType ? request.executionType : OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: request && request.runId ? request.runId : '',
      origin: 'SECRETARIA_NOTIFICACAO',
    },
    function(control) {
      const secretaryEmails = off_getSecretaryEmails_();
      if (!secretaryEmails.length) {
        return { sent: false, message: 'Nenhum e-mail de secretaria encontrado.' };
      }

      if (control.dryRun) {
        return { sent: false, dryRun: true, message: 'DRY_RUN: notificação à secretaria não enviada.' };
      }

      const htmlBody = [
        '<div style="font-family:Arial,sans-serif;line-height:1.6;">',
        '<p>Foi registrado um novo pedido no módulo de desligamentos e suspensões.</p>',
        '<p>',
        '<strong>Tipo:</strong> ' + off_escapeHtml_(request.requestType) + '<br>',
        '<strong>Modalidade:</strong> ' + off_escapeHtml_(request.executionMode) + '<br>',
        '<strong>Membro:</strong> ' + off_escapeHtml_(request.name) + '<br>',
        '<strong>RGA:</strong> ' + off_escapeHtml_(request.rga) + '<br>',
        '<strong>E-mail:</strong> ' + off_escapeHtml_(request.email),
        '</p>',
        '<p>A linha já foi roteada para a fila operacional correspondente e está pronta para análise da diretoria.</p>',
        '</div>',
      ].join('');

      off_sendHtmlEmail_({
        to: secretaryEmails.join(','),
        subject: OFF_CFG.EMAIL.SECRETARY_SUBJECT_PREFIX + ': ' + request.requestType,
        htmlBody: htmlBody,
      });

      return { sent: true, message: 'Secretaria notificada sobre o novo pedido.' };
    }
  );
}

/**
 * Envia o e-mail de decisao sem acoplar a execucao final do desligamento.
 */
function off_sendDecisionEmail_(payload) {
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.EMAIL_NOTIFICACAO,
    OFF_OPS.CAPABILITIES.EMAIL,
    {
      executionType: payload && payload.executionType ? payload.executionType : OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: payload && payload.runId ? payload.runId : '',
      origin: 'DECISION_EMAIL',
    },
    function(control) {
      if (!off_isValidEmail_(payload.email)) {
        return { sent: false, message: 'E-mail do membro ausente ou inválido para enviar a decisão.' };
      }

      if (control.dryRun) {
        return { sent: false, dryRun: true, message: 'DRY_RUN: e-mail de decisão não enviado.' };
      }

      const actionLabel = payload.approved ? 'deferida' : 'indeferida';
      const followup = payload.approved
        ? 'A solicitação seguirá o fluxo operacional correspondente.'
        : 'Caso necessário, a diretoria poderá orientar os próximos passos manualmente.';
      const decisionNotesHtml = off_buildDecisionNotesHtml_(payload.decisionNotes);

      const htmlBody = [
        '<div style="font-family:Arial,sans-serif;line-height:1.6;">',
        '<p>Olá, <strong>' + off_escapeHtml_(payload.name || 'membro(a)') + '</strong>.</p>',
        '<p>Sua solicitação de <strong>' + off_escapeHtml_(String(payload.requestType || '').toLowerCase()) + '</strong> foi <strong>' + actionLabel + '</strong> pela diretoria.</p>',
        '<p><strong>Modalidade:</strong> ' + off_escapeHtml_(payload.executionMode || '') + '</p>',
        '<p>' + off_escapeHtml_(followup) + '</p>',
        decisionNotesHtml,
        '<p>Atenciosamente,<br><strong>Diretoria do GEAPA</strong></p>',
        '</div>',
      ].join('');

      off_sendHtmlEmail_({
        to: payload.email,
        subject: OFF_CFG.EMAIL.DECISION_SUBJECT_PREFIX + ': ' + payload.requestType,
        htmlBody: htmlBody,
      });

      return { sent: true, message: 'E-mail de decisão enviado.' };
    }
  );
}

/**
 * Envia o e-mail final de desligamento efetivado.
 */
function off_sendFinalDismissalEmail_(payload, options) {
  options = options || {};
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.EMAIL_FINAL,
    OFF_OPS.CAPABILITIES.EMAIL,
    {
      executionType: options.executionType || OFF_OPS.EXECUTION_TYPES.MANUAL,
      runId: options.runId,
      origin: options.origin || 'FINAL_EMAIL',
    },
    function(control) {
      if (!off_isValidEmail_(payload.email)) {
        return { sent: false, message: 'E-mail final não enviado por ausência de e-mail válido.' };
      }

      if (control.dryRun) {
        return { sent: false, dryRun: true, message: 'DRY_RUN: e-mail final não enviado.' };
      }

      const finalEmailContent = off_buildFinalDismissalEmailContent_(payload);

      off_sendHtmlEmail_({
        to: payload.email,
        subject: finalEmailContent.subject,
        htmlBody: finalEmailContent.htmlBody,
      });

      return {
        sent: true,
        message: 'E-mail final de desligamento enviado.',
        variant: finalEmailContent.variant,
      };
    }
  );
}

/**
 * Seleciona o template final conforme o direito do membro a horas complementares.
 */
function off_buildFinalDismissalEmailContent_(payload) {
  const rights = off_normalizeHoursRightsFlag_(payload && payload.hasHoursRights ? payload.hasHoursRights : '');
  const nameHtml = off_escapeHtml_(payload && payload.name ? payload.name : 'membro(a)');
  const decisionNotesHtml = off_buildDecisionNotesHtml_(payload && payload.decisionNotes ? payload.decisionNotes : '');
  const shellStart = [
    '<div style="margin:0;padding:0;background:#f5f6f7;font-family:Arial,Helvetica,sans-serif;color:#1f1f1f;">',
    '<div style="max-width:640px;margin:0 auto;padding:16px 12px;">',
    '<div style="background:#ffffff;border-radius:18px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08);">',
    '<div style="padding:28px 24px 20px 24px;">',
  ].join('');
  const shellEnd = [
    '<p>Atenciosamente,<br><strong>Diretoria do GEAPA</strong></p>',
    '</div></div></div></div></div>',
  ].join('');

  if (rights === OFF_CFG.VALUES.YES) {
    return {
      variant: 'COM_HORAS',
      subject: OFF_CFG.EMAIL.FINAL_DISMISSAL_WITH_HOURS_SUBJECT,
      htmlBody: [
        shellStart,
        '<div style="font-size:26px;font-weight:800;line-height:1.15;color:#000000;margin-bottom:20px;">Desligamento confirmado</div>',
        '<div style="font-size:16px;line-height:1.6;color:#222222;">',
        '<p>Olá, <strong>' + nameHtml + '</strong>.</p>',
        '<p>Confirmamos que seu pedido de desligamento do GEAPA foi homologado e efetivado.</p>',
        '<p>Seu vínculo com o grupo encontra-se oficialmente encerrado a partir desta comunicação.</p>',
        '<p>Conforme o registro deste período, você terá direito ao aproveitamento das horas complementares correspondentes.</p>',
        '<p>Se houver necessidade de comprovação formal, a diretoria poderá orientar a emissão ou o encaminhamento institucional cabível.</p>',
        decisionNotesHtml,
        shellEnd,
      ].join(''),
    };
  }

  if (rights === OFF_CFG.VALUES.NO) {
    return {
      variant: 'SEM_HORAS',
      subject: OFF_CFG.EMAIL.FINAL_DISMISSAL_WITHOUT_HOURS_SUBJECT,
      htmlBody: [
        shellStart,
        '<div style="font-size:26px;font-weight:800;line-height:1.15;color:#000000;margin-bottom:20px;">Desligamento confirmado</div>',
        '<div style="font-size:16px;line-height:1.6;color:#222222;">',
        '<p>Olá, <strong>' + nameHtml + '</strong>.</p>',
        '<p>Confirmamos que seu pedido de desligamento do GEAPA foi homologado e efetivado.</p>',
        '<p>Seu vínculo com o grupo encontra-se oficialmente encerrado a partir desta comunicação.</p>',
        '<p>Conforme o registro deste período, não há direito a horas complementares vinculadas a esta solicitação de desligamento.</p>',
        '<p>Se houver alguma dúvida sobre esse enquadramento, você pode solicitar esclarecimento diretamente à diretoria.</p>',
        decisionNotesHtml,
        shellEnd,
      ].join(''),
    };
  }

  if (rights === OFF_CFG.VALUES.PENDENTE) {
    return {
      variant: 'PENDENTE',
      subject: OFF_CFG.EMAIL.FINAL_DISMISSAL_PENDING_HOURS_SUBJECT,
      htmlBody: [
        shellStart,
        '<div style="font-size:26px;font-weight:800;line-height:1.15;color:#000000;margin-bottom:20px;">Desligamento confirmado</div>',
        '<div style="font-size:16px;line-height:1.6;color:#222222;">',
        '<p>Olá, <strong>' + nameHtml + '</strong>.</p>',
        '<p>Confirmamos que seu pedido de desligamento do GEAPA foi homologado e efetivado.</p>',
        '<p>Seu vínculo com o grupo encontra-se oficialmente encerrado a partir desta comunicação.</p>',
        '<p>A análise sobre horas complementares deste período permanece pendente, pois ainda há informações ou pendências a serem concluídas.</p>',
        '<p>A diretoria informará o resultado quando a análise for finalizada.</p>',
        decisionNotesHtml,
        shellEnd,
      ].join(''),
    };
  }

  return {
    variant: 'PADRAO',
    subject: OFF_CFG.EMAIL.FINAL_DISMISSAL_SUBJECT,
    htmlBody: [
      shellStart,
      '<div style="font-size:26px;font-weight:800;line-height:1.15;color:#000000;margin-bottom:20px;">Desligamento confirmado</div>',
      '<div style="font-size:16px;line-height:1.6;color:#222222;">',
      '<p>Olá, <strong>' + nameHtml + '</strong>.</p>',
      '<p>Confirmamos que seu pedido de desligamento do GEAPA foi homologado e efetivado.</p>',
      '<p>Seu vínculo com o grupo encontra-se oficialmente encerrado a partir desta comunicação.</p>',
      '<p>Agradecemos sua participação e desejamos sucesso em sua trajetória acadêmica.</p>',
      decisionNotesHtml,
      shellEnd,
    ].join(''),
  };
}

/**
 * Monta o bloco opcional com a observação registrada pela diretoria.
 */
function off_buildDecisionNotesHtml_(notes) {
  const clean = String(notes || '').trim();
  if (!clean) {
    return '';
  }
  return '<p><strong>Mensagem da diretoria:</strong> ' + off_escapeHtml_(clean) + '</p>';
}

/**
 * Normaliza o campo manual de direito a horas complementares.
 */
function off_normalizeHoursRightsFlag_(value) {
  const key = off_normalizeTextKey_(value || '');
  if (key === OFF_CFG.VALUES.YES) {
    return OFF_CFG.VALUES.YES;
  }
  if (key === OFF_CFG.VALUES.NO) {
    return OFF_CFG.VALUES.NO;
  }
  if (key === OFF_CFG.VALUES.PENDENTE) {
    return OFF_CFG.VALUES.PENDENTE;
  }
  return '';
}

/**
 * Busca os e-mails institucionais da secretaria.
 */
function off_getSecretaryEmails_() {
  if (
    typeof GEAPA_CORE !== 'undefined' &&
    GEAPA_CORE &&
    typeof GEAPA_CORE.coreGetCurrentEmailsByEmailGroup === 'function'
  ) {
    const emails = GEAPA_CORE.coreGetCurrentEmailsByEmailGroup('SECRETARIA') || [];
    return emails.filter(Boolean).filter(function(email, idx, arr) {
      return arr.indexOf(email) === idx;
    });
  }
  return [];
}

/**
 * Wrapper unico para envio de e-mail HTML.
 */
function off_sendHtmlEmail_(payload) {
  if (
    typeof GEAPA_CORE !== 'undefined' &&
    GEAPA_CORE &&
    typeof GEAPA_CORE.coreSendHtmlEmail === 'function'
  ) {
    GEAPA_CORE.coreSendHtmlEmail(payload);
    return;
  }

  MailApp.sendEmail({
    to: payload.to,
    subject: payload.subject,
    htmlBody: payload.htmlBody,
    name: OFF_CFG.EMAIL.FROM_NAME,
  });
}

/**
 * Validacao leve de e-mail para evitar excecoes evitaveis.
 */
function off_isValidEmail_(email) {
  if (
    typeof GEAPA_CORE !== 'undefined' &&
    GEAPA_CORE &&
    typeof GEAPA_CORE.coreIsValidEmail === 'function'
  ) {
    return GEAPA_CORE.coreIsValidEmail(email);
  }
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

/**
 * Escape simples de HTML para templates.
 */
function off_escapeHtml_(value) {
  return String(value || '').replace(/[<>&"]/g, function(ch) {
    return { '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[ch];
  });
}
