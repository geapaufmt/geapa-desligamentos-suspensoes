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
        return { sent: false, dryRun: true, message: 'DRY_RUN: notificacao da secretaria nao enviada.' };
      }

      const htmlBody = [
        '<div style="font-family:Arial,sans-serif;line-height:1.6;">',
        '<p>Foi registrado um novo pedido no modulo de desligamentos e suspensoes.</p>',
        '<p>',
        '<strong>Tipo:</strong> ' + off_escapeHtml_(request.requestType) + '<br>',
        '<strong>Modalidade:</strong> ' + off_escapeHtml_(request.executionMode) + '<br>',
        '<strong>Membro:</strong> ' + off_escapeHtml_(request.name) + '<br>',
        '<strong>RGA:</strong> ' + off_escapeHtml_(request.rga) + '<br>',
        '<strong>E-mail:</strong> ' + off_escapeHtml_(request.email),
        '</p>',
        '<p>A linha ja foi roteada para a fila operacional correspondente e esta pronta para analise da diretoria.</p>',
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
        return { sent: false, message: 'E-mail do membro ausente ou invalido para enviar decisao.' };
      }

      if (control.dryRun) {
        return { sent: false, dryRun: true, message: 'DRY_RUN: e-mail de decisao nao enviado.' };
      }

      const actionLabel = payload.approved ? 'deferida' : 'indeferida';
      const followup = payload.approved
        ? 'A solicitacao seguira o fluxo operacional correspondente.'
        : 'Caso necessario, a diretoria podera orientar os proximos passos manualmente.';

      const htmlBody = [
        '<div style="font-family:Arial,sans-serif;line-height:1.6;">',
        '<p>Ola, <strong>' + off_escapeHtml_(payload.name || 'membro(a)') + '</strong>.</p>',
        '<p>Sua solicitacao de <strong>' + off_escapeHtml_(String(payload.requestType || '').toLowerCase()) + '</strong> foi <strong>' + actionLabel + '</strong> pela diretoria.</p>',
        '<p><strong>Modalidade:</strong> ' + off_escapeHtml_(payload.executionMode || '') + '</p>',
        '<p>' + off_escapeHtml_(followup) + '</p>',
        '<p>Atenciosamente,<br><strong>Diretoria do GEAPA</strong></p>',
        '</div>',
      ].join('');

      off_sendHtmlEmail_({
        to: payload.email,
        subject: OFF_CFG.EMAIL.DECISION_SUBJECT_PREFIX + ': ' + payload.requestType,
        htmlBody: htmlBody,
      });

      return { sent: true, message: 'E-mail de decisao enviado.' };
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
        return { sent: false, message: 'E-mail final nao enviado por ausencia de e-mail valido.' };
      }

      if (control.dryRun) {
        return { sent: false, dryRun: true, message: 'DRY_RUN: e-mail final nao enviado.' };
      }

      const htmlBody = [
        '<div style="margin:0;padding:0;background:#f5f6f7;font-family:Arial,Helvetica,sans-serif;color:#1f1f1f;">',
        '<div style="max-width:640px;margin:0 auto;padding:16px 12px;">',
        '<div style="background:#ffffff;border-radius:18px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.08);">',
        '<div style="padding:28px 24px 20px 24px;">',
        '<div style="font-size:26px;font-weight:800;line-height:1.15;color:#000000;margin-bottom:20px;">Desligamento confirmado</div>',
        '<div style="font-size:16px;line-height:1.6;color:#222222;">',
        '<p>Ola, <strong>' + off_escapeHtml_(payload.name || 'membro(a)') + '</strong>.</p>',
        '<p>Confirmamos que seu pedido de desligamento do GEAPA foi homologado e efetivado.</p>',
        '<p>Seu vinculo com o grupo encontra-se oficialmente encerrado a partir desta comunicacao.</p>',
        '<p>Agradecemos sua participacao e desejamos sucesso em sua trajetoria academica.</p>',
        '<p>Atenciosamente,<br><strong>Diretoria do GEAPA</strong></p>',
        '</div></div></div></div></div>',
      ].join('');

      off_sendHtmlEmail_({
        to: payload.email,
        subject: OFF_CFG.EMAIL.FINAL_DISMISSAL_SUBJECT,
        htmlBody: htmlBody,
      });

      return { sent: true, message: 'E-mail final de desligamento enviado.' };
    }
  );
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
