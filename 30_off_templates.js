/***************************************
 * 30_off_email_templates.gs
 * Templates e envios de e-mail
 ***************************************/

function off_notifySecretaries_(secretaryEmails, req) {
  if (!secretaryEmails || !secretaryEmails.length) return;

  const subject = 'GEAPA — Novo pedido aguardando análise';

  const htmlBody = `
    <div style="font-family:Arial,sans-serif; line-height:1.6;">
      <p>Olá, <strong>Secretários(as)</strong>.</p>

      <p>
        Foi registrado um novo pedido via formulário e ele está aguardando análise da Diretoria.
      </p>

      <p>
        <strong>Tipo de solicitação:</strong> ${off_safe_(req.tipo)}<br>
        <strong>Membro:</strong> ${off_safe_(req.nome)}<br>
        <strong>RGA:</strong> ${off_safe_(req.rga)}<br>
        <strong>E-mail:</strong> ${off_safe_(req.email)}
      </p>

      <p>
        <strong>Ação:</strong> acompanhar a deliberação e registrar a decisão na planilha.
      </p>

      <p>
        Atenciosamente,<br>
        <strong>Sistema GEAPA</strong>
      </p>
    </div>
  `;

  GEAPA_CORE.coreSendHtmlEmail({
    to: secretaryEmails.join(','),
    subject,
    htmlBody,
  });
}

function off_sendFinalEmailToMember_(req) {
  if (!GEAPA_CORE.coreIsValidEmail(req.email)) return;

  const isOff = String(req.tipo || '').trim() === OFF_CFG.VALUES.TYPE_OFFBOARD;
  const subject = isOff ? OFF_CFG.EMAIL.FINAL_SUBJECT_OFFBOARD : OFF_CFG.EMAIL.FINAL_SUBJECT_SUSPEND;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;">
      <h2>${isOff ? 'Desligamento confirmado' : 'Suspensão confirmada'}</h2>
      <p>Olá, <b>${off_firstName_(req.nome)}</b>.</p>
      <p>Confirmamos o processamento do seu pedido no GEAPA.</p>
      <p><b>Tipo:</b> ${off_safe_(req.tipo)}</p>
      <hr>
      <p>Atenciosamente,<br>GEAPA</p>
    </div>
  `;

  GEAPA_CORE.coreSendHtmlEmail({
    to: req.email,
    subject,
    htmlBody,
  });
}

function off_firstName_(full) {
  const s = String(full || '').trim();
  return s ? s.split(/\s+/)[0] : '';
}

function off_safe_(s) {
  return String(s || '').replace(/[<>&"]/g, ch => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[ch]));
}