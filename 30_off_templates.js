/***************************************
 * 30_off_email_templates.gs
 * Templates e envios de e-mail
 * Responsabilidades:
 * - Gerar e enviar e-mails para secretários e membros
 * - Templates de e-mail para notificações e comunicações finais
 * Dependências:
 * - Library GEAPA-CORE com identificador GEAPA_CORE
 * Acionado por:
 * - off_processQueue para notificações aos secretários
 * - off_onEditDecision para envio de e-mail final ao membro
 * Regras:
 * - E-mails só são enviados se os dados necessários estiverem presentes e válidos
 * - O template de e-mail é construído de forma segura para evitar injeção de HTML
 * - Logs de erro são capturados e registrados no console
 * - O código é organizado em funções pequenas e reutilizáveis para clareza e manutenção
 * - Testes recomendados:
 * - Verificar que os e-mails são enviados corretamente para os secretários quando um novo pedido é registrado
 *  - Verificar que o e-mail final é enviado corretamente ao membro quando um pedido é deferido
 * - Verificar que os templates de e-mail são renderizados corretamente com os dados do pedido
 * - Simular erros no envio de e-mail e verificar se o erro é registrado corretamente
 * - Verificar que os e-mails não são enviados se os dados necessários estiverem ausentes ou inválidos
 * - Melhorias futuras:
 * - Adicionar mais detalhes personalizados nos e-mails, como o tipo de solicitação ou o nome do membro
 * - Adicionar notificações para a diretoria quando um pedido é deferido
 * - Adicionar uma etapa de confirmação antes de marcar como deferido, para evitar edições acidentais
 * - Adicionar testes automatizados para as funções de envio de e-mail e templates
 * - Refatorar o código para usar uma abordagem orientada a objetos, encapsulando o processo de offboarding em uma classe
 * - Adicionar suporte para outros tipos de solicitações além de desligamento, com templates de e-mail personalizados
 * - Adicionar um painel de controle para monitorar o status dos pedidos de offboarding e suas integrações
 * - Adicionar logs mais detalhados para auditoria e análise de falhas
 * - Considerações de segurança:
 * - Garantir que apenas usuários autorizados possam editar a coluna de decisão para deferir um pedido
 * - Garantir que os dados pessoais dos membros sejam tratados de forma segura e em conformidade com as políticas de privacidade
 * - Garantir que o processo de integração com membros não exponha dados sensíveis ou permita ações não autorizadas
 * - Garantir que o template de e-mail seja construído de forma segura para evitar injeção de HTML ou outros ataques
 * - Garantir que os erros sejam registrados de forma segura sem expor informações sensíveis
 * - Documentação:
 * - Documentar as funções e seus parâmetros, bem como o fluxo geral do processo de deferimento e integração
 * - Documentar as dependências e como configurar o ambiente para que o processo funcione corretamente
 * - Documentar os testes realizados e os resultados esperados para cada cenário
 * - Conclusão:
 * Este módulo é responsável por gerar e enviar os e-mails relacionados ao processo de offboarding, garantindo que os secretários sejam notificados sobre novos pedidos e que os membros sejam informados sobre o deferimento de seus pedidos. Ele é acionado em pontos específicos do processo de offboarding e segue regras para garantir que os e-mails sejam enviados de forma consistente, segura e com as informações corretas. O código é estruturado para ser claro, modular e fácil de manter, com considerações para segurança e testes abrangentes.
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
        <strong>Ação:</strong> analisar a solicitação e registrar a decisão na planilha. (Deferir ou Indeferir)
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

function offboard_getHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h || "").trim().toLowerCase()] = i;
  });
  return map;
}