/***************************************
 * IGNORE ---
 * 11_off_finalize_deferred.js
 * IGNORE ---
 *
 * Finaliza o processo de deferimento imediato, enviando e-mail final e integrando com membros.
 * Acionado por off_onEditDecision ao detectar edição para DEFERIDO.
 *
 * Regras para finalizar: 
 * - Tipo de solicitação deve ser "Desligamento"
 * - Desejo de desligamento deve ser "Imediatamente"
 * - Decisão da diretoria deve ser "DEFERIDO"
 * - Se e-mail final ainda não foi enviado, envia o e-mail final e marca como enviado
 * - Integra com membros se ainda não integrou
 * - Usa LockService para evitar concorrência na finalização de uma mesma solicitação
 * - O e-mail final é um template simples confirmando o deferimento e agradecendo a participação
 * - Depende de:
 *   - offboard_processMembersIntegration_ para integração com membros
 *   - GEAPA_MEMBROS para integração de membros
 * - O template de e-mail é construído de forma segura para evitar injeção de HTML
 * - Logs de erro são capturados e registrados no console
 * - O código é organizado em funções pequenas e reutilizáveis para clareza e manutenção
 * - Testes recomendados:
 *  - Editar a coluna "Decisão da Diretoria" para "DEFERIDO" e verificar se o e-mail é enviado e a integração ocorre
 * - Editar para outros valores e verificar que nada é acionado
 * - Editar a célula de decisão para "DEFERIDO" em uma linha já processada e verificar que o processo não é repetido
 * - Simular erros na função de integração e verificar se o erro é registrado e não impede o envio do e-mail
 * - Simular erros no envio de e-mail e verificar se o erro é registrado e não impede a integração
 * - Verificar que a data/hora do deferimento é registrada corretamente
 * - Verificar que a data/hora do deferimento não é sobrescrita se o e-mail já estava marcado como enviado
 * - Verificar que a integração com membros só ocorre uma vez por solicitação deferida
 * - Melhorias futuras:
 * - Adicionar mais detalhes personalizados no e-mail final, como o tipo de solicitação ou o nome do membro
 * - Adicionar notificações para a diretoria quando um pedido é deferido
 * - Adicionar uma etapa de confirmação antes de marcar como deferido, para evitar edições acidentais
 * - Adicionar testes automatizados para as funções de finalização e integração
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
 * Este módulo é responsável por finalizar o processo de deferimento imediato, garantindo que o membro seja notificado e que a integração com membros ocorra de forma consistente e segura. Ele é acionado por uma edição na coluna de decisão da diretoria e segue regras específicas para garantir que apenas os pedidos de desligamento imediato sejam processados. O código é estruturado para ser claro, modular e fácil de manter, com considerações para segurança e testes abrangentes.
 ***************************************/

function off_finalizeDeferredRequest_(sh, sheetRow) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error("Não foi possível obter lock para finalizar deferimento.");
  }

  try {
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const map = offboard_getHeaderMap_(headers);
    const row = sh.getRange(sheetRow, 1, 1, lastCol).getValues()[0];

    function getCell_(candidates) {
      const idx = offboard_findHeaderIndex_(map, candidates);
      return idx >= 0 ? row[idx] : "";
    }

    const tipo = String(getCell_([OFF_CFG.RESP.TYPE, "Quero solicitar:"]) || "").trim();
    const nome = String(getCell_([OFF_CFG.RESP.NAME, "Nome Completo"]) || "").trim();
    const rga = String(getCell_([OFF_CFG.RESP.RGA, "RGA"]) || "").trim();
    const email = String(getCell_([OFF_CFG.RESP.EMAIL, "Endere�o de e-mail", "Endereco de e-mail", "E-mail", "Email"]) || "").trim();
    const leaveTiming = String(getCell_(["Desejo me desligar:"]) || "").trim();
    const approved = String(getCell_([OFF_CFG.RESP.APPROVED, "Decis�o da Diretoria"]) || "").trim().toUpperCase();
    const sentFinal = String(getCell_([OFF_CFG.RESP.SENT_FINAL, "E-mail de desligamento enviado?"]) || "").trim().toUpperCase();
    const membersProcessed = String(getCell_(["Integra��o membros processada?", "Integracao membros processada?"]) || "").trim().toUpperCase();

    if (tipo !== "Desligamento") return;
    if (leaveTiming !== "Imediatamente") return;
    if (approved !== "DEFERIDO") return;

    const now = new Date();

    // 1) Se ainda não enviou o e-mail final, envia
    if (sentFinal !== "SIM") {
      off_sendFinalEmailToMember_({ tipo, nome, rga, email });

      const sentFinalCol = offboard_findHeaderIndex_(map, [OFF_CFG.RESP.SENT_FINAL, "E-mail de desligamento enviado?"]);
      if (sentFinalCol >= 0) {
        sh.getRange(sheetRow, sentFinalCol + 1).setValue("SIM");
      }

      const deferCell = offboard_findHeaderIndex_(map, [OFF_CFG.RESP.DEFER_DATE, "Data/Hora do deferimento"]);
      if (deferCell >= 0) {
        sh.getRange(sheetRow, deferCell + 1).setValue(now);
      }
    } else {
      // Se já estava como enviado, garante pelo menos a data se estiver vazia
      const deferCell = offboard_findHeaderIndex_(map, [OFF_CFG.RESP.DEFER_DATE, "Data/Hora do deferimento"]);
      if (deferCell >= 0) {
        const currentDeferDate = sh.getRange(sheetRow, deferCell + 1).getValue();
        if (!currentDeferDate) {
          sh.getRange(sheetRow, deferCell + 1).setValue(now);
        }
      }
    }

    // 2) Integra com membros se ainda não integrou
    if (membersProcessed !== "SIM") {
      offboard_processMembersIntegration_(sh, sheetRow);
    }

  } finally {
    lock.releaseLock();
  }
}

function off_onEditDecision(e) {
  try {
    if (!e || !e.range) return;

    const editedSheet = e.range.getSheet();
    const { sh } = off_readResponsesTable_();
    if (!sh) return;

    if (editedSheet.getSheetId() !== sh.getSheetId()) return;
    if (e.range.getRow() < 2) return;

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
    const map = offboard_getHeaderMap_(headers);

    const cDecision = offboard_findHeaderIndex_(map, [OFF_CFG.RESP.APPROVED, "Decis�o da Diretoria"]);
    if (cDecision < 0) return;

    // só reage se a célula editada for a coluna "Decisão da Diretoria"
    if (e.range.getColumn() !== cDecision + 1) return;

    const newValue = String(e.value || "").trim().toUpperCase();
    if (newValue !== "DEFERIDO") return;

    off_finalizeDeferredRequest_(sh, e.range.getRow());

  } catch (err) {
    console.error("off_onEditDecision erro:", err);
  }
}

function buildOffFinalEmailHtml_(data) {
  const safeName = escapeOffHtml_(data.nome || "membro(a)");
  const safeTipo = escapeOffHtml_(data.tipo || "Desligamento");

  return `
    <div style="margin:0; padding:0; background:#f5f6f7; font-family:Arial, Helvetica, sans-serif; color:#1f1f1f; -webkit-text-size-adjust:100%; text-size-adjust:100%;">
      <div style="max-width:640px; margin:0 auto; padding:16px 12px;">
        <div style="background:#ffffff; border-radius:18px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.08);">

          <div style="padding:28px 24px 20px 24px;">
            <div style="font-size:26px; font-weight:800; line-height:1.15; color:#000000; margin-bottom:20px;">
              ${safeTipo}<br>confirmado
            </div>

            <div style="font-size:16px; line-height:1.6; color:#222222;">
              <p style="margin:0 0 18px 0;">Olá, <strong>${safeName}</strong>!</p>

              <p style="margin:0 0 18px 0;">
                Confirmamos que seu pedido de <strong>desligamento do GEAPA</strong> foi
                analisado e <strong>DEFERIDO</strong>.
              </p>

              <p style="margin:0 0 18px 0;">
                A partir deste momento, seu vínculo com o grupo encontra-se oficialmente encerrado.
              </p>

              <p style="margin:0 0 24px 0;">
                Agradecemos sua participação e desejamos sucesso em sua trajetória acadêmica.
              </p>

              <div style="border-top:1px solid #d9d9d9; padding-top:18px; margin-top:10px;">
                <p style="margin:0;">
                  Atenciosamente,<br>
                  <strong>Diretoria do GEAPA</strong>
                </p>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>
  `;
}

function escapeOffHtml_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function off_sendFinalEmailToMember_(data) {
  if (!GEAPA_CORE.coreIsValidEmail(data.email)) return;

  const html = buildOffFinalEmailHtml_(data);

  MailApp.sendEmail({
    to: data.email,
    subject: "Desligamento confirmado - GEAPA",
    htmlBody: html,
    name: "GEAPA"
  });
}
