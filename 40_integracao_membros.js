/***************************************
 * 30_offboarding_members_integration.gs
 * Processa a integração de membros para desligamentos imediatos
 * deferidos, garantindo que o membro seja movido de MEMBERS_ATUAIS para MEMBERS_HIST
 * e registrando o resultado no formulário de offboarding.
 * Responsabilidades:
 * - Processar integração de membros para casos deferidos de desligamento imediato
 * - Registrar data/hora e resultado da integração no formulário
 * - Garantir que a integração só ocorra uma vez por solicitação
 * Dependências:
 * - Library GEAPA-MEMBROS com identificador GEAPA_MEMBROS
 * Acionado por:
 * - off_onEditDecision ao detectar edição para DEFERIDO
 ***************************************/

function offboard_processMembersIntegration_(sheet, rowNumber) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  const map = offboard_getHeaderMap_(headers);
  const row = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

  const requestType = String(row[map["quero solicitar:"]] || "").trim();
  const leaveTiming = String(row[map["desejo me desligar:"]] || "").trim();
  const decision = String(row[map["decisão da diretoria"]] || "").trim().toUpperCase();
  const finalEmailSent = String(row[map["e-mail de desligamento enviado?"]] || "").trim().toUpperCase();
  const alreadyProcessed = String(row[map["integração membros processada?"]] || "").trim().toUpperCase();

  if (requestType !== "Desligamento") return;
  if (leaveTiming !== "Imediatamente") return;
  if (decision !== "DEFERIDO") return;
  if (finalEmailSent !== "SIM") return;
  if (alreadyProcessed === "SIM") return;

  try {
    const payload = {
      requestRowNumber: rowNumber,
      requestTimestamp: row[map["carimbo de data/hora"]],
      memberName: row[map["nome completo"]],
      memberEmail: row[map["endereço de e-mail"]],
      memberRga: row[map["rga"]],
      requestType: row[map["quero solicitar:"]],
      leaveTiming: row[map["desejo me desligar:"]],
      reason: row[map["informe o motivo do pedido de desligamento:"]],
      approvedAt: row[map["data/hora do deferimento"]],
      decision: row[map["decisão da diretoria"]],
      finalEmailSent: row[map["e-mail de desligamento enviado?"]]
    };

    const result = GEAPA_MEMBROS.members_offboardApprovedImmediateExit(payload);

    sheet.getRange(rowNumber, map["integração membros processada?"] + 1).setValue("SIM");
    sheet.getRange(rowNumber, map["data/hora integração membros"] + 1).setValue(new Date());
    sheet.getRange(rowNumber, map["mensagem integração membros"] + 1).setValue(
      result.duplicatedHistory
        ? "Registro já existia em MEMBERS_HIST; membro removido de MEMBERS_ATUAIS."
        : "Membro movido com sucesso de MEMBERS_ATUAIS para MEMBERS_HIST."
    );

  } catch (err) {
    sheet.getRange(rowNumber, map["integração membros processada?"] + 1).setValue("ERRO");
    sheet.getRange(rowNumber, map["data/hora integração membros"] + 1).setValue(new Date());
    sheet.getRange(rowNumber, map["mensagem integração membros"] + 1).setValue(
      String(err && err.message ? err.message : err)
    );
  }
}