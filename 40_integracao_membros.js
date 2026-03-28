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

  function getCell(candidates) {
    const idx = offboard_findHeaderIndex_(map, candidates);
    return idx >= 0 ? row[idx] : "";
  }

  function setCellIfPresent(candidates, value) {
    const idx = offboard_findHeaderIndex_(map, candidates);
    if (idx >= 0) {
      sheet.getRange(rowNumber, idx + 1).setValue(value);
      return true;
    }
    return false;
  }

  const requestType = String(getCell([OFF_CFG.RESP.TYPE, "Quero solicitar:"]) || "").trim();
  const leaveTiming = String(getCell(["Desejo me desligar:"]) || "").trim();
  const decision = String(getCell([OFF_CFG.RESP.APPROVED, "Decisao da Diretoria", "Decisão da Diretoria"]) || "").trim().toUpperCase();
  const finalEmailSent = String(getCell([OFF_CFG.RESP.SENT_FINAL, "E-mail de desligamento enviado?"]) || "").trim().toUpperCase();
  const alreadyProcessed = String(getCell(["Integracao membros processada?", "Integração membros processada?"]) || "").trim().toUpperCase();

  if (requestType !== "Desligamento") return;
  if (leaveTiming !== "Imediatamente") return;
  if (decision !== "DEFERIDO") return;
  if (finalEmailSent !== "SIM") return;
  if (alreadyProcessed === "SIM") return;

  try {
    const payload = {
      requestRowNumber: rowNumber,
      requestTimestamp: getCell(["Carimbo de data/hora", "Timestamp"]),
      memberName: getCell([OFF_CFG.RESP.NAME, "Nome Completo"]),
      memberEmail: getCell([OFF_CFG.RESP.EMAIL, "Endereco de e-mail", "E-mail", "Email"]),
      memberRga: getCell([OFF_CFG.RESP.RGA, "RGA"]),
      requestType: getCell([OFF_CFG.RESP.TYPE, "Quero solicitar:"]),
      leaveTiming: getCell(["Desejo me desligar:"]),
      reason: getCell(["Informe o motivo do pedido de desligamento:"]),
      approvedAt: getCell([OFF_CFG.RESP.DEFER_DATE, "Data/Hora do deferimento"]),
      decision: getCell([OFF_CFG.RESP.APPROVED, "Decisao da Diretoria", "Decisão da Diretoria"]),
      finalEmailSent: getCell([OFF_CFG.RESP.SENT_FINAL, "E-mail de desligamento enviado?"])
    };

    const result = GEAPA_MEMBROS.members_offboardApprovedImmediateExit(payload);

    setCellIfPresent(["Integracao membros processada?", "Integração membros processada?"], "SIM");
    setCellIfPresent(["Data/Hora integracao membros", "Data/Hora integração membros"], new Date());
    setCellIfPresent(
      ["Mensagem integracao membros", "Mensagem integração membros"],
      result.duplicatedHistory
        ? "Registro já existia em MEMBERS_HIST; membro removido de MEMBERS_ATUAIS."
        : "Membro movido com sucesso de MEMBERS_ATUAIS para MEMBERS_HIST."
    );

  } catch (err) {
    setCellIfPresent(["Integracao membros processada?", "Integração membros processada?"], "ERRO");
    setCellIfPresent(["Data/Hora integracao membros", "Data/Hora integração membros"], new Date());
    setCellIfPresent(
      ["Mensagem integracao membros", "Mensagem integração membros"],
      String(err && err.message ? err.message : err)
    );
  }
}