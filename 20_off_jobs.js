/***************************************
 * 20_off_jobs.gs
 * Jobs para triggers
 ***************************************/
function off_processQueue() {
  const runId = GEAPA_CORE.coreRunId();
  GEAPA_CORE.coreLogInfo(runId, 'off_processQueue: INÍCIO');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    GEAPA_CORE.coreLogWarn(runId, 'off_processQueue: lock não obtido');
    return;
  }

  try {
    const { sh, headerMap, rows, startRow } = off_readResponsesTable_();
    if (!rows.length) {
      GEAPA_CORE.coreLogInfo(runId, 'off_processQueue: sem linhas');
      return;
    }

    // colunas (1-based)
    const cType = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.TYPE);
    const cName = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.NAME);
    const cRga = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.RGA);
    const cEmail = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.EMAIL);
    const cNotified = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.NOTIFIED_SECS);
    const cApproved = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.APPROVED);
    const cSentFinal = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.SENT_FINAL);
    const cDeferDt = GEAPA_CORE.coreGetCol(headerMap, OFF_CFG.RESP.DEFER_DATE);

    // valida mínimos
    const required = { cType, cName, cEmail, cNotified, cApproved, cSentFinal };
    const missing = Object.entries(required).filter(([, v]) => !v).map(([k]) => k);
    if (missing.length) {
      throw new Error('OFF: cabeçalhos obrigatórios ausentes (mapeamento falhou): ' + missing.join(', '));
    }

    const secs = off_getSecretaryEmails_();

    let notifiedCount = 0;
    let finalSentCount = 0;

    rows.forEach((row, idx) => {
      const sheetRow = startRow + idx;

      const tipo = String(row[cType - 1] || '').trim();
      const nome = String(row[cName - 1] || '').trim();
      const rga = cRga ? String(row[cRga - 1] || '').trim() : '';
      const email = String(row[cEmail - 1] || '').trim();
      const notified = String(row[cNotified - 1] || '').trim().toUpperCase();
      const approved = String(row[cApproved - 1] || '').trim().toUpperCase();
      const sentFinal = String(row[cSentFinal - 1] || '').trim().toUpperCase();

      // 1) Notificar secretários (uma vez)
      if (notified !== OFF_CFG.VALUES.YES) {
        if (secs.length) {
          off_notifySecretaries_(secs, { tipo, nome, rga, email });
          sh.getRange(sheetRow, cNotified).setValue(OFF_CFG.VALUES.YES);
          notifiedCount++;
        } else {
          GEAPA_CORE.coreLogWarn(runId, 'OFF: nenhum e-mail de secretário encontrado', {
            sheetRow,
            nome,
            rga
          });
        }
      }

      // 2) Retaguarda de integração com membros
      if (approved === OFF_CFG.VALUES.DECISAO_DEFERIDO) {
        offboard_processMembersIntegration_(sh, sheetRow);
      }
    });

    GEAPA_CORE.coreLogInfo(runId, 'off_processQueue: FIM OK', {
      notifiedCount,
      finalSentCount
    });

  } catch (e) {
    GEAPA_CORE.coreLogError(runId, 'off_processQueue: ERRO', {
      err: String(e),
      stack: e && e.stack
    });
    throw e;
  } finally {
    lock.releaseLock();
  }
}