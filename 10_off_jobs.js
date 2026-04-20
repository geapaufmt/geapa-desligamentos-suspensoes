/***************************************
 * 10_off_jobs.gs
 * Jobs diarios e wrapper legado.
 ***************************************/

/**
 * Wrapper legado mantido para compatibilidade externa.
 */
function off_processQueue() {
  return off_processDailyLifecycleQueue();
}

/**
 * Job diario unico para iniciar suspensoes, efetivar retornos e executar desligamentos agendados.
 */
function off_processDailyLifecycleQueue() {
  const runId = off_newRunId_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    off_logWarn_(runId, 'off_processDailyLifecycleQueue: lock nao obtido');
    return null;
  }

  try {
    off_ensureOperationalSheets_();

    const summary = {
      suspensionsStarted: off_processDueSuspensionStarts_(runId),
      suspensionReturns: off_processDueSuspensionReturns_(runId),
      scheduledDismissals: off_processDueScheduledDismissals_(runId),
    };

    off_logInfo_(runId, 'off_processDailyLifecycleQueue: fim ok', summary);
    return summary;
  } catch (err) {
    off_logError_(runId, 'off_processDailyLifecycleQueue: erro', {
      err: String(err),
      stack: err && err.stack,
    });
    throw err;
  } finally {
    lock.releaseLock();
  }
}
