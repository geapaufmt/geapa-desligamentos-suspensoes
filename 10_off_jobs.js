/***************************************
 * 10_off_jobs.gs
 * Jobs diarios e wrapper legado.
 ***************************************/

/**
 * Wrapper legado mantido para compatibilidade externa.
 */
function off_processQueue(e) {
  return off_processDailyLifecycleQueue(e);
}

/**
 * Job diario unico para iniciar suspensoes, efetivar retornos e executar desligamentos agendados.
 */
function off_processDailyLifecycleQueue(e) {
  const runId = off_newRunId_();
  return off_runControlledFlow_(
    OFF_OPS.FLOWS.JOB_DIARIO_FILAS,
    OFF_OPS.CAPABILITIES.SYNC,
    {
      executionType: off_getExecutionTypeFromEvent_(e),
      runId: runId,
      origin: 'DAILY_JOB',
    },
    function(control) {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) {
        off_logWarn_(runId, 'off_processDailyLifecycleQueue: lock nao obtido');
        return null;
      }

      try {
        off_ensureOperationalSheets_();

        const summary = {
          dryRun: control.dryRun === true,
          suspensionsStarted: off_processDueSuspensionStarts_(runId, {
            executionType: control.executionType,
          }),
          suspensionReturns: off_processDueSuspensionReturns_(runId, {
            executionType: control.executionType,
          }),
          scheduledDismissals: off_processDueScheduledDismissals_(runId, {
            executionType: control.executionType,
          }),
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
  );
}
