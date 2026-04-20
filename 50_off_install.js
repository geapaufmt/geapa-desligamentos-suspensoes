/**
 * Reinstala os triggers do modulo de forma segura e idempotente.
 */
function off_installTriggers() {
  const ref = off_getRegistryRef_(OFF_KEYS.RESPONSES, true);
  if (!ref.id) {
    throw new Error('Registry OFFBOARD_RESPONSES.id esta vazio no GEAPA_CORE.');
  }

  off_ensureOperationalSheets_();
  if (typeof off_tryApplyOperationalSheetsUx_ === 'function') {
    off_tryApplyOperationalSheetsUx_();
  }

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const fn = trigger.getHandlerFunction();
    if (
      fn === 'off_handleFormSubmit' ||
      fn === 'off_onEditDecision' ||
      fn === 'off_processQueue' ||
      fn === 'off_processDailyLifecycleQueue'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('off_handleFormSubmit')
    .forSpreadsheet(ref.id)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('off_onEditDecision')
    .forSpreadsheet(ref.id)
    .onEdit()
    .create();

  ScriptApp.newTrigger('off_processDailyLifecycleQueue')
    .timeBased()
    .everyDays(1)
    .atHour(OFF_CFG.DAILY_TRIGGER_HOUR)
    .create();
}
