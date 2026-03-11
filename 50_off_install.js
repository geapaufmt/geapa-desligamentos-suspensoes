function off_installTriggers() {
  const reg = GEAPA_CORE.coreGetRegistry();
  const ref = reg.OFFBOARD;

  if (!ref?.id) throw new Error("Registry OFFBOARD.id está vazio no GEAPA_CORE.");

  // Remove triggers antigos desse módulo
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "off_handleFormSubmit" || fn === "off_handleEdit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("off_handleFormSubmit")
    .forSpreadsheet(ref.id)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("off_handleEdit")
    .forSpreadsheet(ref.id)
    .onEdit()
    .create();
}