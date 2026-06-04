function createHourlyTrigger() {

  // Prevent duplicate triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "retryPendingTracking") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("retryPendingTracking")
    .timeBased()
    .everyHours(1)
    .create();

}