function createEditTrigger() {

  ScriptApp.newTrigger("triggerImmediate")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

}