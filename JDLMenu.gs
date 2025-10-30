/** JDLMenu_fix.gs — UI-safe menu patch (minimal, non-breaking) */
function _safeUi_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}
function addJDLMenuOnce() {
  var ss = SpreadsheetApp.getActive();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === 'onOpenJDL') return;
  }
  ScriptApp.newTrigger('onOpenJDL').forSpreadsheet(ss).onOpen().create();
}
function onOpenJDL(e) {
  var ui = _safeUi_();
  if (!ui) return;
  addJDLMenu_UI_(ui);
}
function addJDLMenu_UI_(ui) {
  try {
    ui.createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (err) {
    try {
      ui.createMenu('JDL試算表(2)')
        .addItem('JDL試算表作成', 'buildJDLTrialBalance')
        .addToUi();
    } catch (e) {}
  }
}
