/** JDLMenu.gs — 上書き版（ScriptApp不使用・UIセーフ・onOpenから呼ぶだけ） */
function addJDLMenu_() {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    return;
  }
  try {
    ui.createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (err) {
    try {
      ui.createMenu('JDL試算表(2)')
        .addItem('JDL試算表作成', 'buildJDLTrialBalance')
        .addToUi();
    } catch (e2) {}
  }
}
