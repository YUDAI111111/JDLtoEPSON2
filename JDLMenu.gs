/*******************************************************
 * JDLMenu.gs — メニュー生成（UIセーフ／トリガー不要）
 *******************************************************/
function addJDLMenu_() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { return; }
  try {
    ui.createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (e1) {
    try {
      ui.createMenu('JDL試算表(2)')
        .addItem('JDL試算表作成', 'buildJDLTrialBalance')
        .addToUi();
    } catch (e2) {}
  }
}

function onOpen(e) {
  try { addJDLMenu_(); } catch (err) {
    try { SpreadsheetApp.getActive().toast('メニュー初期化失敗: ' + String(err).slice(0,120), 'JDL試算表', 5); } catch (_) {}
  }
}

function manualAddMenu() { addJDLMenu_(); }
