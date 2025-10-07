/*******************************************************
 * JDLMenu.gs — メニュー追加
 *******************************************************/
function onOpen_JDLMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (_) {}
}
function addJDLMenuOnce() {
  SpreadsheetApp.getUi()
    .createMenu('JDL試算表')
    .addItem('JDL試算表作成', 'buildJDLTrialBalance')
    .addToUi();
  SpreadsheetApp.getActive().toast('このセッションだけメニューを追加しました','JDL試算表',5);
}
