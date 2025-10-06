/*******************************************************
 * JDLMenu.gs — メニュー追加（2通り）
 * A) ensureJDLMenuTrigger(): インストール型トリガー（恒常表示）
 *    ※ 要 scope: https://www.googleapis.com/auth/script.scriptapp
 * B) addJDLMenuOnce(): セッション限定の即時メニュー追加（トリガー不要）
 *******************************************************/
function onOpen_JDLMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (_) {}
}

function ensureJDLMenuTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function(t){
    return t.getHandlerFunction && t.getHandlerFunction() === 'onOpen_JDLMenu_' &&
           t.getEventType && t.getEventType() === ScriptApp.EventType.ON_OPEN;
  });
  if (!exists) {
    ScriptApp.newTrigger('onOpen_JDLMenu_')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }
  SpreadsheetApp.getActive().toast('JDLメニューを恒常化しました（次回以降に表示）','JDL試算表',5);
}

function addJDLMenuOnce() {
  SpreadsheetApp.getUi()
    .createMenu('JDL試算表')
    .addItem('JDL試算表作成', 'buildJDLTrialBalance')
    .addToUi();
  SpreadsheetApp.getActive().toast('このセッションだけメニューを追加しました','JDL試算表',5);
}
