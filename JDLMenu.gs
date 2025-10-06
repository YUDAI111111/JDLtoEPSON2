/*******************************************************
 * JDLMenu.gs — 既存 onOpen に干渉せず「JDL試算表作成」を追加
 * 使い方：一度だけ ensureJDLMenuTrigger() を実行（権限許可）
 *******************************************************/
function onOpen_JDLMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (_) {}
}

/** インストール型トリガーを作成（重複は作らない） */
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
  SpreadsheetApp.getActive().toast('JDLメニューを追加しました（次回以降に表示）','JDL試算表',5);
}
