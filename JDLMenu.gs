/*******************************************************
 * JDLMenu.gs — メニュー追加（トリガー不要・1ボタン）
 *******************************************************/
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('JDL試算表')
      .addItem('JDL試算表作成', 'buildJDLTrialBalance')
      .addToUi();
  } catch (e) {}
}
