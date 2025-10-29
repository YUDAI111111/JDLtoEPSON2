/*******************************************************
 * JDLMenu.gs — 衝突回避版（このファイルでは onOpen を定義しない）
 * 使い方：Core.gs 側の onOpen() の末尾で addJDLMenu_(); を1行呼び出してください。
 *******************************************************/

/** メニュー追加本体（トリガー不要・UIのみ） */
function addJDLMenu_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('JDL試算表')
    .addItem('JDL試算表作成', 'buildJDLTrialBalance') // FinancialStatements.gs のエントリを呼ぶ
    .addToUi();
}

/** 手動追加（デバッグ用）：エディタから実行すれば即メニューが出ます */
function addJDLMenuOnce() {
  addJDLMenu_();
  SpreadsheetApp.getActive().toast('メニューを追加しました','JDL試算表',5);
}
