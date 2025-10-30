/** ToastGuard.gs — UI-safe toast helper (no-op if UI missing) */
function toast_(m) {
  try { SpreadsheetApp.getActive().toast(m, 'JDL試算表', 5); } catch (e) {}
}
