/**
 * Core.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN utils.gs (sha256:8fc4d7233db05131) ===== */
/** utils.gs */
function _now(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function _dateStr(v){
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  var s=String(v).trim(), m=s.match(/^(\d{4})[-\/]?(\d{1,2})[-\/]?(\d{1,2})$/);
  if (m) return m[1]+'/'+('0'+m[2]).slice(-2)+'/'+('0'+m[3]).slice(-2);
  return s;
}

function _statusExact(jName, epName){
  jName = (jName||'').toString().trim();
  epName= (epName||'').toString().trim();
  if (!jName && !epName) return '未選択';
  if (jName && epName && jName === epName) return '完全一致';
  return '不一致';
}

function _paintRowsFromMatrix_(sheet, startRow, rows) {
  if (!rows || !rows.length) return;
  var childRow     = (typeof COLORS !== 'undefined' ? COLORS.subRow       : '#E8F5E9');
  var statusBlue   = (typeof COLORS !== 'undefined' ? COLORS.statusBlue   : '#BBDEFB');
  var statusYellow = (typeof COLORS !== 'undefined' ? COLORS.statusYellow : '#FFF9C4');
  var statusRed    = (typeof COLORS !== 'undefined' ? COLORS.statusRed    : '#FFCDD2');

  var r0 = startRow;

  var childBg = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var hasChild = (row[1] && String(row[1]).trim()) || (row[3] && String(row[3]).trim());
    childBg.push([ hasChild ? childRow : null ]);
  }
  sheet.getRange(r0, 1, rows.length, 8).setBackgrounds(
    childBg.map(function(c){ return Array(8).fill(c[0]); })
  );

  var statusBg = [];
  for (var j = 0; j < rows.length; j++) {
    var s = (rows[j][8] || '').toString().trim();
    var color = null;
    if (typeof STATUS !== 'undefined') {
      if (s === STATUS.MATCH)        color = statusBlue;
      else if (s === STATUS.PARTIAL) color = statusYellow;
      else if (s === STATUS.MISMATCH) color = statusRed;
    } else {
      if (s === '完全一致')      color = statusBlue;
      else if (s === '部分一致') color = statusYellow;
      else if (s === '不一致')   color = statusRed;
    }
    statusBg.push([color]);
  }
  sheet.getRange(r0, 9, rows.length, 1).setBackgrounds(statusBg);
}

function _paintSingleRow_(sh, r){
  var vals = sh.getRange(r,1,1,9).getValues()[0];
  _paintRowsFromMatrix_(sh, r, [vals]);
}

function _log_(level, proc, srcRow, side, jPar, jSub, reason, key){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.LOGS) || ss.insertSheet(SHEETS.LOGS);
  if (sh.getLastRow()===0){
    sh.getRange(1,1,1,9).setValues([['時刻','レベル','処理','元行','借貸','親科目','補助科目','理由','キー']]);
    sh.setFrozenRows(1);
  }
  sh.appendRow([_now(), level, proc, srcRow, side, jPar, jSub, reason, key]);
}

/** ===== 進捗：エラーも含めた「処理済み件数」を％化 ===== */
var _PROG_NS_ = 'convert_progress';

function _progressReset_(total) {
  PropertiesService.getUserProperties().setProperty(_PROG_NS_, JSON.stringify({
    phase: 'running',
    total: total || 0,
    processed: 0,       // エラー含めて処理した件数
    success: 0,         // 正常変換できた件数
    error: 0,           // エラー（行）件数
    started: Date.now(),
    etaSec: 0,
    message: ''
  }));
}

/**
 * processed/success/error を都度上書き
 * pct は processed/total で計算される（＝エラーでも進む）
 */
function _progressHit_(processed, total, startedMs, success, error) {
  var now = Date.now();
  var elapsed = Math.max(1, (now - startedMs) / 1000);
  var rate = processed > 0 ? processed / elapsed : 0; // rows/sec
  var remain = (rate > 0 && total > processed) ? Math.max(0, Math.round((total - processed) / rate)) : 0;
  PropertiesService.getUserProperties().setProperty(_PROG_NS_, JSON.stringify({
    phase: 'running',
    total: total,
    processed: processed,
    success: success || 0,
    error: error || 0,
    started: startedMs,
    etaSec: remain,
    message: ''
  }));
}

function _progressFinish_(ok, message) {
  var p = PropertiesService.getUserProperties().getProperty(_PROG_NS_);
  var cur = p ? JSON.parse(p) : {};
  cur.phase = ok ? 'done' : 'error';
  cur.message = message || '';
  PropertiesService.getUserProperties().setProperty(_PROG_NS_, JSON.stringify(cur));
}

function progressSnapshot() {
  var p = PropertiesService.getUserProperties().getProperty(_PROG_NS_);
  return p ? JSON.parse(p) : { phase:'idle', total:0, processed:0, success:0, error:0, etaSec:0, message:'' };
}

/** サイドバー（エラー含む処理済み件数で％表示・成功/エラー内訳も表示） */
function _showProgressSidebar_() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px system-ui, -apple-system, Segoe UI, Roboto, sans-serif;padding:12px;min-width:300px">' +
    '<div style="font-weight:600;margin-bottom:6px">変換の進捗</div>' +
    '<div id="label" style="opacity:.8">準備中…</div>' +
    '<div style="height:10px;background:#eee;border-radius:6px;margin:10px 0;overflow:hidden">' +
      '<div id="bar" style="height:100%;width:0%;background:#1a73e8;transition:width .25s"></div>' +
    '</div>' +
    '<div id="detail" style="color:#444;white-space:pre-line"></div>' +
    '<script>' +
      'function fmt(sec){sec=Math.max(0,Math.floor(sec));var m=Math.floor(sec/60),s=sec%60;return (m>0?m+\"分\":\"\")+s+\"秒\";}' +
      'async function tick(){' +
        'google.script.run.withSuccessHandler(function(p){' +
          'if(!p) return;' +
          'var total=p.total||0, done=p.processed||0;' +   // ← processed を％の分子に
          'var pct = total? Math.floor(done*100/total): (p.phase===\"done\"?100:0);' +
          'document.getElementById(\"bar\").style.width=pct+\"%\";' +
          'document.getElementById(\"label\").textContent = (p.phase===\"running\"?\"変換中\":\"状態\") + \" … \" + pct + \"%\";' +
          'var det=\"\";' +
          'if(total){ det += done+\"/\"+total+\"  残り約 \"+fmt(p.etaSec); }' +
          'det += \"\\n成功:\"+(p.success||0)+\" / エラー:\"+(p.error||0);' +
          'if(p.message) det += \"\\n\"+p.message;' +
          'document.getElementById(\"detail\").textContent = det;' +
          'if(p.phase===\"done\"||p.phase===\"error\"){clearInterval(window._iv);}' +
        '}).progressSnapshot();' +
      '}' +
      'window._iv=setInterval(tick,800);tick();' +
    '</script>' +
    '</div>'
  ).setTitle('進捗');
  SpreadsheetApp.getUi().showSidebar(html);
}
/** ===== END utils.gs ===== */

/** ===== BEGIN debug.gs (sha256:39f82d65bcfe957f) ===== */

/** debug.gs */
function buildDebugSubjects(){
  setupOrRepairSheets();
  var ss = SpreadsheetApp.getActive();
  var imp = ss.getSheetByName(SHEETS.IMPORT);
  if(!imp) throw new Error('1_Data_import がありません');
  var vals = imp.getDataRange().getValues();
  var start = IMPORT_HEADER_ROW + 1;
  var seen = new Map();
  function addRec(jCode,jName,jSub,jSubName){
    jCode=(jCode||'').toString().trim();
    jName=(jName||'').toString().trim();
    jSub =(jSub ||'').toString().trim();
    jSubName=(jSubName||'').toString().trim();
    if(!jName) return;
    var key=jCode+'|'+jSub+'|'+jName;
    if(!seen.has(key)){ seen.set(key,{jCode:jCode,jName:jName,jSub:jSub,jSubName:jSubName,cnt:0}); }
    seen.get(key).cnt++;
  }
  for(var r=start-1;r<vals.length;r++){
    var row=vals[r];
    addRec(row[IMPORT_COLS.debitCode-1], row[IMPORT_COLS.debitName-1], row[IMPORT_COLS.dSubCode-1], row[IMPORT_COLS.dSubName-1]);
    addRec(row[IMPORT_COLS.creditCode-1],row[IMPORT_COLS.creditName-1],row[IMPORT_COLS.cSubCode-1], row[IMPORT_COLS.cSubName-1]);
  }
  var list = Array.from(seen.values());
  list.sort(function(a,b){ return (a.jCode||'').localeCompare(b.jCode||'') || (a.jSub||'').localeCompare(b.jSub||'') || (a.jName||'').localeCompare(b.jName||''); });
  var out = ss.getSheetByName(SHEETS.DEBUG_SUBJ) || ss.insertSheet(SHEETS.DEBUG_SUBJ);
  out.clear();
  out.getRange(1,1,1,6).setValues([['区分','親科目名','JDL科目コード','補助科目名','JDL補助コード','件数']]);
  out.setFrozenRows(1);
  if(list.length){
    out.getRange(2,1,list.length,6).setValues(list.map(function(x){ return ['', x.jName, x.jCode, x.jSubName, x.jSub, x.cnt]; }));
  }
}
/** ===== END debug.gs ===== */

/** ===== BEGIN setup.gs (sha256:0a9b1e4c745e6c9e) ===== */

/** setup.gs */
function setupOrRepairSheets(){
  var ss = SpreadsheetApp.getActive();
  function ensure(name, header, freeze){
    var sh = ss.getSheetByName(name);
    if(!sh){ sh = ss.insertSheet(name); }
    if(header && header.length && (!sh.getLastRow())){
      sh.getRange(1,1,1,header.length).setValues([header]);
      if(freeze) sh.setFrozenRows(1);
    }
    return sh;
  }
  ensure(SHEETS.IMPORT, [], false);
  ensure(SHEETS.MAPPING, [
    'JDL科目コード','JDL補助コード','JDL科目名','JDL補助科目名',
    'EPSON科目コード','EPSON補助コード','EPSON科目名（選択）','EPSON補助科目名','状態'
  ], true);
  ensure(SHEETS.MAP_STORE, [
    'JDL科目名','JDL補助科目名','EPSON科目コード','EPSON科目名','EPSON補助コード','EPSON補助名','状態','キー(内部)'
  ], true);
  ensure(SHEETS.CONVERTED, [], true);
  ensure(SHEETS.DEBUG_SUBJ, ['区分','親科目名','JDL科目コード','補助科目名','JDL補助コード','件数'], true);

  ensure(SHEETS.EPSON, ['コード','正 式 科 目 名','（他列任意）'], true);
  ensure(SHEETS.EPSON_SUBS, ['親科目コード','補助コード','補助名','同義語(カンマ区切り)'], true);

  var dv = ss.getSheetByName(SHEETS.DV_NAMES); if(!dv) ss.insertSheet(SHEETS.DV_NAMES);
  var dvs = ss.getSheetByName(SHEETS.DV_SUBS); if(!dvs) ss.insertSheet(SHEETS.DV_SUBS);
  ss.getSheetByName(SHEETS.DV_NAMES).hideSheet();
  ss.getSheetByName(SHEETS.DV_SUBS).hideSheet();

  var logs = ss.getSheetByName(SHEETS.LOGS);
  if(!logs){
    logs = ss.insertSheet(SHEETS.LOGS);
    logs.getRange(1,1,1,9).setValues([['時刻','レベル','処理','元行','借貸','親科目','補助科目','理由','キー']]);
    logs.setFrozenRows(1);
  }
  return true;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('仕訳変換ツール')
    .addItem('0 初期化/修復（必須シート作成）', 'setupOrRepairSheets')
    .addSeparator()
    .addItem('① マッピング（固定ヘッダ版）', 'buildMappingGrid')
    .addItem('② マッピング保存', 'saveMappingStore')
    .addItem('③ 仕訳をEPSON変換（厳格）', 'convertImportToEpson_STRICT')
    .addSeparator()
    .addItem('デバッグ（科目親子集計）', 'buildDebugSubjects')
    .addToUi()
    .addJDLMenu_();
}
/** ===== END setup.gs ===== */

