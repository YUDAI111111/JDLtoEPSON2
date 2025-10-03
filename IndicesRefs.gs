/**
 * IndicesRefs.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN ep_index.gs (sha256:4f9cb79f0161712a) ===== */
/** ep_index.gs — EPSON科目/補助 インデックスの薄い実装（不足関数の補完） */
var __EP_CACHE__ = { nameToCode:null, codeToName:null };
var __EP_SUB_CACHE__ = Object.create(null); // key: parentCode -> {codeToName, nameToCode}

function _ep_nameToCode(){
  _ensureEpParentIndex_();
  return __EP_CACHE__.nameToCode || {};
}
function _ep_codeToName(){
  _ensureEpParentIndex_();
  return __EP_CACHE__.codeToName || {};
}
function _ep_sub_codeToName(parentCode){
  var p = String(parentCode||'').trim();
  if(!p) return {};
  if (__EP_SUB_CACHE__[p]) return __EP_SUB_CACHE__[p].codeToName || {};
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON_SUBS);
  if(!sh || sh.getLastRow()<2){
    __EP_SUB_CACHE__[p] = {codeToName:{}, nameToCode:{}};
    return {};
  }
  var vals = sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues(); // 親コード, 補助コード, 補助名
  var c2n = {}, n2c = {};
  for (var i=0;i<vals.length;i++){
    var pc = (vals[i][0]||'').toString().trim();
    if (pc !== p) continue;
    var sc = (vals[i][1]||'').toString().trim();
    var sn = (vals[i][2]||'').toString().trim();
    if (!sn) continue;
    if (sc) c2n[sc] = sn;
    n2c[sn] = sc || '';
  }
  __EP_SUB_CACHE__[p] = {codeToName:c2n, nameToCode:n2c};
  return c2n;
}
function _ensureEpParentIndex_(){
  if (__EP_CACHE__.nameToCode && __EP_CACHE__.codeToName) return;
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
  __EP_CACHE__.nameToCode = {};
  __EP_CACHE__.codeToName = {};
  if(!sh || sh.getLastRow()<2) return;
  var last = sh.getLastRow();
  var codes = sh.getRange(2, EPSON_CHART_COLS.code,         last-1, 1).getDisplayValues();
  var names = sh.getRange(2, EPSON_CHART_COLS.name_display,  last-1, 1).getDisplayValues();
  for (var i=0;i<codes.length;i++){
    var c = (codes[i][0]||'').toString().trim();
    var n = (names[i][0]||'').toString().trim();
    if (!c || !n) continue;
    __EP_CACHE__.codeToName[c] = n;
    __EP_CACHE__.nameToCode[n] = c;
  }
}
/** ===== END ep_index.gs ===== */

/** ===== BEGIN subs_index.gs (sha256:d5f6b74d3b4eb558) ===== */
/** subs_index.gs — Epson_subs を引きやすい形に */
function _subs_indexByParent(){
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON_SUBS);
  var out = {};
  if(!sh || sh.getLastRow()<2) return out;
  var vals = sh.getRange(2,1,sh.getLastRow()-1,4).getValues(); // 親科目コード, 補助コード, 補助名, 同義語
  for (var i=0;i<vals.length;i++){
    var p=(''+vals[i][0]).trim(), c=(''+vals[i][1]).trim(), n=(''+vals[i][2]).trim();
    var syn=(''+(vals[i][3]||'')).trim();
    if(!p || !n) continue;
    if(!out[p]) out[p] = { nameToCode:{}, synonymToName:{} };
    if(c) out[p].nameToCode[n] = c;
    if(syn){
      syn.split(/\s*,\s*/).filter(function(x){return !!x;}).forEach(function(s){ out[p].synonymToName[s]=n; });
    }
  }
  return out;
}
/** ===== END subs_index.gs ===== */

/** ===== BEGIN epson_lookup.gs (sha256:4ade28fda6711a37) ===== */
/** convert_strict.gs — 全置き換え（名称欠落ゼロ化：E/F⇄G/H 双方向補完＋補助名だけキー対応） */
var __PROG_NS = 'MAPPING_CONVERT_PROGRESS_V3';

function progInit_(total){
  var o={total:Number(total||0),done:0,startMs:Date.now(),lastMs:Date.now(),
    errMapD:0,errMapC:0,errTaxD:0,errTaxC:0,canExport:false};
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(o));
}
function progBump_(delta, add){
  var raw=PropertiesService.getScriptProperties().getProperty(__PROG_NS); if(!raw) return;
  var o=JSON.parse(raw); o.done+=Number(delta||0); o.lastMs=Date.now();
  if(add){ if(add.errMapD) o.errMapD+=add.errMapD; if(add.errMapC) o.errMapC+=add.errMapC; if(add.errTaxD) o.errTaxD+=add.errTaxD; if(add.errTaxC) o.errTaxC+=add.errTaxC; }
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(o));
}
function progMarkExportable_(ok){
  var raw=PropertiesService.getScriptProperties().getProperty(__PROG_NS); if(!raw) return;
  var o=JSON.parse(raw); o.canExport=!!ok;
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(o));
}
function progSnapshot_(){
  var raw=PropertiesService.getScriptProperties().getProperty(__PROG_NS);
  if(!raw) return {total:0,done:0,percent:0,etaSec:0,err:{},canExport:false};
  var o=JSON.parse(raw), pct=o.total?Math.floor(o.done/o.total*100):0;
  var el=(o.lastMs-o.startMs)/1000, eta=(o.done>0&&o.total>o.done)?Math.max(0,Math.round(el/o.done*(o.total-o.done))):0;
  return {total:o.total,done:o.done,percent:pct,etaSec:eta,
    err:{mapD:o.errMapD,mapC:o.errMapC,taxD:o.errTaxD,taxC:o.errTaxC},canExport:!!o.canExport};
}
function showProgressSidebar_(){
  var html=HtmlService.createHtmlOutputFromFile('ProgressSidebar').setTitle('変換 進行状況');
  SpreadsheetApp.getUi().showSidebar(html);
}
function getProgressSnapshot(){ return progSnapshot_(); }

var EPSON_OFFICIAL_HEADER=[ '月種別','種類','形式','作成方法','付箋',
  '伝票日付','伝票番号','伝票摘要','枝番',
  '借方部門','借方部門名','借方科目','借方科目名','借方補助','借方補助科目名','借方金額','借方消費税コード','借方消費税業種','借方消費税税率','借方資金区分','借方任意項目１','借方任意項目２','借方インボイス情報',
  '貸方部門','貸方部門名','貸方科目','貸方科目名','貸方補助','貸方補助科目名','貸方金額','貸方消費税コード','貸方消費税業種','貸方消費税税率','貸方資金区分','貸方任意項目１','貸方任意項目２','貸方インボイス情報',
  '摘要','期日','証番号','入力マシン','入力ユーザ','入力アプリ','入力会社','入力日付'
];
function _buildConvertedHeader_(){
  var BASE=[
    '伝票日付','伝票番号',
    '借方科目','借方科目名','借方補助','借方補助科目名','借方金額','借方消費税コード','借方消費税税率',
    '貸方科目','貸方科目名','貸方補助','貸方補助科目名','貸方金額','貸方消費税コード','貸方消費税税率',
    '摘要'
  ];
  var APPEND=EPSON_OFFICIAL_HEADER.filter(function(h){return BASE.indexOf(h)===-1;});
  return {HEADER:BASE.concat(APPEND), APPEND_LEN:APPEND.length};
}

/* ===== 補完用の辞書 ===== */
// 親コード -> 親正式名
function _ep_codeToName_local_(){
  try{ if (typeof _ep_codeToName==='function') return _ep_codeToName(); }catch(_){}
  var m={}, sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
  if(!sh||sh.getLastRow()<2) return m;
  var last=sh.getLastRow();
  var codes=sh.getRange(2, EPSON_CHART_COLS.code, last-1,1).getDisplayValues();
  var names=sh.getRange(2, EPSON_CHART_COLS.name_display, last-1,1).getDisplayValues();
  for (var i=0;i<codes.length;i++){
    var c=(''+codes[i][0]).trim(), n=(''+names[i][0]).trim();
    if(c&&n) m[c]=n;
  }
  return m;
}
// 親コード -> { nameToCode, codeToName }
function _subs_index_codeToName_local_(){
  var out={}, sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON_SUBS);
  if(!sh||sh.getLastRow()<2) return out;
  var vals=sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues(); // 親コード,補助コード,補助名
  for (var i=0;i<vals.length;i++){
    var p=(''+vals[i][0]).trim(), c=(''+vals[i][1]).trim(), n=(''+vals[i][2]).trim();
    if(!p||!n) continue;
    if(!out[p]) out[p]={nameToCode:{}, codeToName:{}};
    if(c) out[p].codeToName[c]=n;
    out[p].nameToCode[n]=c||'';
  }
  return out;
}

/* 変換本体 */
function convertImportToEpson_STRICT(){
  setupOrRepairSheets();
  var ss=SpreadsheetApp.getActive(), imp=ss.getSheetByName(SHEETS.IMPORT);
  if(!imp||imp.getLastRow()<IMPORT_HEADER_ROW+1){ SpreadsheetApp.getUi().alert('1_Data_import が空です'); return; }

  showProgressSidebar_();

  var lastRow=imp.getLastRow(), lastCol=Math.max(24,imp.getLastColumn());
  var vals=imp.getRange(1,1,lastRow,lastCol).getDisplayValues();
  var start=IMPORT_HEADER_ROW+1;

  var COL_DENNO=2, COL_DATE=IMPORT_COLS.date;

  // Mapping_store 読み（A..I）
  var storeSh=ss.getSheetByName(SHEETS.MAP_STORE), mapIndex=new Map();
  if(storeSh&&storeSh.getLastRow()>=2){
    var svals=storeSh.getRange(2,1,storeSh.getLastRow()-1,9).getDisplayValues();
    for(var i=0;i<svals.length;i++){
      var jPar=(svals[i][2]||'').toString().trim(); // C
      var jSub=(svals[i][3]||'').toString().trim(); // D
      var eCod=(svals[i][4]||'').toString().trim(); // E
      var fSub=(svals[i][5]||'').toString().trim(); // F
      var gNam=(svals[i][6]||'').toString().trim(); // G
      var hNam=(svals[i][7]||'').toString().trim(); // H
      var rec={eCode:eCod,fSub:fSub,gName:gNam,hName:hNam};
      mapIndex.set(jPar+'|'+(jSub||''), rec);      // ① 親|補助
      if(!jPar && jSub) mapIndex.set('|'+jSub, rec); // ③ |補助（親名空で保存されてたケース）
      if(jSub==='[補助なし]') mapIndex.set(jPar+'|', rec); // ② 親|
    }
  }

  // 逆引き辞書（E→G / E+F→H）
  var epCodeToName=_ep_codeToName_local_();
  var subsIx=_subs_index_codeToName_local_();

  function mapTaxKnown(raw){
    var s=(raw||'').toString().trim();
    switch(s){
      case '仕　入': case '仕 入': case '仕入': return {code:'32',rate:10};
      case '売一種':                         return {code:'02',rate:10};
      case '売五種':                         return {code:'02',rate:10};
      case '非売上':                         return {code:'20',rate:0};
      case '非仕入':                         return {code:'31',rate:0};
      default:                               return null;
    }
  }
  function decideTax(parName, raw){
    var has=!!(parName||'').trim(), s=(raw||'').toString().trim();
    if(!has) return {code:'',rate:''};   // 複数行の空欄は空で出す
    if(!s)   return {code:'00',rate:0};  // 科目あり＆空欄 → 00/0%
    var k=mapTaxKnown(s); if(k) return k;
    return {code:null,rate:null};
  }

  // 伝票番号採番（B列が空は日付ごとに通番）
  var dateSeq=Object.create(null), outNo=new Array(lastRow+1).fill('');
  for(var r=start;r<=lastRow;r++){
    var d=_dateStr(vals[r-1][COL_DATE-1]), dn=(vals[r-1][COL_DENNO-1]||'').toString().trim();
    if(dn){ outNo[r]=dn; } else { var k=d||'(no-date)'; if(!(k in dateSeq)) dateSeq[k]=0; dateSeq[k]++; outNo[r]=dateSeq[k]; }
  }

  var total=Math.max(0,lastRow-IMPORT_HEADER_ROW); progInit_(total);

  // 出力ヘッダ
  var H=_buildConvertedHeader_(), HEADER=H.HEADER, APPEND_LEN=H.APPEND_LEN;
  var out=ss.getSheetByName(SHEETS.CONVERTED); out.clear();
  out.getRange(1,1,1,HEADER.length).setValues([HEADER]); out.setFrozenRows(1);

  // 税コード列を文字列表示固定（先頭ゼロ保持）
  var idxDCode=HEADER.indexOf('借方消費税コード')+1;
  var idxCCode=HEADER.indexOf('貸方消費税コード')+1;
  if(idxDCode>0) out.getRange(2,idxDCode,Math.max(1,total),1).setNumberFormat('@STRING@');
  if(idxCCode>0) out.getRange(2,idxCCode,Math.max(1,total),1).setNumberFormat('@STRING@');

  // Logs
  var logs=ss.getSheetByName(SHEETS.LOGS), lbuf=[];
  function logPush(level,proc,srcRow,side,jPar,jSub,reason,key){ lbuf.push([_now(),level,proc,srcRow,side,jPar,jSub,reason,key]); }
  function flushLogs(){ if(lbuf.length){ logs.getRange(logs.getLastRow()+1,1,lbuf.length,9).setValues(lbuf); lbuf=[]; } }

  // マッピング取得＋E/F⇄G/H 補完
  function completeByCode(rec){
    // G:親名 欠け → E から補完
    if(!rec.gName && rec.eCode){
      rec.gName = epCodeToName[rec.eCode] || rec.gName || '';
    }
    // H:補助名 欠け → E+F から補完
    if(!rec.hName && rec.eCode && rec.fSub && subsIx[rec.eCode] && subsIx[rec.eCode].codeToName){
      rec.hName = subsIx[rec.eCode].codeToName[rec.fSub] || rec.hName || '';
    }
    return rec;
  }
  function completeByName(rec){
    // E 欠け → G から逆引き
    if(!rec.eCode && rec.gName){
      try{ if (typeof _ep_nameToCode==='function'){ var m=_ep_nameToCode(); rec.eCode=m[rec.gName]||rec.eCode||''; } }catch(_){}
      if(!rec.eCode){ // ローカル逆引き（安全側）
        var sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
        if(sh&&sh.getLastRow()>=2){
          var last=sh.getLastRow();
          var names=sh.getRange(2, EPSON_CHART_COLS.name_display, last-1,1).getDisplayValues();
          var codes=sh.getRange(2, EPSON_CHART_COLS.code, last-1,1).getDisplayValues();
          for (var i=0;i<names.length;i++){
            if((names[i][0]||'').toString().trim()===rec.gName){ rec.eCode=(codes[i][0]||'').toString().trim(); break; }
          }
        }
      }
    }
    // F 欠け → E + H から逆引き
    if(!rec.fSub && rec.eCode && rec.hName && subsIx[rec.eCode] && subsIx[rec.eCode].nameToCode){
      rec.fSub = subsIx[rec.eCode].nameToCode[rec.hName] || rec.fSub || '';
    }
    return rec;
  }

  function getMap(par, sub, side, rowNo){
    if(!par && !sub) return {eCode:'',fSub:'',gName:'',hName:''};
    var rec = mapIndex.get((par||'')+'|'+(sub||'')) ||    // ① 親|補助
              mapIndex.get((par||'')+'|') ||              // ② 親|
              mapIndex.get('|'+(sub||'')) ||              // ③ |補助
              null;
    if(!rec){
      if(par){ if(side==='借方') progBump_(0,{errMapD:1}); else progBump_(0,{errMapC:1}); }
      logPush('WARN','convert', rowNo, side, (par||''), (sub||''), 'マッピング不足', (par||'')+'|'+(sub||''));
      return {eCode:'',fSub:'',gName:'',hName:''};
    }
    // 双方向補完で名称欠落をゼロ化
    rec = {eCode: String(rec.eCode||''), fSub: String(rec.fSub||''), gName: rec.gName||'', hName: rec.hName||''};
    rec = completeByCode(rec);
    rec = completeByName(rec);
    rec = completeByCode(rec); // 最後にもう一度（連鎖補完）
    return rec;
  }

  // 変換
  var outRows=[], bumpEvery=50, curRow=start;
  for(curRow=start; curRow<=lastRow; curRow++){
    var dn=String(outNo[curRow]), jDate=vals[curRow-1][COL_DATE-1];

    var dPar=(vals[curRow-1][IMPORT_COLS.debitName-1]||'').toString().trim();
    var dSub=(vals[curRow-1][IMPORT_COLS.dSubName-1]  ||'').toString().trim();
    var dAmt=Number((vals[curRow-1][IMPORT_COLS.debitAmt-1]||'0').toString().replace(/,/g,''));

    var cPar=(vals[curRow-1][IMPORT_COLS.creditName-1]||'').toString().trim();
    var cSub=(vals[curRow-1][IMPORT_COLS.cSubName-1]  ||'').toString().trim();
    var cAmt=Number((vals[curRow-1][IMPORT_COLS.creditAmt-1]||'0').toString().replace(/,/g,''));

    var dTaxRaw=(vals[curRow-1][IMPORT_COLS.dTax-1]||'').toString().trim();
    var cTaxRaw=(vals[curRow-1][IMPORT_COLS.cTax-1]||'').toString().trim();
    var memo=(vals[curRow-1][IMPORT_COLS.memo-1]||'').toString().trim();

    var dMap=getMap(dPar,dSub,'借方',curRow);
    var cMap=getMap(cPar,cSub,'貸方',curRow);

    var dTax=decideTax(dPar,dTaxRaw), cTax=decideTax(cPar,cTaxRaw);
    if(dTax.code===null){ progBump_(0,{errTaxD:1}); logPush('WARN','convert',curRow,'借方',dPar,dSub,'税区分 未知: '+dTaxRaw,dPar+'|'+dSub); dTax={code:'',rate:''}; }
    if(cTax.code===null){ progBump_(0,{errTaxC:1}); logPush('WARN','convert',curRow,'貸方',cPar,cSub,'税区分 未知: '+cTaxRaw,cPar+'|'+cSub); cTax={code:'',rate:''}; }

    // ← ここが重要：EPSON の「科目名／補助名」を必ず出す（空にならない）
    var row=[
      _dateStr(jDate), dn,
      dMap.eCode, dMap.gName, dMap.fSub, dMap.hName, (dAmt||''), (dTax.code!==''?String(dTax.code):''), (dTax.rate!==''?dTax.rate:''),
      cMap.eCode, cMap.gName, cMap.fSub, cMap.hName, (cAmt||''), (cTax.code!==''?String(cTax.code):''), (cTax.rate!==''?cTax.rate:''),
      memo
    ];
    if(H.APPEND_LEN>0) row=row.concat(new Array(H.APPEND_LEN).fill(''));
    outRows.push(row);

    var doneNow=curRow-start+1;
    if(doneNow%bumpEvery===0){ progBump_(bumpEvery,null); flushLogs(); Utilities.sleep(15); }
  }
  var rem=(lastRow-start+1)%bumpEvery; if(rem) progBump_(rem,null);
  flushLogs();

  var snap=progSnapshot_(), hasErr=(snap.err.mapD+snap.err.mapC+snap.err.taxD+snap.err.taxC)>0;
  if(hasErr){
    progMarkExportable_(false);
    _log_('INFO','convert','-','-','-','-','完了(エラーあり)：'+snap.done+'/'+snap.total,'-');
    SpreadsheetApp.getUi().alert('変換中止（エラーあり）。4_Converted には書き込みません。Logs を確認してください。');
    return;
  }
  if(outRows.length) out.getRange(2,1,outRows.length,H.HEADER.length).setValues(outRows);
  progMarkExportable_(true);

  snap=progSnapshot_();
  ss.getSheetByName(SHEETS.LOGS).appendRow([_now(),'INFO','convert','','','','','完了: '+snap.done+'/'+snap.total+'（エラーなし）','']);
  SpreadsheetApp.getUi().alert('変換完了：'+snap.done+'/'+snap.total+' 行を 4_Converted に出力しました。');
}

/** CSV（Shift_JIS, CRLF）をBase64で返す（既存のまま） */
function buildConvertedCsvBase64(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName(SHEETS.CONVERTED);
  if(!sh||sh.getLastRow()<2) throw new Error('4_Converted が空です');
  var R=sh.getLastRow(), C=sh.getLastColumn();
  var header=sh.getRange(1,1,1,C).getDisplayValues()[0];
  var data=sh.getRange(2,1,R-1,C).getDisplayValues();
  function esc(v){ var s=(v==null?'':String(v)); if(/[",\r\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"'; return s; }
  var lines=[header.map(esc).join(',')];
  for(var i=0;i<data.length;i++) lines.push(data[i].map(esc).join(','));
  var csv=lines.join('\r\n');
  var blob=Utilities.newBlob('', 'text/csv', 'converted.csv'); blob.setDataFromString(csv, 'Shift_JIS');
  return {filename:'converted.csv', base64:Utilities.base64Encode(blob.getBytes()), mime:'text/csv; charset=shift_jis'};
}
/** ===== END epson_lookup.gs ===== */

/** ===== BEGIN shim_fuzzy_rank.gs (sha256:76b87f3325d2222b) ===== */
/** shim_fuzzy_rank.gs — _rankBest が無い環境向けの最小実装（部分一致用） */
function _rankBest(name, pool) {
  if (!name || !pool || !pool.length) return '';
  var a = __shim_norm(name);
  var best = '', bestScore = -1;

  for (var i = 0; i < pool.length; i++) {
    var b = __shim_norm(pool[i]);
    var score = 0;

    if (a === b) {
      score = 100;                           // 完全一致
    } else {
      var l = __shim_lcsContig(a, b);        // 連続一致の長さ
      if (l >= 2) {
        score = 85 + Math.min(10, l - 2);    // 2文字以上の連続一致→部分一致
      } else {
        var diff = Math.abs(a.length - b.length);
        score = Math.max(0, 70 - diff * 5);  // 長さ近似の簡易スコア
      }
    }

    if (score > bestScore) { bestScore = score; best = pool[i]; }
  }
  return (bestScore >= 80) ? best : '';      // 80点以上で採用（＝部分一致扱い）
}

/* 以下は _rankBest 内部用のローカル関数（既存関数と名前が被らないよう shim 接頭辞） */
function __shim_norm(s) {
  return (s || '').toString()
    .normalize('NFKC')
    .replace(/[ 　\t]+/g, '')
    .replace(/[()（）［］【】]/g, '')
    .trim();
}

function __shim_lcsContig(a, b) {
  var n = a.length, m = b.length;
  if (!n || !m) return 0;
  var prev = new Array(m + 1).fill(0), best = 0;
  for (var i = 1; i <= n; i++) {
    var curr = new Array(m + 1).fill(0);
    for (var j = 1; j <= m; j++) {
      if (a[i - 1] === b[j - 1]) {
        curr[j] = prev[j - 1] + 1;
        if (curr[j] > best) best = curr[j];
      }
    }
    prev = curr;
  }
  return best;
}
/** ===== END shim_fuzzy_rank.gs ===== */

