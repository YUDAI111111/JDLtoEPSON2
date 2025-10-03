/**
 * ConvertEngine.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN convert_strict.gs (sha256:560983ec44b3a727) ===== */
/** convert_strict.gs — コード優先＋銀行特例＋買掛/消耗の親コード強制補完＋名称補完（最終確定）
 * 目的：4_Converted の「借方科目名／貸方科目名」が空になる問題を解消。
 * 対応：Mapping_store の G/H が空でも、E/F のコードから EPSON台帳で **科目名／補助名を自動補完**。
 * 方針：
 *  - 照合は A|B（JDL親コード|補助コード）優先。名称フォールバックはしない。
 *  - 買掛金(3141)/消耗品費(8621) は親名から親コードを強制補完して A|B 参照。
 *  - 銀行/積立は特例オーバーライド（普通預金/積立預金のみ）。
 *  - 出力は E/F/G/H（EPSON）だが、G/H が空なら **EPSON台帳から補完**（ここが今回の要修正）。
 */

var __PROG_NS = 'MAPPING_CONVERT_PROGRESS_V3';

/* ===== Progress ===== */
function progInit_(total) {
  var obj = { total:Number(total||0), done:0, startMs:Date.now(), lastMs:Date.now(),
    errMapD:0, errMapC:0, errTaxD:0, errTaxC:0, canExport:false };
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(obj));
}
function progBump_(deltaDone, add) {
  var raw = PropertiesService.getScriptProperties().getProperty(__PROG_NS);
  if(!raw) return;
  var o = JSON.parse(raw);
  o.done += (deltaDone||0);
  if(add){ if(add.errMapD) o.errMapD+=add.errMapD; if(add.errMapC) o.errMapC+=add.errMapC; if(add.errTaxD) o.errTaxD+=add.errTaxD; if(add.errTaxC) o.errTaxC+=add.errTaxC; }
  o.lastMs = Date.now();
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(o));
}
function progMarkExportable_(ok){
  var raw = PropertiesService.getScriptProperties().getProperty(__PROG_NS); if(!raw) return;
  var o = JSON.parse(raw); o.canExport = !!ok;
  PropertiesService.getScriptProperties().setProperty(__PROG_NS, JSON.stringify(o));
}
function progSnapshot_(){
  var raw = PropertiesService.getScriptProperties().getProperty(__PROG_NS);
  if(!raw) return { total:0, done:0, percent:0, etaSec:0, err:{}, canExport:false };
  var o = JSON.parse(raw), pct = o.total ? Math.floor(o.done/o.total*100) : 0;
  var el=(o.lastMs-o.startMs)/1000, eta=(o.done>0&&o.total>o.done)?Math.max(0,Math.round(el/o.done*(o.total-o.done))):0;
  return { total:o.total, done:o.done, percent:pct, etaSec:eta,
    err:{mapD:o.errMapD,mapC:o.errMapC,taxD:o.errTaxD,taxC:o.errTaxC}, canExport:!!o.canExport };
}
function showProgressSidebar_(){
  var html=HtmlService.createHtmlOutputFromFile('ProgressSidebar').setTitle('変換 進行状況');
  SpreadsheetApp.getUi().showSidebar(html);
}
function getProgressSnapshot(){ return progSnapshot_(); }

/* ===== Header ===== */
var EPSON_OFFICIAL_HEADER = [
  '月種別','種類','形式','作成方法','付箋',
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
  return {HEADER:BASE.concat(APPEND), APPEND:APPEND};
}

/* ===== 税区分 ===== */
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
function decideTax_(parName, raw){
  var has=!!(parName||'').trim(), s=(raw||'').toString().trim();
  if(!has) return {code:'',rate:''};
  if(!s)   return {code:'00',rate:0};
  var k=mapTaxKnown(s); if(k) return k;
  return {code:null,rate:null};
}

/* ===== EPSON 台帳インデックス ===== */
function _ep_codeToName_local_(){
  var m={}, sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
  if(!sh||sh.getLastRow()<2) return m;
  var last=sh.getLastRow();
  var codes=sh.getRange(2, EPSON_CHART_COLS.code,         last-1, 1).getDisplayValues();
  var names=sh.getRange(2, EPSON_CHART_COLS.name_display, last-1, 1).getDisplayValues();
  for (var i=0;i<codes.length;i++){
    var c=(''+codes[i][0]).trim(), n=(''+names[i][0]).trim();
    if(c&&n) m[c]=n;
  }
  return m;
}
function _subs_index_codeToName_local_(){
  var out={}, sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON_SUBS);
  if(!sh||sh.getLastRow()<2) return out;
  var vals=sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues(); // 親コード,補助コード,補助名
  for (var i=0;i<vals.length;i++){
    var p=(''+vals[i][0]).trim(), c=(''+vals[i][1]).trim(), n=(''+vals[i][2]).trim();
    if(!p||!n) continue;
    if(!out[p]) out[p]={codeToName:{}, nameToCode:{}};
    if(c) out[p].codeToName[c]=n;
    out[p].nameToCode[n]=c||'';
  }
  return out;
}
function _completeNamesFromCodes_(rec, epCodeToName, subsIx){
  // rec: {eCode,fSub,gName,hName}
  if(!rec) return rec;
  if(!rec.gName && rec.eCode){
    rec.gName = epCodeToName[rec.eCode] || rec.gName || '';
  }
  if(!rec.hName && rec.eCode && rec.fSub && subsIx[rec.eCode] && subsIx[rec.eCode].codeToName){
    rec.hName = subsIx[rec.eCode].codeToName[rec.fSub] || rec.hName || '';
  }
  return rec;
}

/* ===== 銀行系オーバーライド ===== */
function _bankOverride_(parCode, subCode, parName, subName){
  // 普通預金(1312)
  if (parCode==='1312' || (parName||'')==='普通預金'){
    if ((subCode==='11') || (subName||'').indexOf('八十二銀行')>=0){
      return {eCode:'116', fSub:'', gName:'普通八十二', hName:''};
    }
    if ((subCode==='12') || (subName||'').indexOf('長野県信用組合')>=0){
      return {eCode:'115', fSub:'', gName:'普通長野県信', hName:''};
    }
    if ((subName||'').indexOf('長野信用金庫須坂2')>=0){
      return {eCode:'117', fSub:'', gName:'普通長野信金2', hName:''};
    }
    if (!subCode && (!subName || subName==='[補助なし]')){
      return {eCode:'114', fSub:'', gName:'普通長野信金', hName:''};
    }
  }
  // 積立預金(1421)
  if (parCode==='1421' || (parName||'')==='積立預金'){
    if ((subName||'').indexOf('長野信用金庫1')>=0){
      return {eCode:'125', fSub:'1', gName:'定積長野信金', hName:'しんきん'};
    }
    if (subCode==='14' || (subName||'').indexOf('長野県信用組合2')>=0){
      return {eCode:'126', fSub:'1', gName:'定積長野県信', hName:'けんしん'};
    }
  }
  return null;
}

/* ===== Util ===== */
function _normName_(s){ return (s||'').toString().replace(/\s+/g,'').trim(); }
function _forceParentCodeIfKaidakaOrShomohin_(parCode, parName){
  if (parCode) return parCode;
  var n = _normName_(parName);
  if (n==='買掛金') return '3141';
  if (n==='消耗品費') return '8621';
  return parCode;
}

/* ===== 本体 ===== */
function convertImportToEpson_STRICT(){
  setupOrRepairSheets();
  var ss=SpreadsheetApp.getActive(), imp=ss.getSheetByName(SHEETS.IMPORT);
  if(!imp||imp.getLastRow()<IMPORT_HEADER_ROW+1){ SpreadsheetApp.getUi().alert('1_Data_import が空です'); return; }

  showProgressSidebar_();

  var lastRow=imp.getLastRow(), lastCol=Math.max(24,imp.getLastColumn());
  var vals=imp.getRange(1,1,lastRow,lastCol).getDisplayValues();
  var start=IMPORT_HEADER_ROW+1;

  var COL_DENNO=2, COL_DATE=IMPORT_COLS.date;

  // === Mapping_store 読み（A..I）: コードキーのみ ===
  var storeSh=ss.getSheetByName(SHEETS.MAP_STORE);
  var mapByCode=new Map(); // key: (A:JDL親コード '|' B:JDL補助コード) → rec(E,F,G,H)
  if(storeSh&&storeSh.getLastRow()>=2){
    var svals=storeSh.getRange(2,1,storeSh.getLastRow()-1,9).getDisplayValues();
    for(var i=0;i<svals.length;i++){
      var jParCode=(svals[i][0]||'').toString().trim(); // A
      var jSubCode=(svals[i][1]||'').toString().trim(); // B
      var eCod=(svals[i][4]||'').toString().trim();     // E
      var fSub=(svals[i][5]||'').toString().trim();     // F
      var gNam=(svals[i][6]||'').toString().trim();     // G
      var hNam=(svals[i][7]||'').toString().trim();     // H
      if(!jParCode) continue;
      var rec={eCode:eCod,fSub:fSub,gName:gNam,hName:hNam};
      mapByCode.set(jParCode+'|'+(jSubCode||''), rec);
      if(!jSubCode){ mapByCode.set(jParCode+'|', rec); }
    }
  }

  // EPSON台帳（名称補完用）
  var epCodeToName=_ep_codeToName_local_();
  var subsIx=_subs_index_codeToName_local_();

  // ログ準備
  var logs=ss.getSheetByName(SHEETS.LOGS), lbuf=[];
  function _now(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
  function logPush(level,proc,srcRow,side,jPar,jSub,reason,key){ lbuf.push([_now(),level,proc,srcRow,side,jPar,jSub,reason,key]); }
  function flushLogs(){ if(lbuf.length){ logs.getRange(logs.getLastRow()+1,1,lbuf.length,9).setValues(lbuf); lbuf=[]; } }

  function pickByCodeOrOverride_(parCode, subCode, sideLabel, rowNo, parNameForLog, subNameForLog){
    // ★ 買掛金/消耗品費は親名から親コードを強制補完
    parCode = _forceParentCodeIfKaidakaOrShomohin_(parCode, parNameForLog);

    var rec = null;
    if(parCode){
      rec = mapByCode.get(parCode+'|'+(subCode||'')) || mapByCode.get(parCode+'|');
    }
    if(!rec){
      var ov = _bankOverride_(parCode, subCode, parNameForLog, subNameForLog);
      if (ov) rec = ov;
    }
    if(rec){
      // ★ 名称補完（G/Hが空なら、E/Fコードから台帳で補完）
      rec = _completeNamesFromCodes_(rec, epCodeToName, subsIx);
      return { eCode:String(rec.eCode||''), fSub:String(rec.fSub||''), gName:rec.gName||'', hName:rec.hName||'' };
    }
    // 未ヒットはエラー
    if(parNameForLog || subNameForLog){
      if(sideLabel==='借方') progBump_(0,{errMapD:1}); else progBump_(0,{errMapC:1});
      var keyShown = (parCode?parCode:'(no-parent)')+'|'+(subCode||'');
      logPush('WARN','convert', rowNo, sideLabel, parNameForLog||'', subNameForLog||'', 'マッピング不足(A|B)', keyShown);
    }
    return {eCode:'',fSub:'',gName:'',hName:''};
  }

  // 出力ヘッダ
  var H=_buildConvertedHeader_(), HEADER=H.HEADER, APPEND_LEN=H.APPEND.length;
  var out=ss.getSheetByName(SHEETS.CONVERTED); out.clear();
  out.getRange(1,1,1,HEADER.length).setValues([HEADER]); out.setFrozenRows(1);

  // 税コード列は文字列
  var totalWork=Math.max(0,lastRow-IMPORT_HEADER_ROW);
  var idxDCode=HEADER.indexOf('借方消費税コード')+1, idxCCode=HEADER.indexOf('貸方消費税コード')+1;
  if(idxDCode>0) out.getRange(2,idxDCode,Math.max(1,totalWork),1).setNumberFormat('@STRING@');
  if(idxCCode>0) out.getRange(2,idxCCode,Math.max(1,totalWork),1).setNumberFormat('@STRING@');

  // 伝票番号（B列）空は日付ごとに通番
  var dateSeq=Object.create(null), outNo=new Array(lastRow+1).fill('');
  for(var r=start;r<=lastRow;r++){
    var d=_dateStr(vals[r-1][COL_DATE-1]), dn=(vals[r-1][COL_DENNO-1]||'').toString().trim();
    if(dn){ outNo[r]=dn; } else { var k=d||'(no-date)'; if(!(k in dateSeq)) dateSeq[k]=0; dateSeq[k]++; outNo[r]=dateSeq[k]; }
  }

  // 変換ループ
  var outRows=[], bumpEvery=50; progInit_(totalWork);
  for(var r=start; r<=lastRow; r++){
    var den=String(outNo[r]), jDate=vals[r-1][COL_DATE-1];

    // 借方
    var dParCode=(vals[r-1][IMPORT_COLS.debitCode-1]||'').toString().trim();
    var dParName=(vals[r-1][IMPORT_COLS.debitName-1]||'').toString().trim();
    var dSubCode=(vals[r-1][IMPORT_COLS.dSubCode-1]  ||'').toString().trim();
    var dSubName=(vals[r-1][IMPORT_COLS.dSubName-1]  ||'').toString().trim();
    var dAmt=Number((vals[r-1][IMPORT_COLS.debitAmt-1]||'0').toString().replace(/,/g,''));

    // 貸方
    var cParCode=(vals[r-1][IMPORT_COLS.creditCode-1]||'').toString().trim();
    var cParName=(vals[r-1][IMPORT_COLS.creditName-1]||'').toString().trim();
    var cSubCode=(vals[r-1][IMPORT_COLS.cSubCode-1]  ||'').toString().trim();
    var cSubName=(vals[r-1][IMPORT_COLS.cSubName-1]  ||'').toString().trim();
    var cAmt=Number((vals[r-1][IMPORT_COLS.creditAmt-1]||'0').toString().replace(/,/g,''));

    var dTaxRaw=(vals[r-1][IMPORT_COLS.dTax-1]||'').toString().trim();
    var cTaxRaw=(vals[r-1][IMPORT_COLS.cTax-1]||'').toString().trim();
    var memo=(vals[r-1][IMPORT_COLS.memo-1]||'').toString().trim();

    var dMap=pickByCodeOrOverride_(dParCode,dSubCode,'借方',r,dParName,dSubName);
    var cMap=pickByCodeOrOverride_(cParCode,cSubCode,'貸方',r,cParName,cSubName);

    var dTax=decideTax_(dParName,dTaxRaw), cTax=decideTax_(cParName,cTaxRaw);
    if(dTax.code===null){ progBump_(0,{errTaxD:1}); logPush('WARN','convert',r,'借方',dParName,dSubName,'税区分 未知: '+dTaxRaw,dParCode+'|'+dSubCode); dTax={code:'',rate:''}; }
    if(cTax.code===null){ progBump_(0,{errTaxC:1}); logPush('WARN','convert',r,'貸方',cParName,cSubName,'税区分 未知: '+cTaxRaw,cParCode+'|'+cSubCode); cTax={code:'',rate:''}; }

    var row=[
      _dateStr(jDate), den,
      dMap.eCode, dMap.gName, dMap.fSub, dMap.hName, (dAmt||''), (dTax.code!==''?String(dTax.code):''), (dTax.rate!==''?dTax.rate:''),
      cMap.eCode, cMap.gName, cMap.fSub, cMap.hName, (cAmt||''), (cTax.code!==''?String(cTax.code):''), (cTax.rate!==''?cTax.rate:''),
      memo
    ];
    if(APPEND_LEN>0) row=row.concat(new Array(APPEND_LEN).fill(''));
    outRows.push(row);

    var doneNow=r-(start-1);
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
  if(outRows.length) ss.getSheetByName(SHEETS.CONVERTED).getRange(2,1,outRows.length,HEADER.length).setValues(outRows);
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
  for (var i=0;i<data.length;i++) lines.push(data[i].map(esc).join(','));
  var csv=lines.join('\r\n');
  var blob=Utilities.newBlob('', 'text/csv', 'converted.csv'); blob.setDataFromString(csv, 'Shift_JIS');
  return {filename:'converted.csv', base64:Utilities.base64Encode(blob.getBytes()), mime:'text/csv; charset=shift_jis'};
}
/** ===== END convert_strict.gs ===== */

