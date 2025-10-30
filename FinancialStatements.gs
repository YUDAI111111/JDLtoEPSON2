/*******************************************************
 * FinancialStatements.gs — ES5セーフ版（write_不使用・末尾カンマ無し）
 *******************************************************/
function buildJDLTrialBalance() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('1_Data_import');
  if (!src) throw new Error('シート「1_Data_import」が見つかりません。');

  var headerRow = 4, dataStart = headerRow + 1;
  var lastRow = src.getLastRow(), lastCol = src.getLastColumn();
  if (lastRow < dataStart) { return toast_('1_Data_import にデータがありません'); }
  var values = src.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getDisplayValues();
  var header = values.shift(), rows = values;
  var spec = [
    { key: 'DrCode',  names: ['借方科目','借方科目コード'] },
    { key: 'DrName',  names: ['借方科目名称'] },
    { key: 'DrSubCd', names: ['借方補助','借方補助コード'] },
    { key: 'DrSubNm', names: ['借方補助名称'] },
    { key: 'DrAmt',   names: ['借方金額'] },
    { key: 'CrCode',  names: ['貸方科目','貸方科目コード'] },
    { key: 'CrName',  names: ['貸方科目名称'] },
    { key: 'CrSubCd', names: ['貸方補助','貸方補助コード'] },
    { key: 'CrSubNm', names: ['貸方補助名称'] },
    { key: 'CrAmt',   names: ['貸方金額'] }
  ];
  var col = colIndex_(header, spec);

  var out = ss.getSheetByName('JDL試算表') || ss.insertSheet('JDL試算表');
  var skeleton = readSkeleton_(out);
  var opening  = readOpeningFromSheet_(out);

  var MAP = buildNameClassMap_();
  var agg = aggregate_(rows, col, MAP);

  Object.keys(opening).forEach(function(k){
    if (!agg.bs[k]) {
      var parts = k.split('|');
      var name  = parts[1] || '';
      var cx    = classify_(name, MAP);
      var cls   = cx.cls;
      var specx = cx.special || null;
      agg.bs[k] = {code:(parts[0]||''), name:name, subCd:(parts[2]||''), subNm:(parts[3]||''), cls:(cls==='UNASSIGNED'?'ASSET':cls), special:specx, dr:0, cr:0};
    }
  });

  out.clear();

  var bsTitle   = ['【貸借対照表（BS）】','','','','','','',''];
  var bsHeaders = ['科目コード','科目名','補助コード','補助名','期首','借方発生','貸方発生','期末残高'];
  var bsRowsObj = buildStrictBS_(skeleton.bs, opening, agg.bs);
  var bsMatrix  = [bsTitle, bsHeaders].concat(bsRowsObj.values);
  if (bsMatrix.length > 0) {
    out.getRange(1, 1, bsMatrix.length, 8).setValues(bsMatrix);
    formatBlock_(out, 1, 1, bsMatrix.length, 8, [5,6,7,8], 8);
    setBsFormulas_(out, 3, 1, bsRowsObj.meta);
  }

  var plTitle   = ['【損益計算書（PL）】','','','','','',''];
  var plHeaders = ['科目コード','科目名','補助コード','補助名','借方発生','貸方発生','当期損益'];
  var plRows    = buildStrictPL_(skeleton.pl, agg.pl);
  var plMatrix  = [plTitle, plHeaders].concat(plRows);
  if (plMatrix.length > 0) {
    out.getRange(1, 10, plMatrix.length, 7).setValues(plMatrix);
    formatBlock_(out, 1, 10, plMatrix.length, 7, [5,6,7], 7);
  }

  toast_('JDL試算表をCSVの並び順で再描画しました');
}

function readSkeleton_(sh){
  var sk = { bs:[], pl:[] };
  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 3) return sk;

  for (var r=2; r<vals.length; r++){
    var code=safe(vals[r][0]), name=safe(vals[r][1]), subCd=safe(vals[r][2]), subNm=safe(vals[r][3]);
    if (code||name||subCd||subNm) sk.bs.push({code:code,name:name,subCd:subCd,subNm:subNm});
  }
  for (var r2=2; r2<vals.length; r2++){
    var code2=safe(vals[r2][9]), name2=safe(vals[r2][10]), subCd2=safe(vals[r2][11]), subNm2=safe(vals[r2][12]);
    if (code2||name2||subCd2||subNm2) sk.pl.push({code:code2,name:name2,subCd:subCd2,subNm:subNm2});
  }
  return sk;
}
function readOpeningFromSheet_(sh){
  var map={}, vals=sh.getDataRange().getValues();
  if (!vals || vals.length < 3) return map;
  for (var r=2; r<vals.length; r++){
    var code=safe(vals[r][0]), name=safe(vals[r][1]), subCd=safe(vals[r][2]), subNm=safe(vals[r][3]);
    var k = key_(code,name,subCd,subNm);
    var open = vals[r][4];
    if (open!=='' && open!=null) map[k] = Number(open) || 0;
  }
  return map;
}

function aggregate_(rows, col, MAP){
  var bs={}, pl={};
  function upsert(bucket, o, cls, special){
    var k = key_(o.code,o.name,o.subCd,o.subNm);
    var t = bucket[k];
    if (!t) t = {code:o.code,name:o.name,subCd:o.subCd,subNm:o.subNm, cls:cls, special:(special||null), dr:0, cr:0};
    t.dr += o.dr; t.cr += o.cr; bucket[k]=t;
  }
  rows.forEach(function(r){
    var dr = {code:safe(r[col.DrCode]), name:safe(r[col.DrName]), subCd:safe(r[col.DrSubCd]), subNm:safe(r[col.DrSubNm]), dr:num(r[col.DrAmt])||0, cr:0};
    var cr = {code:safe(r[col.CrCode]), name:safe(r[col.CrName]), subCd:safe(r[col.CrSubCd]), subNm:safe(r[col.CrSubNm]), dr:0, cr:num(r[col.CrAmt])||0};
    if (dr.name || dr.dr) { var c=classify_(dr.name, MAP); if (isBS_(c.cls)) upsert(bs,dr,c.cls,c.special); else upsert(pl,dr,c.cls,c.special); }
    if (cr.name || cr.cr) { var c2=classify_(cr.name, MAP); if (isBS_(c2.cls)) upsert(bs,cr,c2.cls,c2.special); else upsert(pl,cr,c2.cls,c2.special); }
  });
  return {bs:bs, pl:pl};
}

function buildStrictBS_(skeleton, opening, bsAgg){
  var rows=[], meta=[], seen={};
  var MAP = buildNameClassMap_();

  for (var i=0;i<skeleton.length;i++){
    var s = skeleton[i];
    var k = key_(s.code,s.name,s.subCd,s.subNm);
    var cx = classify_(s.name, MAP);
    var a  = bsAgg[k] || {dr:0,cr:0,cls:cx.cls, special:cx.special};
    var open = opening.hasOwnProperty(k) ? opening[k] : 0;
    rows.push([ s.code, (s.name||''), (s.subCd||''), (s.subNm ? '　→ ' + s.subNm : ''), open, a.dr, a.cr, '' ]);
    meta.push({cls:a.cls, special:(a.special||null)});
    seen[k]=true;
  }

  var keys = Object.keys(bsAgg);
  for (var j=0;j<keys.length;j++){
    var kk = keys[j];
    if (seen[kk]) continue;
    var a2 = bsAgg[kk];
    var open2 = opening[kk] || 0;
    rows.push([ a2.code, (a2.name||''), (a2.subCd||''), (a2.subNm ? '　→ ' + a2.subNm : ''), open2, a2.dr, a2.cr, '' ]);
    meta.push({cls:a2.cls, special:(a2.special||null)});
  }

  return {values:rows, meta:meta};
}

function buildStrictPL_(skeleton, plAgg){
  var rows=[], seen={};
  var MAP = buildNameClassMap_();

  for (var i=0;i<skeleton.length;i++){
    var s = skeleton[i];
    var k = key_(s.code,s.name,s.subCd,s.subNm);
    var cx = classify_(s.name, MAP);
    var a  = plAgg[k] || {dr:0,cr:0,cls:cx.cls, special:cx.special};
    var net = (a.cls==='REVENUE') ? (a.cr - a.dr) : (a.dr - a.cr);
    if (a.special==='KAJI' || s.name==='家事消費') net = -Math.abs(net);
    rows.push([ s.code, (s.name||''), (s.subCd||''), (s.subNm ? '　→ ' + s.subNm : ''), a.dr, a.cr, net ]);
    seen[k]=true;
  }

  var keys = Object.keys(plAgg);
  for (var j=0;j<keys.length;j++){
    var kk = keys[j];
    if (seen[kk]) continue;
    var a2 = plAgg[kk];
    var net2 = (a2.cls==='REVENUE') ? (a2.cr - a2.dr) : (a2.dr - a2.cr);
    if (a2.special==='KAJI') net2 = -Math.abs(net2);
    rows.push([ a2.code, (a2.name||''), (a2.subCd||''), (a2.subNm ? '　→ ' + a2.subNm : ''), a2.dr, a2.cr, net2 ]);
  }
  return rows;
}

function setBsFormulas_(sh, dataStartRow, startCol, meta){
  var Hcol = startCol + 7;
  var Ecol = startCol + 4;
  var Fcol = startCol + 5;
  var Gcol = startCol + 6;
  for (var i=0;i<meta.length;i++){
    var r = dataStartRow + i;
    var m = meta[i] || {};
    var isAsset = (m.cls === 'ASSET');
    var isContra = (m.special === 'CONTRA_ASSET');
    var formula;
    if (isAsset && !isContra) {
      formula = '='+colLetter_(Ecol)+r+'+('+colLetter_(Fcol)+r+'-'+colLetter_(Gcol)+r+')';
    } else {
      formula = '='+colLetter_(Ecol)+r+'+('+colLetter_(Gcol)+r+'-'+colLetter_(Fcol)+r+')';
    }
    sh.getRange(r, Hcol).setFormula(formula);
  }
}

function formatBlock_(sh, r, c, rows, width, numCols, boldColCount){
  sh.getRange(r, c, 1, width).setFontWeight('bold');
  sh.getRange(r+1, c, 1, width).setFontWeight('bold');
  for (var i=0;i<numCols.length;i++){
    var off = numCols[i];
    sh.getRange(r+2, c+off-1, rows-2, 1).setNumberFormat('#,##0;[Red]-#,##0;"-"');
  }
  sh.getRange(r, c, rows, width).setWrap(false);
  sh.setRowHeights(r+2, Math.max(0, rows-2), 18);

  var data=sh.getRange(r+2, c, rows-2, 4).getValues();
  for (var j=0;j<data.length;j++){
    var subNm=data[j][3];
    if (!subNm) sh.getRange(r+2+j, c, 1, boldColCount).setFontWeight('bold');
  }
}

function buildNameClassMap_(){
  var map = {};
  function set(list, cls, opt){
    for (var i=0;i<list.length;i++){
      var n = list[i];
      map[n] = {cls:cls, special: (opt && opt.special) ? opt.special : null};
    }
  }
  function A(list,opt){ set(list,'ASSET',opt); }
  function L(list){ set(list,'LIAB'); }
  function E(list){ set(list,'EQUITY'); }
  function R(list,opt){ set(list,'REVENUE',opt); }
  function X(list){ set(list,'EXPENSE'); }

  A(['土地','建物','建物付属設備','建物附属設備','構築物','機械装置','機械設備','車両運搬具','工具器具備品','器具備品','ソフトウェア','電話加入権','のれん']);
  A(['減価償却累計額','建物減価償却累計額','建物付属設備減価償却累計額','建物附属設備減価償却累計額','機械装置減価償却累計額','機械設備減価償却累計額','車両運搬具減価償却累計額','工具器具備品減価償却累計額'], {special:'CONTRA_ASSET'});
  A(['現金','小口現金','普通預金','積立預金','定期預金','立替金','未収入金','仮払税','事業主貸']);

  L(['買掛金','未払金','預り金','長期借入','事業主借']);

  R(['売上高','雑収入','受取配当']);
  R(['家事消費'], {special:'KAJI'});

  X(['仕入高','給与手当','賞与','法定福利','福利厚生','外注費','旅費交通','通信費','交際費','会議費',
     '賃借料','地代家賃','リース料','保険料','修繕費','水道光熱','燃料費','消耗品費','租税公課','事務用品',
     '広告宣伝','支払手数','諸会費','新聞図書','雑費','支払利息','雑損失']);

  return map;
}
function classify_(name, MAP){
  if (MAP[name]) return MAP[name];
  var keys=Object.keys(MAP);
  for (var i=0;i<keys.length;i++){ var k=keys[i]; if(name.indexOf(k)>=0) return MAP[k]; }
  return {cls:'UNASSIGNED', special:null};
}
function isBS_(cls){ return (cls==='ASSET'||cls==='LIAB'||cls==='EQUITY'); }

function toast_(m){ SpreadsheetApp.getActive().toast(m,'JDL試算表',5); }
function safe(v){ return (v==null)?'':String(v).trim(); }
function num(v){ if (v==null||v==='') return 0; var n=Number(String(v).replace(/,/g,'')); return isFinite(n)?n:0; }
function colIndex_(header, spec){
  var map={}, i, j;
  for (i=0;i<spec.length;i++){
    var s = spec[i];
    var idx=-1;
    for (j=0;j<header.length;j++){ if(header[j]===s.names[0]){idx=j;break;} }
    if(idx<0){
      for (j=0;j<header.length;j++){
        var h=header[j]; if(!h) continue;
        var k;
        for (k=0;k<s.names.length;k++){ if(String(h).indexOf(s.names[k])>=0){ idx=j; break; } }
        if(idx>=0) break;
      }
    }
    map[s.key]=(idx<0)?-1:idx;
  }
  var req = ['DrCode','DrName','DrAmt','CrCode','CrName','CrAmt'];
  for (i=0;i<req.length;i++){ var r=req[i]; if(map[r]<0) throw new Error('必須列が見つかりません: '+r); }
  return map;
}
function key_(code, name, subCd, subNm){ return [code||'',name||'',subCd||'',subNm||''].join('|'); }
function colLetter_(n){
  var s=""; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s;
}
