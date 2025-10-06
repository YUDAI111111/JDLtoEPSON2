/*******************************************************
 * FinancialStatements.gs — 1_Data_import から JDL試算表を作成（左BS｜右PL）
 * 仕様：ヘッダ=4行目、データは5行目〜／出力は1枚シート（左BS｜右PL）
 * 科目分類は“名称ベースのみ”。特例：事業主貸=資産／事業主借=負債／家事消費=収益(マイナス表示)
 * 表示：空行なし／行高18px／折返しオフ／PLは「当期損益」に統一
 *******************************************************/
function buildJDLTrialBalance() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('1_Data_import');
  if (!src) throw new Error('シート「1_Data_import」が見つかりません。');

  var headerRow = 4; // 1-indexed
  var dataStart = headerRow + 1;
  var lastRow = src.getLastRow(), lastCol = src.getLastColumn();
  if (lastRow < dataStart) { toast_('1_Data_import にデータがありません'); return; }

  var values = src.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getDisplayValues();
  var header = values.shift(), rows = values;

  var col = colIndex_(header, [
    { key: 'DrCode',   names: ['借方科目','借方科目コード'] },
    { key: 'DrName',   names: ['借方科目名称'] },
    { key: 'DrSubCd',  names: ['借方補助','借方補助コード'] },
    { key: 'DrSubNm',  names: ['借方補助名称'] },
    { key: 'DrAmt',    names: ['借方金額'] },
    { key: 'CrCode',   names: ['貸方科目','貸方科目コード'] },
    { key: 'CrName',   names: ['貸方科目名称'] },
    { key: 'CrSubCd',  names: ['貸方補助','貸方補助コード'] },
    { key: 'CrSubNm',  names: ['貸方補助名称'] },
    { key: 'CrAmt',    names: ['貸方金額'] },
  ]);

  var MAP = buildNameClassMap_();

  var bsAgg = {}; // key=code|name|subCd|subNm → {cls,dr,cr,...}
  var plAgg = {};

  rows.forEach(function (r) {
    var dr = {name:safe(r[col.DrName]), code:safe(r[col.DrCode]), subCd:safe(r[col.DrSubCd]), subNm:safe(r[col.DrSubNm]), amt:num(r[col.DrAmt])};
    var cr = {name:safe(r[col.CrName]), code:safe(r[col.CrCode]), subCd:safe(r[col.CrSubCd]), subNm:safe(r[col.CrSubNm]), amt:num(r[col.CrAmt])};

    if (dr.name || dr.amt) addLine_(bsAgg, plAgg, classify_(dr.name, MAP), dr, dr.amt, 0);
    if (cr.name || cr.amt) addLine_(bsAgg, plAgg, classify_(cr.name, MAP), cr, 0, cr.amt);
  });

  var out = ss.getSheetByName('JDL試算表') || ss.insertSheet('JDL試算表');
  out.clear();

  var priorOpening = readOpening_(out); // 既存期首の引継ぎ

  // BS（左）
  var bsBlock = buildBs_(bsAgg, priorOpening);
  write_(out, 1, 1, bsBlock.values);
  formatBs_(out, 1, 1, bsBlock.values.length);

  // PL（右）
  var plBlock = buildPl_(plAgg);
  write_(out, 1, 10, plBlock.values);
  formatPl_(out, 1, 10, plBlock.values.length);

  toast_('JDL試算表を作成（左：BS／右：PL）');
}

/* ====== 分類マップ（名称ベースのみ／ご指定反映）====== */
function buildNameClassMap_(){
  var ASSET='ASSET', LIAB='LIAB', EQUITY='EQUITY', REV='REVENUE', EXP='EXPENSE';
  var map = {};
  function set(list, cls, opt){ list.forEach(function(n){ map[n]= {cls:cls, special:(opt&&opt.special)||null}; }); }

  set(['現金','小口現金','普通預金','積立預金','立替金','未収入金','仮払税','事業主貸'], ASSET);
  set(['買掛金','未払金','預り金','長期借入','事業主借'], LIAB);

  set(['売上高','雑収入','受取配当'], REV);
  set(['家事消費'], REV, {special:'KAJI'}); // 収益側でマイナス表示

  set(['仕入高','給与手当','賞与','法定福利','福利厚生','外注費','旅費交通','通信費','交際費','会議費',
       '賃借料','地代家賃','リース料','保険料','修繕費','水道光熱','燃料費','消耗品費','租税公課','事務用品',
       '広告宣伝','支払手数','諸会費','新聞図書','雑費','支払利息','雑損失'], EXP);
  return map;
}

/* ====== 集計・表構築 ====== */
function addLine_(bsAgg, plAgg, klass, acc, drAmt, crAmt){
  var key = key_(acc.code, acc.name, acc.subCd, acc.subNm);
  var cls = klass.cls;

  if (cls === 'ASSET' || cls === 'LIAB' || cls === 'EQUITY') {
    var o = bsAgg[key] || {code:acc.code,name:acc.name,subCd:acc.subCd,subNm:acc.subNm, cls:cls, dr:0, cr:0};
    o.dr += drAmt; o.cr += crAmt; bsAgg[key] = o;
  } else if (cls === 'REVENUE' || cls === 'EXPENSE'){
    var p = plAgg[key] || {code:acc.code,name:acc.name,subCd:acc.subCd,subNm:acc.subNm, cls:cls, dr:0, cr:0, special:klass.special||null};
    p.dr += drAmt; p.cr += crAmt; plAgg[key] = p;
  } else {
    var u = plAgg[key] || {code:acc.code,name:acc.name,subCd:acc.subCd,subNm:acc.subNm, cls:'UNASSIGNED', dr:0, cr:0, special:null};
    u.dr += drAmt; u.cr += crAmt; plAgg[key] = u;
  }
}

function buildBs_(agg, priorOpening){
  var rows = [];
  rows.push(['【貸借対照表（BS）】','','','','','','','']);
  rows.push(['科目コード','科目名','補助コード','補助名','期首','借方発生','貸方発生','期末']);

  var groups = {ASSET:[], LIAB:[], EQUITY:[]};
  Object.keys(agg).forEach(function(k){ var a=agg[k]; (groups[a.cls]=groups[a.cls]||[]).push(a); });

  ['ASSET','LIAB','EQUITY'].forEach(function(cls){
    var parents = groupParent_(groups[cls]||[]);
    parents.forEach(function(p){
      var openKey = key_(p.code,p.name,'','');
      var opening = priorOpening[openKey] || 0;
      var net = (cls==='ASSET') ? (p.dr - p.cr) : (p.cr - p.dr);
      rows.push([p.code, p.name, '', '', opening, p.dr, p.cr, opening + net]);

      p.subs.forEach(function(s){
        var sKey = key_(s.code,s.name,s.subCd,s.subNm);
        var sOpen = priorOpening[sKey] || 0;
        var sNet = (cls==='ASSET') ? (s.dr - s.cr) : (s.cr - s.dr);
        rows.push([s.code, '', s.subCd, '　→ ' + s.subNm, sOpen, s.dr, s.cr, sOpen + sNet]);
      });
      // 空行は入れない
    });
  });
  return {values: rows};
}

function buildPl_(agg){
  var rows = [];
  rows.push(['【損益計算書（PL）】','','','','','','']);
  rows.push(['科目コード','科目名','補助コード','補助名','借方発生','貸方発生','当期損益']); // 表記を統一

  var rev=[], exp=[], unk=[];
  Object.keys(agg).forEach(function(k){
    var o=agg[k];
    if (o.cls==='REVENUE') rev.push(o);
    else if (o.cls==='EXPENSE') exp.push(o);
    else unk.push(o);
  });

  [{title:'収益',list:rev,cls:'REVENUE'},{title:'費用',list:exp,cls:'EXPENSE'},{title:'未割当',list:unk,cls:'UNASSIGNED'}]
    .forEach(function(sec){
      var parents = groupParent_(sec.list);
      parents.forEach(function(p){
        var net = (sec.cls==='REVENUE') ? (p.cr - p.dr) : (p.dr - p.cr);
        if (p.special==='KAJI' || hasKajishi_(p)) net = -Math.abs(net); // 家事消費は収益マイナス表示
        rows.push([p.code, p.name, '', '', p.dr, p.cr, net]);
        p.subs.forEach(function(s){
          var sNet = (sec.cls==='REVENUE') ? (s.cr - s.dr) : (s.dr - s.cr);
          if (s.special==='KAJI') sNet = -Math.abs(sNet);
          rows.push([s.code, '', s.subCd, '　→ ' + s.subNm, s.dr, s.cr, sNet]);
        });
        // 空行は入れない
      });
    });

  return {values: rows};
}

/* ====== 並び・整形・保存 ====== */
function groupParent_(list){
  list = (list||[]).slice().sort(function(a,b){
    var A=(a.code||'')+(a.name||'')+(a.subCd||'')+(a.subNm||'');
    var B=(b.code||'')+(b.name||'')+(b.subCd||'')+(b.subNm||'');
    return A>B?1:(A<B?-1:0);
  });
  var map={};
  list.forEach(function(o){
    var pk=(o.code||'')+'|'+(o.name||'');
    if(!map[pk]) map[pk]={code:o.code,name:o.name,dr:0,cr:0,subs:[],special:o.special||null};
    map[pk].dr+=o.dr; map[pk].cr+=o.cr;
    if (o.subCd || o.subNm) map[pk].subs.push({code:o.code,name:o.name,subCd:o.subCd,subNm:o.subNm,dr:o.dr,cr:o.cr,special:o.special||null});
  });
  return Object.keys(map).map(function(k){return map[k];});
}

function readOpening_(sh){
  var map={}, vals=sh.getDataRange().getValues();
  for (var r=3; r<=vals.length; r++){
    var row=vals[r-1], code=safe(row[0]), name=safe(row[1]), subCd=safe(row[2]), subNm=safe(row[3]);
    if (!code && !name && !subCd && !subNm) continue;
    var opening = Number(row[4]) || 0;
    map[key_(code,name,subCd,subNm)] = opening;
  }
  return map;
}

function write_(sh, r, c, values){ sh.getRange(r, c, values.length, values[0].length).setValues(values); }

function formatBs_(sh, r, c, rows){
  sh.getRange(r, c, 1, 8).setFontWeight('bold');
  sh.getRange(r+1, c, 1, 8).setFontWeight('bold');
  [5,6,7,8].forEach(function(off){
    sh.getRange(r+2, c+off-1, rows-2, 1).setNumberFormat('#,##0;[Red]-#,##0;"-"');
  });
  // 折返しオフ＆行高詰め
  sh.getRange(r, c, rows, 8).setWrap(false);
  sh.setRowHeights(r+2, Math.max(0, rows-2), 18);
  // 親行を太字（補助名セルが空の行を親とみなす）
  var data=sh.getRange(r+2, c, rows-2, 4).getValues();
  for (var i=0;i<data.length;i++){
    var subNm=data[i][3];
    if (!subNm) sh.getRange(r+2+i, c, 1, 8).setFontWeight('bold');
  }
}

function formatPl_(sh, r, c, rows){
  sh.getRange(r, c, 1, 7).setFontWeight('bold');
  sh.getRange(r+1, c, 1, 7).setFontWeight('bold');
  [5,6,7].forEach(function(off){
    sh.getRange(r+2, c+off-1, rows-2, 1).setNumberFormat('#,##0;[Red]-#,##0;"-"');
  });
  // 折返しオフ＆行高詰め
  sh.getRange(r, c, rows, 7).setWrap(false);
  sh.setRowHeights(r+2, Math.max(0, rows-2), 18);
  // 親行太字
  var data=sh.getRange(r+2, c, rows-2, 4).getValues();
  for (var i=0;i<data.length;i++){
    var subNm=data[i][3];
    if (!subNm) sh.getRange(r+2+i, c, 1, 7).setFontWeight('bold');
  }
}

/* ====== 小物 ====== */
function toast_(m){ SpreadsheetApp.getActive().toast(m,'JDL試算表',5); }
function safe(v){ return (v==null)?'':String(v).trim(); }
function num(v){ if (v==null||v==='') return 0; var n=Number(String(v).replace(/,/g,'')); return isFinite(n)?n:0; }
function colIndex_(header, spec){
  var map={};
  spec.forEach(function(s){
    var idx=-1;
    for (var i=0;i<header.length;i++){ if(header[i]===s.names[0]){idx=i;break;} }
    if(idx<0){
      for (var i=0;i<header.length;i++){
        var h=header[i]; if(!h) continue;
        for (var j=0;j<s.names.length;j++){ if(String(h).indexOf(s.names[j])>=0){ idx=i; break; } }
        if(idx>=0) break;
      }
    }
    map[s.key]=(idx<0)?-1:idx;
  });
  ['DrCode','DrName','DrAmt','CrCode','CrName','CrAmt'].forEach(function(k){
    if(map[k]<0) throw new Error('必須列が見つかりません: '+k);
  });
  return map;
}
function classify_(name, MAP){
  if (MAP[name]) return MAP[name];
  var keys=Object.keys(MAP);
  for (var i=0;i<keys.length;i++){ var k=keys[i]; if(name.indexOf(k)>=0) return MAP[k]; }
  return {cls:'UNASSIGNED', special:null};
}
function hasKajishi_(p){
  if (p.name === '家事消費') return true;
  for (var i=0;i<(p.subs||[]).length;i++){ if (p.subs[i].subNm === '家事消費') return true; }
  return false;
}
// 追加：キー生成（コード|科目名|補助コード|補助名）
function key_(code, name, subCd, subNm) {
  return [code || '', name || '', subCd || '', subNm || ''].join('|');
}
