/*******************************************************
 * FinancialStatements.gs — CSV（＝既存JDL試算表）の並び順で再描画
 * 仕様（合意済）：
 *  - 並び順：既存「JDL試算表」の親順＋直下補助順を完全踏襲（BS/PLとも）。新規は末尾にのみ追加。
 *  - BS：E=期首, F=借方発生, G=貸方発生, H=期末残高（※Hは“式”で、期首を必ず含む）
 *      資産            ：H = E + (F - G)
 *      負債・純資産    ：H = E + (G - F)
 *      控除資産(償却累計)：H = E + (G - F)
 *  - PL：期首なし。当期のみ（収益=G−F、費用=F−G、家事消費は常にマイナス表示）
 *  - 行間詰め／折返しオフ／親太字／補助は「　→ 補助名」
 *  - 期首のみの科目もBSに必ず出す（openingと集計を合流）
 *******************************************************/
function buildJDLTrialBalance() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('1_Data_import');
  if (!src) throw new Error('シート「1_Data_import」が見つかりません。');

  // 1) 仕訳読込（ヘッダ=4行目）
  var headerRow = 4, dataStart = headerRow + 1;
  var lastRow = src.getLastRow(), lastCol = src.getLastColumn();
  if (lastRow < dataStart) { return toast_('1_Data_import にデータがありません'); }
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

  // 2) 既存JDL試算表から“並びひな型”と“期首”を読む（＝CSV順）
  var out = ss.getSheetByName('JDL試算表') || ss.insertSheet('JDL試算表');
  var skeleton = readSkeleton_(out);           // 並び順
  var opening  = readOpeningFromSheet_(out);   // 期首（E列）

  // 3) 分類マップ（固定資産・控除資産・家事消費対応）
  var MAP = buildNameClassMap_();

  // 4) 当期集計（借方=増、貸方=減の解釈は科目区分で後段に適用）
  var agg = aggregate_(rows, col, MAP);

  // 5) BSは“期首キー”を必ず含める（期中ゼロでも出す）
  Object.keys(opening).forEach(function(k){
    if (!agg.bs[k]) {
      var parts = k.split('|');
      var name  = parts[1] || '';
      var cls   = classify_(name, MAP).cls;
      var spec  = classify_(name, MAP).special || null;
      agg.bs[k] = {code:parts[0]||'', name:name, subCd:parts[2]||'', subNm:parts[3]||'', cls:cls==='UNASSIGNED'?'ASSET':cls, special:spec, dr:0, cr:0};
    }
  });

  // 6) クリアして再描画（並びはskeleton通り、新規は末尾）
  out.clear();

  // 左：BS
  var bsHeaders = ['科目コード','科目名','補助コード','補助名','期首','借方発生','貸方発生','期末残高'];
  var bsRows = buildStrictBS_(skeleton.bs, opening, agg.bs);
  write_(out, 1, 1, [ ['【貸借対照表（BS）】','','','','','','',''], bsHeaders ].concat(bsRows.values));
  formatBlock_(out, 1, 1, bsRows.values.length + 2, 8, [5,6,7,8], 8);
  // 期末残高（H列）へ式を設定
  setBsFormulas_(out, 3, 1, bsRows.meta); // データ開始行=3行目

  // 右：PL
  var plHeaders = ['科目コード','科目名','補助コード','補助名','借方発生','貸方発生','当期損益'];
  var plRows = buildStrictPL_(skeleton.pl, agg.pl);
  write_(out, 1, 10, [ ['【損益計算書（PL）】','','','','','',''], plHeaders ].concat(plRows));
  formatBlock_(out, 1, 10, plRows.length + 2, 7, [5,6,7], 7);

  toast_('JDL試算表をCSVの並び順で再描画しました');
}

/* ====== 既存JDL試算表の“並びひな型”＆期首 ====== */
function readSkeleton_(sh){
  var sk = { bs:[], pl:[] };
  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 3) return sk;

  // BS：A-D
  for (var r=2; r<vals.length; r++){
    var code=safe(vals[r][0]), name=safe(vals[r][1]), subCd=safe(vals[r][2]), subNm=safe(vals[r][3]);
    if (code||name||subCd||subNm) sk.bs.push({code:code,name:name,subCd:subCd,subNm:subNm});
  }
  // PL：J-M
  for (var r=2; r<vals.length; r++){
    var code2=safe(vals[r][9]), name2=safe(vals[r][10]), subCd2=safe(vals[r][11]), subNm2=safe(vals[r][12]);
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

/* ====== 集計 ====== */
function aggregate_(rows, col, MAP){
  var bs={}, pl={};
  function upsert(bucket, o, cls, special){
    var k = key_(o.code,o.name,o.subCd,o.subNm);
    var t = bucket[k] || {code:o.code,name:o.name,subCd:o.subCd,subNm:o.subNm, cls:cls, special:special||null, dr:0, cr:0};
    t.dr += o.dr; t.cr += o.cr; bucket[k]=t;
  }
  rows.forEach(function(r){
    var dr = {code:safe(r[col.DrCode]), name:safe(r[col.DrName]), subCd:safe(r[col.DrSubCd]), subNm:safe(r[col.DrSubNm]), dr:num(r[col.DrAmt])||0, cr:0};
    var cr = {code:safe(r[col.CrCode]), name:safe(r[col.CrName]), subCd:safe(r[col.CrSubCd]), subNm:safe(r[col.CrSubNm]), dr:0, cr:num(r[col.CrAmt])||0};
    if (dr.name || dr.dr) { var c=classify_(dr.name, MAP); (isBS_(c.cls))?upsert(bs,dr,c.cls,c.special):upsert(pl,dr,c.cls,c.special); }
    if (cr.name || cr.cr) { var c2=classify_(cr.name, MAP); (isBS_(c2.cls))?upsert(bs,cr,c2.cls,c2.special):upsert(pl,cr,c2.cls,c2.special); }
  });
  return {bs:bs, pl:pl};
}

/* ====== BS出力（並びはひな型通り。新規は末尾）＋ 期末数式メタ出力 ====== */
function buildStrictBS_(skeleton, opening, bsAgg){
  var rows=[], meta=[], seen={};

  // ひな型行をそのまま走査
  skeleton.forEach(function(s){
    var k = key_(s.code,s.name,s.subCd,s.subNm);
    var a = bsAgg[k] || {dr:0,cr:0,cls:classify_(s.name, buildNameClassMap_()).cls, special:classify_(s.name, buildNameClassMap_()).special};
    var open = (opening.hasOwnProperty(k)) ? opening[k] : 0;

    rows.push([
      s.code, s.name || '', s.subCd || '', s.subNm ? '　→ ' + s.subNm : '',
      open, a.dr, a.cr, '' // H列は後で式で入れる
    ]);
    meta.push({cls:a.cls, special:a.special}); // 同じ行順で保持
    seen[k]=true;
  });

  // ひな型に無い新規（末尾へ）
  Object.keys(bsAgg).forEach(function(k){
    if (seen[k]) return;
    var a=bsAgg[k], open = opening[k] || 0;
    rows.push([a.code, a.name || '', a.subCd || '', a.subNm ? '　→ ' + a.subNm : '', open, a.dr, a.cr, '']);
    meta.push({cls:a.cls, special:a.special||null});
  });

  return {values:rows, meta:meta};
}

/* ====== PL出力（並びはひな型通り。新規は末尾） ====== */
function buildStrictPL_(skeleton, plAgg){
  var rows=[], seen={};

  skeleton.forEach(function(s){
    var k = key_(s.code,s.name,s.subCd,s.subNm);
    var a = plAgg[k] || {dr:0,cr:0,cls:classify_(s.name, buildNameClassMap_()).cls, special:classify_(s.name, buildNameClassMap_()).special};
    var net = (a.cls==='REVENUE') ? (a.cr - a.dr) : (a.dr - a.cr);
    if (a.special==='KAJI' || s.name==='家事消費') net = -Math.abs(net);
    rows.push([ s.code, s.name || '', s.subCd || '', s.subNm ? '　→ ' + s.subNm : '', a.dr, a.cr, net ]);
    seen[k]=true;
  });

  Object.keys(plAgg).forEach(function(k){
    if (seen[k]) return;
    var a=plAgg[k];
    var net = (a.cls==='REVENUE') ? (a.cr - a.dr) : (a.dr - a.cr);
    if (a.special==='KAJI') net = -Math.abs(net);
    rows.push([ a.code, a.name || '', a.subCd || '', a.subNm ? '　→ ' + a.subNm : '', a.dr, a.cr, net ]);
  });

  return rows;
}

/* ====== H列に“式”を入れる（期首を必ず含む） ====== */
function setBsFormulas_(sh, dataStartRow, startCol, meta){
  // dataStartRow: データ開始の行番号（見出し＋ヘッダで2行使用 → 通常3）
  // startCol    : BSの開始列（通常 1 = 列A）
  var Hcol = startCol + 7; // H
  var Ecol = startCol + 4; // E
  var Fcol = startCol + 5; // F
  var Gcol = startCol + 6; // G
  for (var i=0;i<meta.length;i++){
    var r = dataStartRow + i;
    var m = meta[i] || {};
    var isAsset = (m.cls === 'ASSET');
    var isContra = (m.special === 'CONTRA_ASSET');
    var formula;
    if (isAsset && !isContra) {
      formula = '='+colLetter_(Ecol)+r+'+('+colLetter_(Fcol)+r+'-'+colLetter_(Gcol)+r+')';
    } else {
      // 負債・純資産・控除資産：期末 = 期首 + (貸方 − 借方)
      formula = '='+colLetter_(Ecol)+r+'+('+colLetter_(Gcol)+r+'-'+colLetter_(Fcol)+r+')';
    }
    sh.getRange(r, Hcol).setFormula(formula);
  }
}

/* ====== 体裁 ====== */
function formatBlock_(sh, r, c, rows, width, numCols, boldColCount){
  sh.getRange(r, c, 1, width).setFontWeight('bold');
  sh.getRange(r+1, c, 1, width).setFontWeight('bold');
  numCols.forEach(function(off){
    sh.getRange(r+2, c+off-1, rows-2, 1).setNumberFormat('#,##0;[Red]-#,##0;"-"');
  });
  sh.getRange(r, c, rows, width).setWrap(false);
  sh.setRowHeights(r+2, Math.max(0, rows-2), 18);

  var data=sh.getRange(r+2, c, rows-2, 4).getValues();
  for (var i=0;i<data.length;i++){
    var subNm=data[i][3];
    if (!subNm) sh.getRange(r+2+i, c, 1, boldColCount).setFontWeight('bold');
  }
}

/* ====== 分類：固定資産＋控除資産＋家事消費 ====== */
function buildNameClassMap_(){
  var map = {};
  function set(list, cls, opt){ list.forEach(function(n){ map[n]= {cls:cls, special:(opt&&opt.special)||null}; }); }
  function A(list,opt){ set(list,'ASSET',opt); }
  function L(list){ set(list,'LIAB'); }
  function E(list){ set(list,'EQUITY'); }
  function R(list,opt){ set(list,'REVENUE',opt); }
  function X(list){ set(list,'EXPENSE'); }

  // 固定資産（通常資産）
  A(['土地','建物','建物付属設備','建物附属設備','構築物','機械装置','機械設備','車両運搬具','工具器具備品','器具備品','ソフトウェア','電話加入権','のれん']);
  // 控除資産（減価償却累計額系）
  A(['減価償却累計額','建物減価償却累計額','建物付属設備減価償却累計額','建物附属設備減価償却累計額','機械装置減価償却累計額','機械設備減価償却累計額','車両運搬具減価償却累計額','工具器具備品減価償却累計額'], {special:'CONTRA_ASSET'});
  // 流動資産など
  A(['現金','小口現金','普通預金','積立預金','定期預金','立替金','未収入金','仮払税','事業主貸']);

  // 負債
  L(['買掛金','未払金','預り金','長期借入','事業主借']);
  // 収益
  R(['売上高','雑収入','受取配当']);
  R(['家事消費'], {special:'KAJI'});
  // 費用
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
function key_(code, name, subCd, subNm){ return [code||'',name||'',subCd||'',subNm||''].join('|'); }
function colLetter_(n){
  var s=""; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s;
}
