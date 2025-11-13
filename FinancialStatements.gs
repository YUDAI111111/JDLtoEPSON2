/**
 * FinancialStatements_STRICT_Full_v3.gs
 * v3 変更点（ユーザー要件反映）
 *  - H3=現金、H4=小口現金、H5=普通預金(親=子HのSUM)、H6〜=普通預金の補助（行内式）
 *  - 小計式は TB_Order ラベルかつ TB_Attributes.SubtotalOnly=TRUE のみ
 *  - 1311 特例: 「普通預金(1311) 補助なし」は 1312 グループに吸収し、補助(10, 長野信用金庫)へ自動付与
 *    ・グルーピングは“親キー=#1312”に統一（表示の科目コードは元のまま）
 *    ・TB_Order側で 1311 の補助行があっても、親の判定は #1312 グループとして扱う（二重防止）
 *  - 親=補助SUM（SUBS_ONLY）。補助なし親は行内式（BS: 資産=E+F−G／その他=E+G−F、PL: E−F）
 *  - セル結合なし／符号反転なし／行列の勝手な追加なし
 */

var FS_ES5_HEADER_ROW = 4;                 // 1_Data_import のヘッダ行
var FS_ES5_ALLOW_SUB_DEFAULT = true;       // 補助許可のデフォルト

/** ===== エントリ ===== */
function buildJDLTrialBalance(){
  var ctx = { startedAt:new Date(), strict:true };
  try{
    assertPreconditions_STRICT_();
    var ss = SpreadsheetApp.getActive();
    var shSrc   = mustSheet_(ss, '1_Data_import');
    var shOrder = mustSheet_(ss, 'TB_Order');
    var shAttr  = mustSheet_(ss, 'TB_Attributes');
    var shTB    = ss.getSheetByName('JDL試算表') || ss.insertSheet('JDL試算表');

    // 読み込み
    var ATTR  = readTBAttributes_(shAttr);
    var ORDER = readTBOrder_(shOrder, ATTR);
    var SIDE  = buildSideMap_(ORDER);
    var OPEN  = readOpening_codeFirst_(shTB, ATTR); // 期首（既存TBに書いてある場合のみ使用）
    var jr    = readJournal_(shSrc);
    guardHtmlLikeContent_STRICT_(jr);
    var AGG   = aggregateSubs_codeFirst_WITH_DEFAULTS_(jr.rows, jr.col, ATTR, SIDE);

    // クリアして描画
    shTB.clear();

    // BS
    var bs = renderSide_('BS', shTB, 1, 1, ORDER.bs, ATTR, OPEN, AGG);
    applySubtotalFormulas_('BS', shTB, bs, 8, ATTR); // 期末=H列

    // PL
    var pl = renderSide_('PL', shTB, 1, 10, ORDER.pl, ATTR, OPEN, AGG);
    applySubtotalFormulas_('PL', shTB, pl, 7, ATTR); // 当期損益=G列

    ctx.elapsedMs = (new Date())-ctx.startedAt;
    logJson_('SUCCESS', ctx, { bsRows:bs.totalRows, plRows:pl.totalRows });
    toast_('JDL試算表を再描画（v3: 1311特例＆小計式ガード）');
  }catch(e){
    logJson_('FAIL',{strict:true},{error:String(e)});
    try{ SpreadsheetApp.getUi().alert('JDL試算表の作成に失敗：'+String(e)); }catch(_){}
    throw e;
  }
}

/** ===== 側ごとの描画 ===== */
function renderSide_(SIDE_NAME, sh, topRow, leftCol, ORDER_SIDE, ATTR, OPEN, AGG){
  var title = (SIDE_NAME==='BS')?'【貸借対照表（BS）】':'【損益計算書（PL）】';
  var headers = (SIDE_NAME==='BS')
    ? ['科目コード','科目名','補助コード','補助名','期首','借方発生','貸方発生','期末残高']
    : ['科目コード','科目名','補助コード','補助名','借方発生','貸方発生','当期損益'];

  // 見出し
  write_(sh, topRow, leftCol, [[title]]);
  sh.getRange(topRow, leftCol, 1, 1).setFontWeight('bold');
  write_(sh, topRow+1, leftCol, [headers]);
  sh.getRange(topRow+1, leftCol, 1, headers.length).setFontWeight('bold');

  // 本体
  var outRows=[], meta=[], currentLabel=null, blocks=[];
  for (var i=0;i<ORDER_SIDE.length;i++){
    var o = ORDER_SIDE[i];
    var attr = ATTR.byDisp[o.dispKey] || {};
    var isExplicitSubtotalLabel = (!o.code && !o.subCd && !o.subNm && o.name && attr.subtotalOnly===true);
    var isPlainLabel            = (!o.code && !o.subCd && !o.subNm && o.name && attr.subtotalOnly!==true);
    var isLabel = isExplicitSubtotalLabel || isPlainLabel;

    if (isLabel){
      if (currentLabel){ currentLabel.endIndex = outRows.length - 1; }
      currentLabel = { labelName:o.name, dispKey:o.dispKey, isSubtotal:(attr.subtotalOnly===true),
                       startIndex: outRows.length + 1, labelIndex: outRows.length };
      outRows.push( SIDE_NAME==='BS' ? ['', o.name, '', '', '', '', '', '']
                                     : ['', o.name, '', '', '', '', ''] );
      meta.push({kind:'label', attr:attr, isSubtotalLabel:(attr.subtotalOnly===true), labelName:o.name});
      blocks.push(currentLabel);
      continue;
    }

    if (o.isSub){
      // 補助行
      var info = (SIDE_NAME==='BS') ? (AGG.bsSub[o.subKeyEff]||{dr:0,cr:0}) : (AGG.plSub[o.subKeyEff]||{dr:0,cr:0});
      var op   = (SIDE_NAME==='BS') ? (OPEN.bsSub[o.subKeyEff]||0) : 0;
      if (SIDE_NAME==='BS'){
        outRows.push([o.code||'', o.name, o.subCd, o.subNm||'', op, info.dr||0, info.cr||0, '']);
      }else{
        outRows.push([o.code||'', o.name, o.subCd, o.subNm||'', info.dr||0, info.cr||0, '']);
      }
      var isAsset = (attr.section1||'').trim()==='資産の部';
      meta.push({kind:'sub', attr:attr, isAsset:isAsset, parentKey:o.groupKey});
      continue;
    }

    // 親行
    var isAssetP = (attr.section1||'').trim()==='資産の部';
    var allowSub = (attr.hasOwnProperty('allowSub')) ? !!attr.allowSub : FS_ES5_ALLOW_SUB_DEFAULT;

    var parentOpen = (SIDE_NAME==='BS') ? (OPEN.bs[o.groupKey]||0) : 0;
    var base       = (SIDE_NAME==='BS') ? (AGG.bs[o.groupKey]||{dr:0,cr:0}) : (AGG.pl[o.groupKey]||{dr:0,cr:0});
    var parentDr   = base.dr||0, parentCr=base.cr||0;

    if (SIDE_NAME==='BS'){
      outRows.push([o.code||'', o.name, '', '', parentOpen, parentDr, parentCr, '']);
    }else{
      outRows.push([o.code||'', o.name, '', '', parentDr, parentCr, '']);
    }
    meta.push({kind:'parent', attr:attr, isAsset:isAssetP, key:o.groupKey, subKeysHint:allowSub?o.childKeysEff:[]});
  }
  if (currentLabel){ currentLabel.endIndex = outRows.length - 1; }

  // 出力 & 罫/数値
  var dataStart = topRow+2; // = 3行目開始（H3/G3 から明細）
  write_(sh, dataStart, leftCol, outRows);
  if (SIDE_NAME==='BS'){
    setNumberFormatCols_(sh, dataStart, leftCol, outRows.length, [5,6,7,8]);
  }else{
    setNumberFormatCols_(sh, dataStart, leftCol, outRows.length, [5,6,7]);
  }
  if(outRows.length>0) sh.setRowHeights(dataStart, outRows.length, 18);

  // 行内式/親=子SUM
  for (var i=0;i<meta.length;i++){
    var m = meta[i]; var r = dataStart + i;
    if (m.kind==='label') continue;
    if (SIDE_NAME==='BS'){
      var E=leftCol+4, F=leftCol+5, G=leftCol+6, H=leftCol+7;
      if (m.kind==='parent'){
        // 子（同グループ）の H を SUM。ヒントが無ければ行内式。
        var childRows = getChildRowIndexesEff_(meta, i, m.key).map(function(idx){ return dataStart + idx; });
        if (childRows.length>0){
          var terms = childRows.map(function(rr){ return colLetter_(H)+rr; }).join(',');
          sh.getRange(r, H).setFormula('=SUM(' + terms + ')');
        }else{
          var f = m.isAsset ? ('='+colLetter_(E)+r+'+('+colLetter_(F)+r+'-'+colLetter_(G)+r+')')
                            : ('='+colLetter_(E)+r+'+('+colLetter_(G)+r+'-'+colLetter_(F)+r+')');
          sh.getRange(r, H).setFormula(f);
        }
      }else{
        var f2 = m.isAsset ? ('='+colLetter_(E)+r+'+('+colLetter_(F)+r+'-'+colLetter_(G)+r+')')
                           : ('='+colLetter_(E)+r+'+('+colLetter_(G)+r+'-'+colLetter_(F)+r+')');
        sh.getRange(r, H).setFormula(f2);
      }
    }else{
      var E2=leftCol+4, F2=leftCol+5, G2=leftCol+6;
      if (m.kind==='parent'){
        var childRows2 = getChildRowIndexesEff_(meta, i, m.key).map(function(idx){ return dataStart + idx; });
        if (childRows2.length>0){
          var terms2 = childRows2.map(function(rr){ return colLetter_(G2)+rr; }).join(',');
          sh.getRange(r, G2).setFormula('=SUM(' + terms2 + ')');
        }else{
          sh.getRange(r, G2).setFormula('='+colLetter_(E2)+r+'-'+colLetter_(F2)+r);
        }
      }else{
        sh.getRange(r, G2).setFormula('='+colLetter_(E2)+r+'-'+colLetter_(F2)+r);
      }
    }
  }

  // ブロック情報（ラベル行の範囲）
  var blocks = buildBlocksFromMeta_(meta, dataStart);
  return { blocks: blocks, rowMeta: meta, totalRows: outRows.length, topRow: topRow, leftCol: leftCol };
}

/** ===== 小計ラベルへSUM式（親のみ対象） ===== */
function applySubtotalFormulas_(SIDE_NAME, sh, sideInfo, valueCol, ATTR){
  var A = sideInfo.leftCol + 0;
  var B = sideInfo.leftCol + 1;
  var C = sideInfo.leftCol + 2;
  var D = sideInfo.leftCol + 3;
  var V = sideInfo.leftCol + (valueCol-1);
  for (var b=0;b<sideInfo.blocks.length;b++){
    var blk = sideInfo.blocks[b];
    if (!blk.isSubtotalLabel) continue;
    // ラベルセルの表示名検証（空は式を入れない／ATTRでSubtotalOnly=TRUE確認）
    var labelName = sh.getRange(blk.labelRow, B).getDisplayValue();
    if (!labelName) continue;
    var dispKey = makeDispKey_(labelName);
    var a = ATTR.byDisp[dispKey];
    if (!a || a.subtotalOnly!==true) continue;

    if (blk.startRow>blk.endRow) {
      sh.getRange(blk.labelRow, V).setValue(0);
      continue;
    }
    // 親行のみ（コードあり & 補助コード/補助名は空）を抽出してSUM
    var f = '=IFERROR(SUM(FILTER(' + colLetter_(V)+blk.startRow+':'+colLetter_(V)+blk.endRow+','
                           + colLetter_(C)+blk.startRow+':'+colLetter_(C)+blk.endRow+'="",'
                           + colLetter_(D)+blk.startRow+':'+colLetter_(D)+blk.endRow+'="",'
                           + colLetter_(A)+blk.startRow+':'+colLetter_(A)+blk.endRow+'<>"")),0)';
    sh.getRange(blk.labelRow, V).setFormula(f);
  }
}

/** ===== メタ→ブロック境界 ===== */
function buildBlocksFromMeta_(meta, startR){
  var blocks=[]; var cur=null;
  for (var i=0;i<meta.length;i++){
    var m=meta[i];
    if (m.kind==='label'){
      if (cur){ cur.endRow = startR+i-1; blocks.push(cur); }
      cur = {labelRow:startR+i, isSubtotalLabel:(m.isSubtotalLabel===true), startRow:startR+i+1, endRow:startR+i};
    }
  }
  if (cur){ cur.endRow = startR+meta.length-1; blocks.push(cur); }
  return blocks;
}

/** ===== 親の配下補助（1311/1312同グループ対応） ===== */
function getChildRowIndexesEff_(meta, parentIdx, parentGroupKey){
  var res=[];
  for (var i=0;i<meta.length;i++){
    if (meta[i].kind==='sub' && meta[i].parentKey===parentGroupKey){ res.push(i); }
  }
  return res;
}

/** ===== TB_Attributes ===== */
function readTBAttributes_(sh){
  var vals = sh.getDataRange().getDisplayValues();
  if (vals.length < 2){ return {byDisp:{}, aliasMap:{}, defaultByCode:{}}; }
  var h = vals[0];
  var idxReq = headerIndex_(h,{ code:'科目コード', name:'科目名' });
  var idxOpt = headerIndexOpt_(h,{
    allowSub:'補助許可', aliasName:'表示名上書き',
    defSubCd:'DefaultSubCodeIfEmpty', defSubNm:'DefaultSubNameIfEmpty',
    side:'Side', sec1:'Section_L1', pMode:'ParentMode', subOnly:'SubtotalOnly'
  });
  var byDisp={}, aliasMap={}, defaultByCode={};
  for (var r=1;r<vals.length;r++){
    var v=vals[r]; var code=safe(v[idxReq.code]); var name=safe(v[idxReq.name]); if(!code && !name) continue;
    var alias=(idxOpt.aliasName>=0)?safe(v[idxOpt.aliasName]):''; alias=alias||name;
    var dispKey=makeDispKey_(alias);
    var allowSub=(idxOpt.allowSub>=0)?toBool_(v[idxOpt.allowSub]):FS_ES5_ALLOW_SUB_DEFAULT;
    var defCd=(idxOpt.defSubCd>=0)?safe(v[idxOpt.defSubCd]):''; var defNm=(idxOpt.defSubNm>=0)?safe(v[idxOpt.defSubNm]):'';
    var side=(idxOpt.side>=0)?String(safe(v[idxOpt.side])):''; var sec1=(idxOpt.sec1>=0)?safe(v[idxOpt.sec1]):'';
    var pMode=(idxOpt.pMode>=0)?String(safe(v[idxOpt.pMode])).toUpperCase():''; var subOnly=(idxOpt.subOnly>=0)?toBool_(v[idxOpt.subOnly]):false;
    byDisp[dispKey]={code:code,name:name,alias:alias,dispKey:dispKey,allowSub:allowSub,defSubCd:defCd,defSubNm:defNm,side:side,section1:sec1,parentMode:pMode,subtotalOnly:subOnly};
    aliasMap[ makeDispKey_(name) ] = alias;
    if (code && defCd && defNm){ defaultByCode[String(code).trim()] = {subCd:defCd, subNm:defNm}; }
  }
  // 1311 の DefaultSub が無い場合のデフォルトを強制（特例：10/長野信用金庫）
  if (!defaultByCode['1311']) defaultByCode['1311'] = {subCd:'10', subNm:'長野信用金庫'};
  return {byDisp:byDisp, aliasMap:aliasMap, defaultByCode:defaultByCode};
}

/** ===== TB_Order ===== */
function readTBOrder_(sh, ATTR){
  var vals = sh.getDataRange().getDisplayValues();
  if (vals.length<2) return {bs:[], pl:[]};
  var idx = headerIndex_(vals[0],{side:'側',code:'科目コード',name:'科目名',subCd:'補助コード',subNm:'補助名'});
  function effGroupKey(code, alias){
    if (String(code).trim()==='1311' && alias==='普通預金') return '#1312';
    if (String(code).trim()==='1312' && alias==='普通預金') return '#1312';
    return (code?('#'+code):makeDispKey_(alias));
  }
  function rowToObj(v){
    var code=safe(v[idx.code]), name=safe(v[idx.name]);
    var alias=ATTR.aliasMap[ makeDispKey_(name) ] || name;
    var dispKey=makeDispKey_(alias);
    var subCd=safe(v[idx.subCd]), subNm=safe(v[idx.subNm]);
    var groupKey = effGroupKey(code, alias);                  // 1311/1312 普通預金は #1312 に統一
    var subKeyEff=(subCd||subNm) ? (groupKey+'||'+subCd+'|'+subNm) : null;
    return {code:code||'', name:alias, dispKey:dispKey, key:'#'+(code||''), groupKey:groupKey,
            subCd:subCd, subNm:subNm, subKeyEff:subKeyEff, isSub:!!subKeyEff, childKeysEff:subKeyEff?[subKeyEff]:[]};
  }
  var bs=[],pl=[];
  for (var r=1;r<vals.length;r++){
    var v=vals[r]; if(!v||v.length===0) continue;
    var side=String(v[idx.side]||'').toUpperCase(); var o=rowToObj(v);
    if(!o.code && !o.name && !o.subCd && !o.subNm) continue;
    if(side==='PL') pl.push(o); else bs.push(o);
  }
  // 親ごとに子キーを集約（同グループ）
  function collectChildKeys(arr){
    var map={};
    arr.forEach(function(o){
      if (o.isSub){
        var parentKey = o.groupKey;
        (map[parentKey]=map[parentKey]||[]).push(o.subKeyEff);
      }
    });
    arr.forEach(function(o){
      if (!o.isSub && map[o.groupKey]) o.childKeysEff = map[o.groupKey];
    });
  }
  collectChildKeys(bs); collectChildKeys(pl);
  return {bs:bs, pl:pl};
}

/** ===== 側マップ ===== */
function buildSideMap_(ORDER){ var side={}; ORDER.bs.forEach(function(o){ side[o.groupKey]='BS'; }); ORDER.pl.forEach(function(o){ side[o.groupKey]='PL'; }); return side; }

/** ===== 期首（既存TBから） ===== */
function readOpening_codeFirst_(sh, ATTR){
  var vals=sh.getDataRange().getDisplayValues(); var bs={}, bsSub={};
  if(!vals || vals.length<3) return {bs:bs, bsSub:bsSub};
  for (var r=2;r<vals.length;r++){
    var code=safe(vals[r][0]), name=safe(vals[r][1]);
    var alias=ATTR.aliasMap[ makeDispKey_(name) ] || name;
    var dispKey=makeDispKey_(alias);
    var rawKey = code ? ('#'+code) : dispKey;
    var groupKey = (code==='1311' && alias==='普通預金') ? '#1312' : rawKey;
    var subCd=safe(vals[r][2]), subNm=safe(vals[r][3]); var open=num(vals[r][4]);
    if(subCd||subNm){ var sk=groupKey+'||'+subCd+'|'+subNm; bsSub[sk]=(bsSub[sk]||0)+open; }
    else{ bs[groupKey]=(bs[groupKey]||0)+open; }
  }
  return {bs:bs, bsSub:bsSub};
}

/** ===== 仕訳読み込み ===== */
function readJournal_(sh){
  var HEADER_ROW=FS_ES5_HEADER_ROW, DATA_START=HEADER_ROW+1;
  var lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow<DATA_START) return {rows:[], col:{}};
  var vals=sh.getRange(HEADER_ROW,1,lastRow-HEADER_ROW+1,lastCol).getDisplayValues();
  var header=vals.shift(), rows=vals;
  var col=colIndex_(header,[
    { key:'DrCode',  names:['借方科目','借方科目コード'] },
    { key:'DrName',  names:['借方科目名称','借方科目正式名称'] },
    { key:'DrSubCd', names:['借方補助','借方補助コード'] },
    { key:'DrSubNm', names:['借方補助名称'] },
    { key:'DrAmt',   names:['借方金額'] },
    { key:'CrCode',  names:['貸方科目','貸方科目コード'] },
    { key:'CrName',  names:['貸方科目名称','貸方科目正式名称'] },
    { key:'CrSubCd', names:['貸方補助','貸方補助コード'] },
    { key:'CrSubNm', names:['貸方補助名称'] },
    { key:'CrAmt',   names:['貸方金額'] },
  ]);
  return {rows:rows, col:col};
}

/** ===== 集計（DefaultSub と 1311 特例を適用） ===== */
function aggregateSubs_codeFirst_WITH_DEFAULTS_(rows, col, ATTR, SIDE){
  var bs={}, pl={}, bsSub={}, plSub={};
  function upsert(bucket, key, info, dr, cr){
    var t=bucket[key]||{code:info.code,name:info.alias,subCd:info.subCd,subNm:info.subNm,dr:0,cr:0};
    t.dr += dr||0; t.cr += cr||0; bucket[key]=t;
  }
  rows.forEach(function(r){
    var sides=[
      {code:safe(r[col.DrCode]), name:safe(r[col.DrName]), subCd:safe(r[col.DrSubCd]), subNm:safe(r[col.DrSubNm]), dr:num(r[col.DrAmt])||0, cr:0},
      {code:safe(r[col.CrCode]), name:safe(r[col.CrName]), subCd:safe(r[col.CrSubCd]), subNm:safe(r[col.CrSubNm]), dr:0, cr:num(r[col.CrAmt])||0},
    ];
    sides.forEach(function(x){
      if(!(x.name || x.code || x.dr || x.cr)) return;
      var alias=ATTR.aliasMap[ makeDispKey_(x.name) ] || x.name;
      var dispKey=makeDispKey_(alias);

      // 1311/1312 普通預金は #1312 グループに統一
      var groupKey = (x.code==='1311' && alias==='普通預金') ? '#1312'
                    : (x.code ? ('#'+x.code) : dispKey);

      // 1311 特例：補助空なら 10/長野信用金庫 を自動付与
      if (x.code==='1311' && alias==='普通預金' && !x.subCd && !x.subNm){
        var def1311 = ATTR.defaultByCode['1311'] || {subCd:'10', subNm:'長野信用金庫'};
        x.subCd = x.subCd || def1311.subCd;
        x.subNm = x.subNm || def1311.subNm;
      }

      // 既定補助（通常の DefaultSub）
      if (!x.subCd && !x.subNm && x.code){
        var defHit = ATTR.defaultByCode[String(x.code).trim()];
        if (defHit && defHit.subCd && defHit.subNm){
          x.subCd = defHit.subCd; x.subNm = defHit.subNm;
        }
      }

      var subKeyEff=(x.subCd||x.subNm)?(groupKey+'||'+x.subCd+'|'+x.subNm):null;
      var info={code:x.code, alias:alias, subCd:x.subCd, subNm:x.subNm};
      var side = SIDE[groupKey] || SIDE[dispKey] || 'PL';
      if(side==='BS'){ upsert(bs, groupKey, info, x.dr, x.cr); if(subKeyEff) upsert(bsSub, subKeyEff, info, x.dr, x.cr); }
      else{ upsert(pl, groupKey, info, x.dr, x.cr); if(subKeyEff) upsert(plSub, subKeyEff, info, x.dr, x.cr); }
    });
  });
  return {bs:bs, pl:pl, bsSub:bsSub, plSub:plSub};
}

/** ===== ユーティリティ ===== */
function mustSheet_(ss,name){ var sh=ss.getSheetByName(name); if(!sh) throw new Error('シート「'+name+'」が見つかりません。'); return sh; }
function write_(sh,r,c,values){ if(values.length) sh.getRange(r,c,values.length,values[0].length).setValues(values); }
function setNumberFormatCols_(sh, r, c, rows, colOffsets){
  colOffsets.forEach(function(off){ if(rows>0) sh.getRange(r,c+off-1,rows,1).setNumberFormat('#,##0;[Red]-#,##0;"-"'); });
}
function colLetter_(n){ var s=''; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=(n-m-1)/26|0; } return s; }
function safe(v){ return (v==null)?'':String(v).trim(); }
function num(v){ if (v==null||v==='') return 0; var n=Number(String(v).replace(/,/g,'')); return isFinite(n)?n:0; }
function toBool_(v){ var s=String(v||'').trim().toLowerCase(); return (s==='true'||s==='t'||s==='1'||s==='yes'||s==='y'||s==='on'); }
function headerIndex_(header, def){ var res={}; Object.keys(def).forEach(function(k){ var name=def[k]; var i=header.indexOf(name); if(i<0) throw new Error('ヘッダが見つかりません → '+name); res[k]=i; }); return res; }
function headerIndexOpt_(header, def){ var res={}; Object.keys(def).forEach(function(k){ var name=def[k]; res[k]=header.indexOf(name); }); return res; }
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
    if(map[k]<0) throw new Error('1_Data_import: 必須列が見つかりません → '+k);
  });
  return map;
}
function assertPreconditions_STRICT_(){
  ['1_Data_import','TB_Order','TB_Attributes','Logs'].forEach(function(n){ mustSheet_(SpreadsheetApp.getActive(), n); });
  var s = SpreadsheetApp.getActive().getSheetByName('1_Data_import');
  var lastRow = s.getLastRow();
  if (lastRow < FS_ES5_HEADER_ROW) throw new Error('1_Data_import: 行数不足（ヘッダがありません）');
  var header = s.getRange(FS_ES5_HEADER_ROW,1,1, s.getLastColumn()).getValues()[0].map(function(v){ return String(v).trim(); });
  var mustAny = [
    ['借方科目','借方科目コード'],
    ['借方金額'],
    ['貸方科目','貸方科目コード'],
    ['貸方金額']
  ];
  mustAny.forEach(function(group){
    var ok = group.some(function(name){ return header.indexOf(name)>=0; });
    if (!ok) throw new Error('1_Data_import ヘッダ欠落：' + group.join('／') + '（どれか一つは必須）');
  });
}
function guardHtmlLikeContent_STRICT_(jr){
  var sample = jr.rows.slice(0, Math.min(jr.rows.length, 10));
  var txt = sample.map(function(r){ return r.join(' '); }).join('\n');
  var lt=(txt.match(/</g)||[]).length, gt=(txt.match(/>/g)||[]).length;
  if (lt>=5 && gt>=5) throw new Error('入力にHTML様の内容を検出。アップロードしたCSV/TSVの中身を確認してください。');
}
function logJson_(status, ctx, extra){
  try{
    var s = mustSheet_(SpreadsheetApp.getActive(), 'Logs');
    var now = new Date();
    s.appendRow([ now, status, JSON.stringify({ ctx: ctx, extra: extra }) ]);
  } catch(_e){}
}
