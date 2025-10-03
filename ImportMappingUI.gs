/**
 * ImportMappingUI.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN dv_helpers.gs (sha256:8878f64d0c1a13f9) ===== */
/** dv_helpers.gs (patched) — プルダウン生成（固定ヘッダ） */
function _ensureDvNames_(){
  var ss = SpreadsheetApp.getActive();
  var ch = ss.getSheetByName(SHEETS.EPSON);
  var dv = ss.getSheetByName(SHEETS.DV_NAMES); if(!dv) dv = ss.insertSheet(SHEETS.DV_NAMES);
  dv.clear();
  dv.getRange(1,1,1,1).setValues([['EPSON科目名一覧']]);
  var list = [];
  if(ch && ch.getLastRow()>=2){
    var last = ch.getLastRow();
    var names = ch.getRange(2, EPSON_CHART_COLS.name_display, Math.max(1, last-1), 1).getDisplayValues().map(function(r){return r[0];});
    list = names.filter(function(x){return !!x;});
  }
  if(list.length){
    dv.getRange(2,1,list.length,1).setValues(list.map(function(x){return [x];}));
  }else{
    dv.getRange(2,1,1,1).setValues([['']]);
  }
  dv.hideSheet();
  return dv.getRange(2,1,Math.max(1, list.length),1);
}

function _ensureDvSubs_(){
  // 返り値の互換性: { byCode: { [parentCode]: { range: Range } } }
  var ss = SpreadsheetApp.getActive();
  var base = ss.getSheetByName(SHEETS.EPSON_SUBS);
  var dv = ss.getSheetByName(SHEETS.DV_SUBS); if(!dv) dv = ss.insertSheet(SHEETS.DV_SUBS);
  dv.clear();
  dv.getRange(1,1,1,1).setValues([['EPSON補助名（親コード別）']]);

  var grouped = {};
  if (base && base.getLastRow() >= 2){
    var vals = base.getRange(2,1,base.getLastRow()-1,3).getValues(); // 親コード, 補助コード, 補助名
    vals.forEach(function(v){
      var p=(''+v[0]).trim(), n=(''+v[2]).trim();
      if(!p || !n) return;
      if(!grouped[p]) grouped[p]=[];
      if(grouped[p].indexOf(n)<0) grouped[p].push(n);
    });
  }

  var row = 2;
  var byCode = {};
  Object.keys(grouped).sort(function(a,b){return (''+a).localeCompare(''+b);}).forEach(function(code){
    var names = grouped[code];
    dv.getRange(row,1,1,1).setValue(code);
    row++;
    var rng = dv.getRange(row,1,Math.max(1, names.length),1);
    dv.getRange(row,1,Math.max(1, names.length),1).setValues((names.length?names:['']).map(function(x){return [x];}));
    byCode[code] = { range: rng };
    row += names.length + 1;
  });

  dv.hideSheet();
  return { byCode: byCode };
}

// DataValidation ヘルパー（互換）
function _newRuleList_(range){
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(true)
    .build();
}

function _setHValidationForRow_(sheet, rowIndex, parentCode, dvSubsMapOrObj, enable){
  var range = sheet.getRange(rowIndex, 8, 1, 1); // H列
  if (enable===false){ range.clearDataValidations(); return; }

  // 両対応: {byCode:{code:{range:Range}}} か、 {code: Range} かを許容
  var map = dvSubsMapOrObj;
  if (map && map.byCode) map = map.byCode;

  var entry = map && map[parentCode];
  var rng = entry && (entry.range || entry); // entry が Range の場合も許容

  if (rng && typeof rng.getA1Notation === 'function'){
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rng, true)
      .setAllowInvalid(true)
      .build();
    range.setDataValidation(rule);
  } else {
    range.clearDataValidations();
  }
}
/** ===== END dv_helpers.gs ===== */

/** ===== BEGIN mapping_grid.gs (sha256:775b23b4bb5ffc75) ===== */
/**
 * mapping_grid.gs（patched minimal）— 部分一致（黄色）復活＆銀行ロジックは既存踏襲
 * 依存：setupOrRepairSheets, SHEETS, IMPORT_HEADER_ROW, IMPORT_COLS, EPSON_CHART_COLS, FIXED_CODES
 *       _subs_indexByParent, _ensureDvSubs_, _setHValidationForRow_, _paintRowsFromMatrix_, _rankBest
 */

function buildMappingGrid(){
  setupOrRepairSheets();
  var ss = SpreadsheetApp.getActive();
  var imp = ss.getSheetByName(SHEETS.IMPORT);
  var map = ss.getSheetByName(SHEETS.MAPPING) || ss.insertSheet(SHEETS.MAPPING);
  if(!imp) throw new Error('1_Data_import がありません');

  var vals = imp.getDataRange().getValues();
  var start = IMPORT_HEADER_ROW + 1;

  function _ep_nameToCode_strict(){
    var m = {};
    var sh = ss.getSheetByName(SHEETS.EPSON);
    if(!sh || sh.getLastRow()<2) return m;
    var last = sh.getLastRow();
    var codes = sh.getRange(2, EPSON_CHART_COLS.code, last-1, 1).getDisplayValues();
    var names = sh.getRange(2, EPSON_CHART_COLS.name_display, last-1, 1).getDisplayValues();
    for (var i=0;i<codes.length;i++){
      var code=(''+codes[i][0]).trim(), name=(''+names[i][0]).trim();
      if(code && name && !m[name]) m[name]=code;
    }
    return m;
  }
  var epNameToCode = _ep_nameToCode_strict();
  var epNames = Object.keys(epNameToCode).sort();
  var subsByParent = _subs_indexByParent();

  // 銀行候補
  var bankFutsu   = epNames.filter(function(n){ return /^普通/.test(n); });
  var bankTetsumi = epNames.filter(function(n){ return /^定積/.test(n); });

  function bankStem(raw){
    var s=(''+(raw||'')).trim();
    s = s.replace(/株式会社|（株）|\(株\)/g,'').replace(/[ 　\t]+/g,'');
    s = s.replace(/長野県信用金庫/g,'長野信金').replace(/長野信用金庫/g,'長野信金');
    s = s.replace(/長野県信用組合/g,'長野県信').replace(/長野信用組合/g,'長野県信');
    var m = s.match(/^(.*?)(銀行|信用金庫|信用組合|信金)(.*?)(\d+)?$/);
    if(m){
      var base = m[1] + (m[2]==='信金' ? '信金' : (m[2]==='信用金庫' ? '信金' : (m[2]==='信用組合' ? '県信' : '')));
      var num  = m[4] || '';
      s = base + num;
    }
    return s;
  }
  function chooseExistingBankName(isFutsu, stem){
    var want = (isFutsu?'普通':'定積') + stem;
    var pool = isFutsu ? bankFutsu : bankTetsumi;
    for (var i=0;i<pool.length;i++){ var n=pool[i]; if(n===want || n.indexOf(want)===0) return n; }
    return (typeof _rankBest==='function') ? _rankBest(want, pool) : '';
  }

  // 親ごとに集計
  var bucket = new Map();
  function addRec(jCode,jName,jSub,jSubName){
    jCode=(jCode||'').toString().trim();
    jName=(jName||'').toString().trim();
    jSub =(jSub ||'').toString().trim();
    jSubName=(jSubName||'').toString().trim();
    if(!jName) return;
    var g=bucket.get(jName);
    if(!g){ g={parent:{jCode:jCode,jName:jName}, children:new Map(), hasNoSub:false, hasSub:false, hasMissingSub:false}; bucket.set(jName,g); }
    var hasAny = !!(jSub || jSubName);
    if(hasAny){
      g.hasSub=true;
      var key=jSub+'|'+jSubName;
      if(!g.children.has(key)) g.children.set(key,{jSub:jSub,jSubName:jSubName});
    }else{
      g.hasNoSub=true;
    }
  }
  for(var r=start-1;r<vals.length;r++){
    var row=vals[r];
    addRec(row[IMPORT_COLS.debitCode-1], row[IMPORT_COLS.debitName-1], row[IMPORT_COLS.dSubCode-1], row[IMPORT_COLS.dSubName-1]);
    addRec(row[IMPORT_COLS.creditCode-1],row[IMPORT_COLS.creditName-1],row[IMPORT_COLS.cSubCode-1], row[IMPORT_COLS.cSubName-1]);
  }
  bucket.forEach(function(g){ g.hasMissingSub = g.hasSub && g.hasNoSub; });

  var parents = Array.from(bucket.values()).map(function(g){ return g.parent; });
  parents.sort(function(a,b){
    return (''+(a.jCode||'')).localeCompare((''+(b.jCode||''))) || (a.jName||'').localeCompare(b.jName||'');
  });

  // 出力初期化
  map.clear();
  try{
    var whole = map.getRange(1,1,map.getMaxRows(),map.getMaxColumns());
    whole.clearDataValidations(); whole.clearFormat();
  }catch(_){}
  map.getRange(1,1,1,9).setValues([[
    'JDL科目コード','JDL補助コード','JDL科目名','JDL補助科目名',
    'EPSON科目コード','EPSON補助コード','EPSON科目名（選択）','EPSON補助科目名','状態'
  ]]);
  map.setFrozenRows(1);

  // 生成
  var rows=[];
  parents.forEach(function(par){
    var name = par.jName;
    var grp  = bucket.get(name);

    var eCode='', eName='', status='未選択';
    if(name==='買掛金' || name==='消耗品費'){
      eCode = FIXED_CODES[name] || ''; eName = name; status='完全一致';
    }else if(name==='普通預金' || name==='積立預金'){
      eCode=''; eName=''; status='未選択';
    }else if(name==='売上高'){
      eName='保険診療収入'; eCode=FIXED_CODES['保険診療収入']||''; status='不一致';
    }else if(epNameToCode[name]){
      eName=name; eCode=epNameToCode[name]; status='完全一致';
    }else{
      var best = (typeof _rankBest==='function') ? _rankBest(name, epNames) : '';
      if(best){ eName=best; eCode=epNameToCode[best]||''; status='部分一致'; }
    }
    rows.push([par.jCode,'', name,'', eCode,'', eName,'', status]);

    var kids = Array.from(grp.children.values());
    kids.sort(function(a,b){ return (a.jSub||'').localeCompare(b.jSub||'') || (a.jSubName||'').localeCompare(b.jSubName||''); });

    kids.forEach(function(k){
      var E='',F='',G='',H='',st='未選択';
      if(name==='買掛金' || name==='消耗品費'){
        E = FIXED_CODES[name] || '';
        G = '';
        var idx = subsByParent[E]||null;
        if(idx){
          if(idx.nameToCode[k.jSubName]){ H=k.jSubName; F=idx.nameToCode[k.jSubName]; st='完全一致'; }
          else{
            var nm = idx.synonymToName[k.jSubName]||'';
            if(nm && idx.nameToCode[nm]){ H=nm; F=idx.nameToCode[nm]; st='部分一致'; }
            else{ st='不一致'; }
          }
        }
      } else if(name==='普通預金' || name==='積立預金'){
        var stem = bankStem(k.jSubName);
        var cand = chooseExistingBankName(name==='普通預金', stem);
        if(cand){
          G = cand; E = epNameToCode[cand] || '';
          st = (cand.indexOf((name==='普通預金'?'普通':'定積') + stem)===0) ? '完全一致' : '部分一致';
        }else{
          G=''; E=''; st='未選択';
        }
      } else {
        E = eCode; G = eName;
        var idx2 = subsByParent[E]||null;
        if(idx2){
          if(idx2.nameToCode[k.jSubName]){ H=k.jSubName; F=idx2.nameToCode[k.jSubName]; st='完全一致'; }
          else{
            var nm2 = idx2.synonymToName[k.jSubName]||'';
            if(nm2 && idx2.nameToCode[nm2]){ H=nm2; F=idx2.nameToCode[nm2]; st='部分一致'; }
            else{
              if(!E && !G){
                var best2 = (typeof _rankBest==='function') ? _rankBest(name, epNames) : '';
                if(best2){ G=best2; E=epNameToCode[best2]||''; st='部分一致'; }
              }else{ st='不一致'; }
            }
          }
        }
      }
      rows.push([par.jCode, k.jSub, '', k.jSubName, E, F, G, H, st]);
    });

    if (grp.hasSub && grp.hasMissingSub){
      var E2='',G2='',st2='未選択';
      if(name==='買掛金'){ E2=FIXED_CODES['買掛金']; G2=''; }
      else if(name==='消耗品費'){ E2=FIXED_CODES['消耗品費']; G2=''; }
      else if(name==='普通預金' || name==='積立預金'){
        var first = kids[0];
        if(first){
          var stem2 = bankStem(first.jSubName);
          var cand2 = chooseExistingBankName(name==='普通預金', stem2);
          if(cand2){ G2=cand2; E2=epNameToCode[cand2]||''; st2='完全一致'; }
        }
      }else{
        if(eCode || eName){ E2=eCode; G2=eName; st2=(eCode? '完全一致' : (G2? '部分一致':'未選択')); }
      }
      rows.push([par.jCode,'','', '[補助なし]', E2,'',G2,'', st2]);
    }
  });

  // ← 部分一致の再採点で黄色が落ちないように最終確認（必要なら）
  if (typeof _recalcStatusesForRows_ === 'function') {
    rows = _recalcStatusesForRows_(rows);
  }

  if(rows.length){
    map.getRange(2,1,rows.length,9).setValues(rows);
    _paintRowsFromMatrix_(map, 2, rows);
  }

  // プルダウン
  if(rows.length){
    var epSh = ss.getSheetByName(SHEETS.EPSON);
    var gRange = epSh.getRange(2, EPSON_CHART_COLS.name_display, Math.max(1, epSh.getLastRow()-1), 1);
    var ruleG = SpreadsheetApp.newDataValidation().requireValueInRange(gRange, true).setAllowInvalid(true).build();
    map.getRange(2,7,rows.length,1).setDataValidation(ruleG);

    var dvSubs = _ensureDvSubs_();
    for (var i=0;i<rows.length;i++){
      var r = 2+i;
      var jSubName = (rows[i][3]||'').toString();
      if (jSubName==='[補助なし]'){ _setHValidationForRow_(map, r, '', dvSubs, false); continue; }
      var eCodeRow = (rows[i][4]||'').toString().trim();
      _setHValidationForRow_(map, r, eCodeRow, dvSubs, true);
    }
  }
}
/** ===== END mapping_grid.gs ===== */

/** ===== BEGIN on_edit_fix.gs (sha256:b132310ae3998ab4) ===== */
/** on_edit_fix.gs — 上書き版：部分一致(>=2連続)を判定して色付けまで反映。銀行ロジックは不変更 **/

function _norm_(s){
  return (s||'').toString().normalize('NFKC')
    .replace(/[ 　\t]+/g,'').replace(/[()（）［］【】]/g,'').trim();
}

/** 連続LCS長（2以上で部分一致扱い） */
function _lcsContig_(a, b){
  a = _norm_(a); b = _norm_(b);
  var n=a.length, m=b.length; if(!n||!m) return 0;
  var prev = new Array(m+1).fill(0), best=0;
  for (var i=1;i<=n;i++){
    var curr = new Array(m+1).fill(0);
    for (var j=1;j<=m;j++){
      if (a[i-1]===b[j-1]) { curr[j]=prev[j-1]+1; if(curr[j]>best) best=curr[j]; }
    }
    prev=curr;
  }
  return best;
}

/** JDL名 vs EPSON名 → ステータス */
function _statusByPair_(jdlName, epName){
  var nj=_norm_(jdlName||''), ne=_norm_(epName||'');
  if (!nj && !ne) return '未選択';
  if (nj && ne && nj===ne) return '完全一致';
  if (nj && ne && _lcsContig_(nj, ne)>=2) return '部分一致';
  return '不一致';
}

/** EPSONコード→表示名（Epson_chart固定列） */
function _ep_codeToName_strict_(){
  var m = {};
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
  if(!sh || sh.getLastRow()<2) return m;
  var last = sh.getLastRow();
  var codes = sh.getRange(2, EPSON_CHART_COLS.code, last-1, 1).getDisplayValues();
  var names = sh.getRange(2, EPSON_CHART_COLS.name_display, last-1, 1).getDisplayValues();
  for (var i=0;i<codes.length;i++){
    var code=(''+codes[i][0]).trim(), name=(''+names[i][0]).trim();
    if(code && name && !m[code]) m[code]=name;
  }
  return m;
}

function _isFourParent_(parentName){
  return parentName==='買掛金' || parentName==='消耗品費' || parentName==='普通預金' || parentName==='積立預金';
}

/** 既存の色付けロジックをそのまま使う */
function _statusColor_(status){
  if (status==='完全一致') return COLORS.statusBlue;
  if (status==='部分一致') return COLORS.statusYellow;
  if (status==='不一致')   return COLORS.statusRed;
  return null;
}
function _paintSingleRowStrict_(sh, r){
  var vals = sh.getRange(r,1,1,9).getValues()[0];
  var isSub = ((vals[2]||'').toString().trim()==='' && (vals[3]||'').toString().trim()!=='') || (vals[3]==='[補助なし]');
  var a2hColor = isSub ? COLORS.subRow : null;
  var iColor   = _statusColor_(vals[8]);
  sh.getRange(r,1,1,8).setBackgrounds([[a2hColor,a2hColor,a2hColor,a2hColor,a2hColor,a2hColor,a2hColor,a2hColor]]);
  sh.getRange(r,9,1,1).setBackgrounds([[iColor]]);
}

/** ===== onEdit 本体（既存列前提） ===== */
function onEdit(e){
  try{
    var sh=e.range.getSheet();
    if(!sh || sh.getName()!==SHEETS.MAPPING) return;
    var r=e.range.getRow(), c=e.range.getColumn(); if(r<=1) return;

    var epCodeToName = _ep_codeToName_strict_();

    var jdlPar = (sh.getRange(r,3).getValue()+'').trim(); // C
    var jdlSub = (sh.getRange(r,4).getValue()+'').trim(); // D
    var eCode  = (sh.getRange(r,5).getValue()+'').trim(); // E
    var fSub   = (sh.getRange(r,6).getValue()+'').trim(); // F
    var gName  = (sh.getRange(r,7).getValue()+'').trim(); // G
    var hName  = (sh.getRange(r,8).getValue()+'').trim(); // H

    var isSub  = ((jdlPar||'')==='') && ((jdlSub||'')!=='') || jdlSub==='[補助なし]';
    var isFour = _isFourParent_(jdlPar) && !isSub;

    // G変更：コード反映（4科目の親だけEは空のまま）
    if(c===7){
      var epNameToCode = (function(){
        var m={};
        var eps  = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPSON);
        if(!eps||eps.getLastRow()<2) return m;
        var last = eps.getLastRow();
        var codes= eps.getRange(2, EPSON_CHART_COLS.code, last-1, 1).getDisplayValues();
        var names= eps.getRange(2, EPSON_CHART_COLS.name_display, last-1, 1).getDisplayValues();
        for(var i=0;i<codes.length;i++){
          var cd=(''+codes[i][0]).trim(), nm=(''+names[i][0]).trim();
          if(cd && nm && !m[nm]) m[nm]=cd;
        }
        return m;
      })();
      var newE = epNameToCode[gName] || '';
      sh.getRange(r,5).setValue(isFour ? '' : newE);

      // ステータス更新
      var status = isSub
        ? _statusByPair_(jdlSub, hName)  // 子行は補助名ベース
        : (isFour ? '未選択' : _statusByPair_(jdlPar, gName));
      sh.getRange(r,9).setValue(status);

      // Hの検証（[補助なし]は検証なし）
      var dvSubs = _ensureDvSubs_();
      var jSubNm = jdlSub;
      if (jSubNm==='[補助なし]'){ _setHValidationForRow_(sh, r, '', dvSubs.byCode, false); }
      else { _setHValidationForRow_(sh, r, (isFour? '' : (newE||'')), dvSubs.byCode, true); }

      _paintSingleRowStrict_(sh, r);
      return;
    }

    // E変更：親コード手入力時の再判定
    if(c===5){
      var epName = gName || (epCodeToName[eCode]||'');
      var status = isSub
        ? _statusByPair_(jdlSub, hName)
        : (isFour ? '未選択' : _statusByPair_(jdlPar, epName));
      sh.getRange(r,9).setValue(status);

      var dvSubs = _ensureDvSubs_();
      if (jdlSub==='[補助なし]'){ _setHValidationForRow_(sh, r, '', dvSubs.byCode, false); }
      else { _setHValidationForRow_(sh, r, (isFour? '' : eCode), dvSubs.byCode, true); }

      _paintSingleRowStrict_(sh, r);
      return;
    }

    // H変更：補助名の部分一致判定
    if(c===8){
      var status = _statusByPair_(jdlSub, hName);
      sh.getRange(r,9).setValue(status);
      _paintSingleRowStrict_(sh, r);
      return;
    }

    if(c===3 || c===4 || c===9){
      if(isFour && !isSub){ sh.getRange(r,9).setValue('未選択'); }
      _paintSingleRowStrict_(sh, r);
    }
  }catch(_){}
}
/** ===== END on_edit_fix.gs ===== */

