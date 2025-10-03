/**
 * StoreIO.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN save_store.gs (sha256:13eccef9431eaa84) ===== */

/** save_store.gs — 固定ヘッダ・完全一致の保存 */
/** save_store.gs — マッピング表をそのまま保存（ヘッダー含む全列） */
/** save_store.gs — 2_Mapping を保存時だけ親→子で C/E/G を補完して 3_Mapping_store へ出力 */
/** save_store.gs — 置き換え版
 *  マッピング表(A:I)をそのまま保存しつつ、保存直前に
 *  「定積」など銀行系の親コード（例：125/126）で補助が一意のとき、
 *  F/H（EPSON補助コード/補助名）を自動補完する。
 *  ※ E は書き換えない（空でもOK）。E が空なら G(親名)→コード逆引きで補完材料だけ取得。
 */
function saveMappingStore(){
  setupOrRepairSheets();
  var ss = SpreadsheetApp.getActive();
  var mapSh = ss.getSheetByName(SHEETS.MAPPING);
  if(!mapSh || mapSh.getLastRow()<2) throw new Error('2_Mapping が空です');

  // 2_Mapping の全行(A:I)をそのまま読み込み
  var last = mapSh.getLastRow();
  var vals = mapSh.getRange(2,1,last-1,9).getValues(); // A..I

  // 逆引き辞書：EPSON名→コード（Eが空の行の補助判断に使うだけ。E自体は書き換えない）
  var epNameToCode = _ep_nameToCode();  // epson_lookup.gs

  // 親コード→{nameToCode} のインデックス（補助の有無・件数判定用）
  var subsIdx = _subs_indexByParent();  // subs_index.gs

  // 親コードごとに「補助がちょうど1つ」だけ定義されているケースを拾う（125:しんきん, 126:けんしん など）
  function singleSubOf(parentCode){
    var rec = subsIdx[parentCode];
    if(!rec) return null;
    var names = Object.keys(rec.nameToCode);
    if(names.length === 1){
      var n = names[0];
      return { subName: n, subCode: String(rec.nameToCode[n]||'') };
    }
    return null;
  }

  // 出力用バッファ（A..Iをそのままコピーしつつ、必要ならF/Hだけ埋める）
  var outRows = [];
  for (var i=0; i<vals.length; i++){
    var row = vals[i].slice(); // A..I
    // 列対応： A:JDL科目コード B:JDL補助コード C:JDL科目名 D:JDL補助科目名
    //          E:EPSON科目コード F:EPSON補助コード G:EPSON科目名（選択） H:EPSON補助科目名 I:状態
    var eCode = (row[4]||'').toString().trim();  // E
    var fSub  = (row[5]||'').toString().trim();  // F
    var gName = (row[6]||'').toString().trim();  // G
    var hName = (row[7]||'').toString().trim();  // H

    // すでに補助が入っていれば触らない
    if (!fSub && !hName) {
      // E が空でも、G(親名)からコードを引けるなら引く（E 自体は保存しない＝変更しない）
      var parentCode = eCode || epNameToCode[gName] || '';
      if (parentCode) {
        var one = singleSubOf(parentCode);
        if (one) {
          // 125 → しんきん(1) / 126 → けんしん(1) …など、親コードに補助が1個だけの時は自動補完
          row[5] = one.subCode; // F:補助コード
          row[7] = one.subName; // H:補助名
        }
      }
    }

    outRows.push(row);
  }

  // 3_Mapping_store をヘッダそのままに全上書き
  var storeSh = ss.getSheetByName(SHEETS.MAP_STORE) || ss.insertSheet(SHEETS.MAP_STORE);
  storeSh.clear();
  storeSh.getRange(1,1,1,9).setValues([[
    'JDL科目コード','JDL補助コード','JDL科目名','JDL補助科目名',
    'EPSON科目コード','EPSON補助コード','EPSON科目名','EPSON補助科目名','状態'
  ]]);
  storeSh.setFrozenRows(1);
  if (outRows.length){
    storeSh.getRange(2,1,outRows.length,9).setValues(outRows);
  }
}
/** ===== END save_store.gs ===== */

