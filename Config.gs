/**
 * Config.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN constants.gs (sha256:aa1850ccc42a8667) ===== */

/** constants.gs */
var SHEETS = {
  IMPORT: '1_Data_import',
  MAPPING: '2_Mapping',
  MAP_STORE: '3_Mapping_store',
  CONVERTED: '4_Converted',
  DEBUG_SUBJ: 'Debug_subjects',
  EPSON: 'Epson_chart',
  EPSON_SUBS: 'Epson_subs',
  DV_NAMES: '_DV_EPNAMES',
  DV_SUBS: '_DV_EPSUBS',
  LOGS: 'Logs'
};

var COLORS = {
  subRow:       '#E8F5E9', // 補助行（A~H）
  statusBlue:   '#BBDEFB', // 完全一致
  statusYellow: '#FFF9C4', // 部分一致
  statusRed:    '#FFCDD2'  // 不一致
};

// 固定ルール：親扱い
var PARENTS_WITH_CHILDREN_SPECIAL = ['普通預金','積立預金','買掛金','消耗品費'];

// 固定列インデックス（1_Data_import）— 4行目がヘッダ、5行目～データ
var IMPORT_HEADER_ROW = 4;
var IMPORT_COLS = {
  // 1-based
  date: 3,
  debitCode: 4,
  debitName: 5,
  dSubCode: 7,
  dSubName: 8,
  dTax: 10,
  debitAmt: 12,
  creditCode: 14,
  creditName: 15,
  cSubCode: 17,
  cSubName: 18,
  cTax: 20,
  creditAmt: 22,
  memo: 24
};

// Epson_chart 固定列（1行目ヘッダ）
var EPSON_CHART_COLS = {
  code: 1,              // コード
  name_display: 2       // 正 式 科 目 名（表示名として採用）
  // 略称や他の列は参照しない
};

// Epson_subs 固定列（1行目ヘッダ）
var EPSON_SUBS_COLS = {
  parentCode: 1,
  subCode: 2,
  subName: 3,
  synonyms: 4 // 「同義語(カンマ区切り)」— 完全一致トークンのみ許容（任意）
};

// 6121=売上高 → EPSON「保険診療収入」に固定（Epson_chartの“正 式 科 目 名”参照）
var FORCE_SALES_TO = '保険診療収入';

// === ステータス定数（塗り分け用） ===
const STATUS = {
  MATCH: '完全一致',
  PARTIAL: '部分一致',
  UNSELECT: '未選択',
  MISMATCH: '不一致',
};
/** ===== END constants.gs ===== */

/** ===== BEGIN ep_cols_const.gs (sha256:a5d345ed9979890e) ===== */
/** ep_cols_const.gs — 固定ヘッダの列番号（1始まり） */
var IMPORT_HEADER_ROW = 4; // 1_Data_import のヘッダは4行目（固定）

// 1_Data_import（あなた指定の固定ヘッダに厳密対応）
var IMPORT_COLS = {
  date: 3,
  debitCode: 4,   debitName: 5,   dSubCode: 7,  dSubName: 8,
  // ▼ 修正：税区分列は I（9）/ S（19）
  dTax: 9,        debitAmt: 12,
  creditCode: 14, creditName: 15, cSubCode: 17, cSubName: 18,
  cTax: 19,       creditAmt: 22,
  memo: 24
};

// Epson_chart（固定：A=コード, B=正 式 科 目 名）
var EPSON_CHART_COLS = {
  code: 1,
  name_display: 2
};

// 固定コード（厳格）
var FIXED_CODES = {
  '買掛金': '201',
  '消耗品費': '523',
  '保険診療収入': '810' // 6121=売上高→保険診療収入
};
/** ===== END ep_cols_const.gs ===== */

