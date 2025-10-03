/**
 * Download.gs
 * このファイルは既存の .gs を再配置した結合ファイルです（機能変更なし／関数本文は原文のまま）。
 * 生成元: test-main.zip /test-main/*.gs
 */

/** ===== BEGIN Download.gs (sha256:2e4c4df315f738fa) ===== */
/**
 * 現在の Apps Script プロジェクトを ZIP にして My Drive に保存します。
 * - まず Apps Script API v1 でプロジェクトの「content」を取得
 * - ZIP 化
 * - DriveApp.createFile をリトライ
 * - それでも失敗したら Drive API(v3) multipart でフォールバック保存
 *
 * 使い方：エディタでこの関数を実行 → ログに保存URLが出ます。
 *
 * 事前準備：
 * 1) エディタ右上の「歯車」→「Google Cloud プロジェクト」→コンソールで
 *    対象 GCP プロジェクトの「Apps Script API」を有効化。
 * 2) エディタの「サービス（+）」→「Drive API」を ON（高度なサービス）。
 */
function exportThisProjectAsZip() {
  const scriptId = ScriptApp.getScriptId();
  const files = fetchProjectContent_(scriptId); // Apps Script API v1

  // type → 拡張子
  const ext = (t) => t === 'SERVER_JS' ? 'gs' : (t === 'HTML' ? 'html' : 'json');
  // ファイル名のサニタイズ（Windows/ZIP互換）
  const safe = (s) => String(s || '').replace(/[\\/:*?"<>|]/g, '_');

  // Blob 群を作成（manifest 名だけは固定 appsscript.json）
  const blobs = files.map(f => {
    const filename =
      f.type === 'JSON' && f.name === 'appsscript'
        ? 'appsscript.json'
        : `${safe(f.name)}.${ext(f.type)}`;
    return Utilities.newBlob(f.source || '', MimeType.PLAIN_TEXT, filename);
  });

  // ZIP 作成
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const projName = getProjectNameSafe_(scriptId);
  let zipBlob = Utilities.zip(blobs, `${projName}_${now}.zip`).setContentType('application/zip');
  Logger.log('ZIP bytes: %s', zipBlob.getBytes().length);

  // DriveApp.createFile をリトライ（指数バックオフ）
  const fileUrl = saveZipWithRetry_(zipBlob);
  if (fileUrl) {
    Logger.log('ZIP saved: %s', fileUrl);
    return;
  }

  // だめなら Drive API multipart でフォールバック
  const url2 = saveZipViaDriveApi_(zipBlob, `${projName}_${now}.zip`);
  Logger.log('ZIP saved via Drive API: %s', url2);
}

/**
 * Apps Script API v1 でプロジェクトの content を取得
 * 失敗時は詳細を投げます。
 */
function fetchProjectContent_(scriptId) {
  const url = `https://script.googleapis.com/v1/projects/${encodeURIComponent(scriptId)}/content`;
  const token = ScriptApp.getOAuthToken();

  // 軽いリトライ
  const MAX_TRY = 4, BASE = 400;
  let lastErr;
  for (let i=0; i<MAX_TRY; i++){
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      const body = res.getContentText();
      if (code !== 200) {
        throw new Error(`GET content failed: ${code}\n${body}`);
      }
      /** @type {{files:Array<{name:string,type:"SERVER_JS"|"HTML"|"JSON",source:string}>}} */
      const payload = JSON.parse(body);
      if (!payload || !payload.files || !payload.files.length) {
        throw new Error('No files returned by Script API.');
      }
      return payload.files;
    } catch (e) {
      lastErr = e;
      Utilities.sleep(BASE * Math.pow(2, i));
    }
  }
  throw new Error(String(lastErr && lastErr.message || lastErr));
}

/** プロジェクト名を Drive から取得（リトライ付き） */
function getProjectNameSafe_(scriptId){
  for (let i=0; i<3; i++){
    try {
      return String(DriveApp.getFileById(scriptId).getName()).replace(/[\\/:*?"<>|]/g, '_');
    } catch (_){
      Utilities.sleep(250);
    }
  }
  return 'AppsScriptProject';
}

/**
 * DriveApp.createFile による保存（指数バックオフで複数回トライ）
 * 成功時はURL、失敗時は空文字を返す
 */
function saveZipWithRetry_(zipBlob){
  const MAX_TRY_SAVE = 6, BASE_MS_SAVE = 700, JITTER_SAVE = 350;
  for (let i=0; i<MAX_TRY_SAVE; i++){
    try {
      const file = DriveApp.createFile(zipBlob);
      return file.getUrl();
    } catch (e){
      const wait = BASE_MS_SAVE * Math.pow(2, i) + Math.floor(Math.random()*JITTER_SAVE);
      Utilities.sleep(wait);
      if (i === MAX_TRY_SAVE - 1) {
        Logger.log('createFile failed (final): %s', e && e.message);
      }
    }
  }
  return '';
}

/**
 * Drive API (v3) で multipart アップロードして保存（高度なサービス Drive を ON にしておく）
 * 成功時は URL を返す。失敗時は Error を投げる。
 */
function saveZipViaDriveApi_(zipBlob, filename){
  // 高度なサービス Drive.Files.create でも良いが、
  // 標準 UrlFetchApp で multipart を明示的に投げる（権限/挙動の差異を避ける）
  const boundary = 'xxxxxxxxxx' + (new Date().getTime());
  const delimiter = '--' + boundary + '\r\n';
  const closeDelim = '\r\n--' + boundary + '--';

  const metadata = {
    name: filename,
    mimeType: 'application/zip'
  };

  // multipart ボディを組み立て
  const metaPart =
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n';
  const mediaPartHeader = 'Content-Type: application/zip\r\n\r\n';

  // パーツ連結
  const bodyHead = Utilities.newBlob(
    delimiter + metaPart + delimiter + mediaPartHeader
  ).getBytes();
  const bodyTail = Utilities.newBlob(closeDelim).getBytes();
  const payload = concatenateBytes_([bodyHead, zipBlob.getBytes(), bodyTail]);

  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'multipart/related; boundary=' + boundary,
    payload: payload,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Drive API upload failed: ' + code + '\n' + res.getContentText());
  }
  const json = JSON.parse(res.getContentText());
  return 'https://drive.google.com/file/d/' + json.id + '/view';
}

/** Uint8Array の配列を連結して 1 本の byte[] にする */
function concatenateBytes_(parts){
  // Apps Script の getBytes() は Java byte[] 相当（数値配列）なので連結して返せばOK
  var total = 0;
  for (var i=0; i<parts.length; i++) total += parts[i].length;

  var out = new Uint8Array(total);
  var offset = 0;
  for (var j=0; j<parts.length; j++){
    out.set(parts[j], offset);
    offset += parts[j].length;
  }
  return out;
}
/** ===== END Download.gs ===== */

