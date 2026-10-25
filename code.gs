// キャラ箱庭 Drive連携 — Google Apps Script
// フォルダ「ご当地キャラを作ろう！」用（フォルダID設定済み）
// script.google.com に貼り付けて「ウェブアプリ」としてデプロイしてください。
// （手順は DRIVE連携_セットアップ.md 参照）
//
// 4つの役割:
//   GET  …/exec           → キャラ一覧のJSON
//   GET  …/exec?img=ID    → 画像本体をbase64のdata:URLテキストで返す
//                            （CORS制限を回避し、アプリ側で白抜きできるようにする）
//   GET  …/exec?names=1   → フォルダ内の生ファイル名一覧（アップローダーの重複チェック用）
//   POST …/exec           → 画像アップロード受付（uploader.html から使用）
//                            body: {"upload":true,"key":"<UPLOAD_KEY>","name":"...","mime":"image/png","data":"<base64>"}
//                            ※スクリプトプロパティ UPLOAD_KEY が未設定だと全て拒否されます

const FOLDER_ID = '1KoexIA7NFdKlvBOOSu5aVyOwQoINUqqV'; // ご当地キャラを作ろう！

function doGet(e) {
  if (e && e.parameter && e.parameter.img) {
    return serveImage(e.parameter.img);
  }
  if (e && e.parameter && e.parameter.names) {
    return serveNames();
  }
  return serveList();
}

// アップロード制限
const ALLOWED_UPLOAD_MIMES = ['image/png', 'image/jpeg', 'image/webp'];
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024; // 8MB（デコード後）

function doPost(e) {
  let out = { ok: false };
  try {
    const req = JSON.parse(e.postData.contents);
    if (!req || !req.upload) {
      out = { ok: false, error: 'bad request' };
    } else if (!checkUploadKey(req.key)) {
      out = { ok: false, error: 'unauthorized' };
    } else if (ALLOWED_UPLOAD_MIMES.indexOf(req.mime) === -1) {
      out = { ok: false, error: 'unsupported type' };
    } else if (typeof req.data !== 'string' || !req.data ||
               req.data.length > Math.ceil(MAX_UPLOAD_BYTES / 3) * 4 + 4) {
      out = { ok: false, error: 'too large' };
    } else {
      const name = sanitizeName(req.name);
      const bytes = Utilities.base64Decode(req.data);
      if (!name) {
        out = { ok: false, error: 'bad name' };
      } else if (bytes.length > MAX_UPLOAD_BYTES) {
        out = { ok: false, error: 'too large' };
      } else {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        if (folder.getFilesByName(name).hasNext()) {
          out = { ok: true, status: 'skipped' }; // 同名ファイルは重複追加しない
        } else {
          folder.createFile(Utilities.newBlob(bytes, req.mime, name));
          out = { ok: true, status: 'created' };
        }
      }
    }
  } catch (err) {
    console.error(err);
    out = { ok: false, error: 'server error' };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// スクリプトプロパティ UPLOAD_KEY と照合。未設定なら常に拒否（フェイルクローズ）
function checkUploadKey(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('UPLOAD_KEY');
  if (!expected || typeof key !== 'string' || key.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) {
    diff |= expected.charCodeAt(i) ^ key.charCodeAt(i);
  }
  return diff === 0;
}

// パス区切り・制御文字などを除去し、最大100文字に切り詰める
function sanitizeName(name) {
  if (typeof name !== 'string') return '';
  let s = name.replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_')
              .replace(/^[.\s]+/, '')
              .trim();
  if (s.length > 100) {
    const m = s.match(/\.[A-Za-z0-9]{1,5}$/);
    const ext = m ? m[0] : '';
    s = s.slice(0, 100 - ext.length) + ext;
  }
  return s;
}

function serveNames() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) out.push(f.getName());
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function serveList() {
  const base = ScriptApp.getService().getUrl(); // この…/exec自身のURL
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf('image/') === 0) {
      const id = f.getId();
      out.push({
        name: displayName(f.getName()),
        district: districtName(f.getName()),                             // 地区名（生息地の振り分け用）
        url: base + '?img=' + id,                                        // 白抜き用（base64）
        b64: true,
        thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w512', // 失敗時フォールバック
        id: id,
        time: f.getDateCreated().getTime()
      });
    }
  }
  out.sort(function (a, b) { return a.time - b.time; }); // 生成時間の古い順
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// FOLDER_ID 直下の画像ファイルだけを返す（それ以外のIDは not found）
function serveImage(id) {
  try {
    const file = DriveApp.getFileById(String(id));
    let inFolder = false;
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === FOLDER_ID) { inFolder = true; break; }
    }
    if (!inFolder || String(file.getMimeType()).indexOf('image/') !== 0) {
      return ContentService.createTextOutput('error: not found')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    const blob = file.getBlob();
    const dataUrl = 'data:' + blob.getContentType() + ';base64,' +
      Utilities.base64Encode(blob.getBytes());
    return ContentService.createTextOutput(dataUrl)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('error: not found')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// 「地区_キャラ名_20260708_151704.png」→「地区」
function districtName(fileName) {
  const base = fileName.replace(/\.[^.]+$/, '');
  const parts = base.split('_');
  return parts.length >= 2 ? parts[0] : '';
}

// 「地区_キャラ名_20260708_151704.png」→「キャラ名」
// キャラ名が「なまえなし」のときは地区名を使う
function displayName(fileName) {
  const base = fileName.replace(/\.[^.]+$/, '');
  const parts = base.split('_');
  if (parts.length >= 2) {
    const chara = parts[1];
    if (chara && chara !== 'なまえなし') return chara;
    return parts[0] || base;
  }
  return base;
}
