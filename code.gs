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
//                            body: {"upload":true,"name":"...","mime":"image/png","data":"<base64>"}

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

function doPost(e) {
  let out = { ok: false };
  try {
    const req = JSON.parse(e.postData.contents);
    if (req && req.upload && req.name && req.data) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      if (folder.getFilesByName(req.name).hasNext()) {
        out = { ok: true, status: 'skipped' }; // 同名ファイルは重複追加しない
      } else {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(req.data), req.mime || 'image/png', req.name);
        folder.createFile(blob);
        out = { ok: true, status: 'created' };
      }
    } else {
      out = { ok: false, error: 'bad request' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
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

function serveImage(id) {
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const dataUrl = 'data:' + blob.getContentType() + ';base64,' +
      Utilities.base64Encode(blob.getBytes());
    return ContentService.createTextOutput(dataUrl)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
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
