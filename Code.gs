// ここだけ名称を統一
const MENU_NAME = '画像一括ダウンロードツール';

function onOpen() {
  DocumentApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('画像一括ダウンロード', 'showSidebar') // ← メニュー項目名
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('画像一括ダウンロードツール'); // ← サイドバーの見出し
  DocumentApp.getUi().showSidebar(html);
}

// 画像の総数
function countImages() {
  return _collectAllImageBlobs().length;
}

/**
 * 画像をバッチでBase64返却
 * @param {number} offset 0-based
 * @param {number} limit
 * @return {{ total:number, items: Array<{mime:string, b64:string}> }}
 */
function getImagesBatch(offset, limit) {
  const blobs = _collectAllImageBlobs();
  const total = blobs.length;
  const start = Math.max(0, offset|0);
  const end   = Math.min(total, start + (limit|0 || 10));
  const items = [];
  for (let i = start; i < end; i++) {
    const blob = blobs[i];
    items.push({
      mime: blob.getContentType(),
      b64: Utilities.base64Encode(blob.getBytes()),
    });
  }
  return { total, items };
}

// 内部：全画像を列挙して Blob[] を返す
function _collectAllImageBlobs() {
  const doc   = DocumentApp.getActiveDocument();
  const body  = doc.getBody();
  const out   = [];

  // 本文のインライン画像
  body.getImages().forEach(img => out.push(img.getBlob()));

  // ヘッダー・フッター
  const header = doc.getHeader();
  if (header) header.getImages().forEach(img => out.push(img.getBlob()));
  const footer = doc.getFooter();
  if (footer) footer.getImages().forEach(img => out.push(img.getBlob()));

  // 段落にアンカーされた配置画像
  body.getParagraphs().forEach(p => {
    p.getPositionedImages().forEach(pi => out.push(pi.getBlob()));
  });

  return out;
}
