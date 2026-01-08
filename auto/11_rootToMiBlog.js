/**
 * G列に入力されたブログ記事ドキュメントIDを
 * 未投稿記事フォルダに移動する
 */
function moveBlogDocsToUnpostedFolder() {
  const SHEET_NAME = 'シート1'; // ←必要に応じて変更
  const COL_BLOG_DOC_ID = 7; // G列
  const UNPOSTED_FOLDER_ID = '1KYutXxQUaMuOZVs34Apt94FLYMI7VoBh';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const folder = DriveApp.getFolderById(UNPOSTED_FOLDER_ID);

  for (let row = 2; row <= lastRow; row++) { // 1行目はヘッダー想定
    const docId = sheet.getRange(row, COL_BLOG_DOC_ID).getValue();

    if (!docId) continue;

    try {
      const file = DriveApp.getFileById(docId);

      // 移動（= フォルダに追加）
      folder.addFile(file);

      // ルートフォルダからは削除（完全に「移動」したい場合）
      DriveApp.getRootFolder().removeFile(file);

    } catch (e) {
      Logger.log(`行${row}のドキュメント移動失敗: ${e.message}`);
    }
  }

  Logger.log('未投稿記事フォルダへの移動が完了しました');
}
