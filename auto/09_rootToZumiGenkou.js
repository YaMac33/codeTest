/**
 * スプレッドシートのH列（2行目以降）に記載された
 * GoogleドキュメントIDのみを
 * 「使用済原稿フォルダ」に移動する
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU';
const SHEET_NAME = ''; // 空なら先頭シート
const USED_FOLDER_ID = '1sHXWVlKHVZzyEM8gikOhulVOH2-aRBnf';

const COL_SCRIPT_DOC_ID = 8; // H列（A=1）

/**
 * メイン処理
 */
function moveUsedScriptDocs() {

  // スプレッドシート取得
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = SHEET_NAME
    ? spreadsheet.getSheetByName(SHEET_NAME)
    : spreadsheet.getSheets()[0];

  if (!sheet) {
    Logger.log('エラー: シートが見つかりません');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('処理対象のデータがありません');
    return;
  }

  // H列のIDを取得
  const docIds = sheet
    .getRange(2, COL_SCRIPT_DOC_ID, lastRow - 1, 1)
    .getValues();

  const usedFolder = DriveApp.getFolderById(USED_FOLDER_ID);
  const rootFolder = DriveApp.getRootFolder();

  docIds.forEach((row, index) => {
    const docId = row[0];
    const rowNumber = index + 2;

    // IDがない行はスキップ
    if (!docId) {
      Logger.log(`行 ${rowNumber}: スキップ（IDなし）`);
      return;
    }

    try {
      const file = DriveApp.getFileById(docId);

      // 念のため Googleドキュメントのみ許可
      if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
        Logger.log(`行 ${rowNumber}: スキップ（Googleドキュメントではない）`);
        return;
      }

      // 使用済フォルダに追加
      usedFolder.addFile(file);

      // 元の親フォルダから削除（完全移動）
      const parents = file.getParents();
      while (parents.hasNext()) {
        parents.next().removeFile(file);
      }

      Logger.log(`行 ${rowNumber}: 移動完了 - ${file.getName()}`);

    } catch (e) {
      Logger.log(`行 ${rowNumber}: エラー - ${e.message}`);
    }
  });

  Logger.log('使用済原稿の移動処理が完了しました');
}
