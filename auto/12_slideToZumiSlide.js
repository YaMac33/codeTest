/**
 * スプレッドシートのI列（2行目以降）に記載された
 * GoogleスライドIDのみを
 * 「原稿ありフォルダ」に移動する
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU';
const SHEET_NAME = ''; // 空なら先頭シート
const SCRIPTED_SLIDE_FOLDER_ID = '1KqLbSMfTPFC4oiXV4iuWvc6gECj36Yna';

const COL_SLIDE_ID = 9; // I列（A=1基準）

/**
 * メイン処理
 */
function moveSlidesWithScript() {

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

  // I列のスライドIDを取得
  const slideIds = sheet
    .getRange(2, COL_SLIDE_ID, lastRow - 1, 1)
    .getValues();

  const targetFolder = DriveApp.getFolderById(SCRIPTED_SLIDE_FOLDER_ID);

  slideIds.forEach((row, index) => {
    const slideId = row[0];
    const rowNumber = index + 2;

    // IDがない行はスキップ
    if (!slideId) {
      Logger.log(`行 ${rowNumber}: スキップ（スライドIDなし）`);
      return;
    }

    try {
      const file = DriveApp.getFileById(slideId);

      // 念のため Googleスライドのみ対象
      if (file.getMimeType() !== MimeType.GOOGLE_SLIDES) {
        Logger.log(`行 ${rowNumber}: スキップ（Googleスライドではない）`);
        return;
      }

      // 原稿ありフォルダに追加
      targetFolder.addFile(file);

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

  Logger.log('原稿ありスライドの移動処理が完了しました');
}
