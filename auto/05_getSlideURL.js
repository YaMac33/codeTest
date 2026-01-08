function exportDriveFilesToSheetById_withLog() {
  Logger.log('===== 処理開始 =====');

  // ===== 設定 =====
  const FOLDER_ID = '1vhp1a-6eUkWPTZ01jnpzCrO4fGbh0WHb';
  const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU';
  const SHEET_NAME = 'フォームの回答 1';

  const START_ROW = 2;
  const TITLE_COL = 2; // B列
  const URL_COL = 4;   // D列

  Logger.log(`フォルダID: ${FOLDER_ID}`);
  Logger.log(`スプレッドシートID: ${SPREADSHEET_ID}`);
  Logger.log(`シート名: ${SHEET_NAME}`);

  // ===== スプレッドシート取得 =====
  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ シートが見つかりません');
    throw new Error(`シート「${SHEET_NAME}」が見つかりません`);
  }

  Logger.log('✅ シート取得成功');

  // ===== フォルダ内ファイル取得 =====
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const titles = [];
  const urls = [];

  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    count++;

    const name = file.getName();
    const url = file.getUrl();

    Logger.log(`取得 ${count}: ${name}`);
    Logger.log(`URL: ${url}`);

    titles.push([name]);
    urls.push([url]);
  }

  Logger.log(`📄 取得ファイル数: ${count}`);

  // ===== シートに書き込み =====
  if (count > 0) {
    sheet.getRange(START_ROW, TITLE_COL, titles.length, 1).setValues(titles);
    sheet.getRange(START_ROW, URL_COL, urls.length, 1).setValues(urls);

    Logger.log(`✍ 書き込み完了: ${titles.length} 行`);
  } else {
    Logger.log('⚠ 書き込むファイルがありません');
  }

  Logger.log('===== 処理終了 =====');
}
