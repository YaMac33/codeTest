function runAllTasks() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    Logger.log('=== Step 1: 原稿生成を開始 ===');
    generateSpeakerNotes();

    SpreadsheetApp.flush();
    Utilities.sleep(10000);

    Logger.log('=== Step 2: スライド転記を開始 ===');
    importSpeakerNotesFromSpreadsheet();

    SpreadsheetApp.flush();
    Utilities.sleep(10000);

    Logger.log('=== Step 3: ブログ生成を開始 ===');
    generateNoteBlogArticles();

    Logger.log('=== 全工程が完了しました ===');
  } finally {
    lock.releaseLock();
  }
}
