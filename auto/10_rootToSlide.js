/**
 * ルートフォルダ直下にある Googleスライドを
 * 指定した「Slideフォルダ」にすべて移動する
 */
function moveRootSlidesToSlideFolder() {

  // ===== 設定 =====
  const TARGET_FOLDER_ID = '1vhp1a-6eUkWPTZ01jnpzCrO4fGbh0WHb'; // SlideフォルダID
  const SLIDE_MIME_TYPE = MimeType.GOOGLE_SLIDES;

  // ルートフォルダ & 移動先フォルダ
  const rootFolder = DriveApp.getRootFolder();
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);

  // ルート直下の Googleスライドを取得
  const files = rootFolder.getFilesByType(SLIDE_MIME_TYPE);

  let count = 0;

  while (files.hasNext()) {
    const file = files.next();

    // 移動（＝追加 → ルートから削除）
    targetFolder.addFile(file);
    rootFolder.removeFile(file);

    Logger.log(`移動完了: ${file.getName()}`);
    count++;
  }

  Logger.log(`処理完了。移動したスライド数: ${count}`);
}
