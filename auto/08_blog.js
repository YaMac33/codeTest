/**
 * 原稿 + スライドを元にブログ記事を生成し、
 * ①ブログ記事Docを作成
 * ②G列にブログ記事IDを書き込み
 * ③未投稿記事フォルダへ移動
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU';
const SHEET_NAME = 'フォームの回答 1';

const COL_BLOG_DOC_ID = 7; // G列
const COL_SCRIPT_DOC_ID = 8; // H列（原稿）
const COL_SLIDE_ID = 9; // I列（スライド）

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('YOUR_GEMINI_API_KEY');
const UNPOSTED_FOLDER_ID = '1KYutXxQUaMuOZVs34Apt94FLYMI7VoBh';

// =================

function generateBlogArticlesAndMove() {
  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName(SHEET_NAME);

  if (!sheet) throw new Error('シートが見つかりません');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const unpostedFolder = DriveApp.getFolderById(UNPOSTED_FOLDER_ID);

  for (let row = 2; row <= lastRow; row++) {

    const blogDocIdCell = sheet.getRange(row, COL_BLOG_DOC_ID).getValue();
    const scriptDocId = sheet.getRange(row, COL_SCRIPT_DOC_ID).getValue();
    const slideId = sheet.getRange(row, COL_SLIDE_ID).getValue();

    // ===== 既にブログ記事がある行はスキップ =====
    if (blogDocIdCell) {
      Logger.log(`行${row}: 既にブログ記事あり → スキップ`);
      continue;
    }

    // ===== 原稿IDがない場合 =====
    if (!scriptDocId) {
      Logger.log(`行${row}: 原稿IDが見つかりません`);
      continue;
    }

    try {
      Logger.log(`行${row}: ブログ記事生成開始`);

      // 原稿本文取得
      const scriptDoc = DocumentApp.openById(scriptDocId);
      const scriptText = scriptDoc.getBody().getText();

      // プロンプト生成
      const prompt = createPrompt(scriptText, slideId);

      // Gemini API 呼び出し
      const blogArticle = callGeminiAPI(prompt);

      // ===== ブログ記事ドキュメント作成 =====
      const blogFileName = `noteブログ記事_${row}_${Date.now()}`;
      const blogDoc = DocumentApp.create(blogFileName);
      blogDoc.getBody().setText(blogArticle);

      const blogDocId = blogDoc.getId();

      // ===== G列にブログ記事IDを書き込み =====
      sheet.getRange(row, COL_BLOG_DOC_ID).setValue(blogDocId);

      // ===== 未投稿記事フォルダへ移動（IDは変わらない）=====
      const blogFile = DriveApp.getFileById(blogDocId);
      blogFile.moveTo(unpostedFolder);

      Logger.log(`行${row}: 完了`);
      Logger.log(`URL: ${blogDoc.getUrl()}`);

    } catch (e) {
      Logger.log(`行${row}: エラー - ${e.message}`);
    }

    Utilities.sleep(1000); // レート制限対策
  }

  Logger.log('全処理完了');
}

/**
 * プロンプト生成
 */
function createPrompt(scriptText, slideId) {
  let prompt = '';
  prompt += 'あなたはAI編集者です。\n\n';

  if (slideId) {
    prompt += `スライドURL: https://docs.google.com/presentation/d/${slideId}\n\n`;
  }

  prompt += '以下の原稿を元に、note向けブログ記事を作成してください。\n';
  prompt += '"""\n';
  prompt += scriptText;
  prompt += '\n"""\n';

  return prompt;
}

/**
 * Gemini API 呼び出し
 */
function callGeminiAPI(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8000
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(response.getContentText());
  }

  const data = JSON.parse(response.getContentText());
  return data.candidates[0].content.parts[0].text;
}
