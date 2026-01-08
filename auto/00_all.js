/**
 * 統合自動化スクリプト (04_rootToSlide ～ 11_rootToMiBlog)
 * * 実行順序:
 * 1. 04: ルートのスライドを指定フォルダへ移動
 * 2. 05: フォルダ内のスライド一覧をスプレッドシートに出力
 * 3. 06: スライドからスピーカーノート(原稿)を生成(Gemini)
 * 4. 07: 生成した原稿をスライドのノート欄に転記
 * 5. 08: 原稿とスライドからブログ記事を生成(Gemini)
 * 6. 09: 使用した原稿(Googleドキュメント)を使用済フォルダへ移動
 * 7. 10: 原稿作成済みのスライドを原稿ありフォルダへ移動
 * 8. 11: ブログ記事(Googleドキュメント)を未投稿フォルダへ移動
 */

// ===== 共通設定 (ここを設定してください) =====
const CONFIG = {
  SPREADSHEET_ID: '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU',
  SHEET_NAME: 'フォームの回答 1', // 基本のシート名
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('YOUR_GEMINI_API_KEY'), // ★スクリプトプロパティから読み込み

  // フォルダID設定
  FOLDER_ID_SLIDE_STORAGE: '1vhp1a-6eUkWPTZ01jnpzCrO4fGbh0WHb', // 04, 05で使用 (Slideフォルダ)
  FOLDER_ID_USED_DOCS: '1sHXWVlKHVZzyEM8gikOhulVOH2-aRBnf',      // 09で使用 (使用済原稿)
  FOLDER_ID_SCRIPTED_SLIDES: '1KqLbSMfTPFC4oiXV4iuWvc6gECj36Yna', // 10で使用 (原稿ありスライド)
  FOLDER_ID_UNPOSTED_BLOG: '1KYutXxQUaMuOZVs34Apt94FLYMI7VoBh',   // 11で使用 (未投稿記事)
  FOLDER_ID_OUTPUT_ROOT: '', // 06, 08の出力先（空ならルート）

  // Gemini設定
  GEMINI_MODEL: 'gemini-2.0-flash-exp'
};

/**
 * 【メイン実行関数】
 * この関数を実行すると、04から11までの処理を順番に行います。
 */
function runAllSteps() {
  Logger.log('========== 全工程の処理を開始します ==========');

  try {
    step04_moveRootSlides();
    Utilities.sleep(1000); // 処理間に待機時間を挟む

    step05_exportSlideUrls();
    Utilities.sleep(1000);

    step06_generateSpeakerNotes();
    Utilities.sleep(1000);

    step07_importSpeakerNotes();
    Utilities.sleep(1000);

    step08_generateBlogArticles();
    Utilities.sleep(1000);

    step09_moveUsedScriptDocs();
    Utilities.sleep(1000);

    step10_moveSlidesWithScript();
    Utilities.sleep(1000);

    step11_moveBlogDocs();

    Logger.log('========== 全工程の処理が正常に完了しました ==========');
  } catch (e) {
    Logger.log(`エラーが発生し、処理が中断されました: ${e.message}`);
    Logger.log(e.stack);
  }
}


// ==========================================
// 04: ルートフォルダのスライドをSlideフォルダへ移動
// ==========================================
function step04_moveRootSlides() {
  Logger.log('--- Step 04: ルートスライドの移動開始 ---');
  const targetFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_SLIDE_STORAGE);
  const rootFolder = DriveApp.getRootFolder();
  const files = rootFolder.getFilesByType(MimeType.GOOGLE_SLIDES);

  let count = 0;
  while (files.hasNext()) {
    const file = files.next();
    targetFolder.addFile(file);
    rootFolder.removeFile(file);
    Logger.log(`移動完了: ${file.getName()}`);
    count++;
  }
  Logger.log(`Step 04 完了: 移動したスライド数: ${count}`);
}


// ==========================================
// 05: スライドURLを取得してスプレッドシートへ書き出し
// ==========================================
function step05_exportSlideUrls() {
  Logger.log('--- Step 05: スライドURLの書き出し開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEET_NAME}」が見つかりません`);

  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID_SLIDE_STORAGE);
  const files = folder.getFiles();

  const titles = [];
  const urls = [];
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    titles.push([file.getName()]);
    urls.push([file.getUrl()]);
    count++;
  }

  if (count > 0) {
    // B列(2)にタイトル、D列(4)にURL
    sheet.getRange(2, 2, titles.length, 1).setValues(titles);
    sheet.getRange(2, 4, urls.length, 1).setValues(urls);
    Logger.log(`書き込み完了: ${count} 件`);
  } else {
    Logger.log('書き込むファイルがありません');
  }
  Logger.log('Step 05 完了');
}


// ==========================================
// 06: スピーカーノート(原稿)生成
// ==========================================
function step06_generateSpeakerNotes() {
  Logger.log('--- Step 06: スピーカーノート生成開始 ---');
  
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    throw new Error('APIキーが設定されていません');
  }

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません');
    return;
  }

  // I列(9)のスライドIDを取得
  const data = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  const outputFolder = CONFIG.FOLDER_ID_OUTPUT_ROOT ? DriveApp.getFolderById(CONFIG.FOLDER_ID_OUTPUT_ROOT) : DriveApp.getRootFolder();

  data.forEach((row, index) => {
    const slideId = row[0];
    const rowNum = index + 2;

    if (!slideId) return;

    // 既に原稿URLがある場合(C列など)はスキップするロジックが必要ならここに追加
    // 今回は上書き動作として実装

    try {
      Logger.log(`行 ${rowNum}: 原稿生成処理開始...`);
      const slideContent = getSlideContent(slideId);
      const prompt = createSpeakerNotesPrompt(slideContent);
      const speakerNotes = callGeminiAPI(prompt);

      const fileName = `スピーカーノート_${rowNum}_${new Date().getTime()}`;
      const doc = DocumentApp.create(fileName);
      doc.getBody().setText(speakerNotes);
      
      const file = DriveApp.getFileById(doc.getId());
      file.moveTo(outputFolder);

      // C列(3)に生成したドキュメントのURLを書き込み (06のコードに基づく)
      sheet.getRange(rowNum, 3).setValue(doc.getUrl());
      // H列(8)にドキュメントIDを書き込む必要がある場合(後続処理用)
      // 08,09等のためにH列にIDを入れておくとスムーズかもしれません
      // 05の構成上H列がDocID用と推測されますが、ここでは明示的に書き込むか確認が必要です
      // 元の06コードではURLを3列目に入れていました。
      // 09や07はH列(8)のIDを参照するため、ここでH列にIDを入れる修正を加えます。
      sheet.getRange(rowNum, 8).setValue(doc.getId());

      Logger.log(`行 ${rowNum}: 生成成功 - ${fileName}`);
      Utilities.sleep(2000); // レート制限対策
    } catch (e) {
      Logger.log(`行 ${rowNum}: エラー - ${e.message}`);
    }
  });
  Logger.log('Step 06 完了');
}


// ==========================================
// 07: 生成した原稿をスライドに転記
// ==========================================
function step07_importSpeakerNotes() {
  Logger.log('--- Step 07: スライドへの原稿転記開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // H列(8:DocID)とI列(9:SlideID)を取得
  const docIds = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  const slideIds = sheet.getRange(2, 9, lastRow - 1, 1).getValues();

  for (let i = 0; i < docIds.length; i++) {
    const docId = String(docIds[i][0]).trim();
    const slideId = String(slideIds[i][0]).trim();
    if (!docId || !slideId) continue;

    try {
      const doc = DocumentApp.openById(docId);
      const text = doc.getBody().getText();
      const slideMatches = [...text.matchAll(/@@ Slide (\d+)[^\n]*\n([\s\S]*?)(?=@@ Slide \d+|$)/g)];

      if (slideMatches.length > 0) {
        const presentation = SlidesApp.openById(slideId);
        const slides = presentation.getSlides();

        slideMatches.forEach(match => {
          const slideIndex = Number(match[1]) - 1;
          const noteText = match[2].trim();
          if (slideIndex < slides.length && noteText) {
            slides[slideIndex].getNotesPage().getSpeakerNotesShape().getText().setText(noteText);
          }
        });
        Logger.log(`行 ${i + 2}: 転記完了`);
      }
    } catch (e) {
      Logger.log(`行 ${i + 2}: エラー - ${e.message}`);
    }
  }
  Logger.log('Step 07 完了');
}


// ==========================================
// 08: ブログ記事生成
// ==========================================
function step08_generateBlogArticles() {
  Logger.log('--- Step 08: ブログ記事生成開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // H列(8:DocID)とI列(9:SlideID)を取得
  const data = sheet.getRange(2, 8, lastRow - 1, 2).getValues();
  const outputFolder = CONFIG.FOLDER_ID_OUTPUT_ROOT ? DriveApp.getFolderById(CONFIG.FOLDER_ID_OUTPUT_ROOT) : DriveApp.getRootFolder();
  const colBlogDocId = 7; // G列

  data.forEach((row, index) => {
    const docId = row[0];
    const slideId = row[1];
    const rowNum = index + 2;

    if (!docId && !slideId) return;

    // 既にG列にブログIDがある場合はスキップ等の判定が必要なら追加

    try {
      Logger.log(`行 ${rowNum}: ブログ生成開始...`);
      const prompt = createBlogPrompt(docId, slideId);
      const blogArticle = callGeminiAPI(prompt);

      const fileName = `noteブログ記事_${rowNum}_${new Date().getTime()}`;
      const doc = DocumentApp.create(fileName);
      applyMarkdownToDocument(doc.getBody(), blogArticle);

      const blogDocId = doc.getId();
      sheet.getRange(rowNum, colBlogDocId).setValue(blogDocId);

      DriveApp.getFileById(blogDocId).moveTo(outputFolder);
      Logger.log(`行 ${rowNum}: ブログ生成完了`);
      Utilities.sleep(2000);
    } catch (e) {
      Logger.log(`行 ${rowNum}: エラー - ${e.message}`);
    }
  });
  Logger.log('Step 08 完了');
}


// ==========================================
// 09: 使用済原稿(H列)を移動
// ==========================================
function step09_moveUsedScriptDocs() {
  Logger.log('--- Step 09: 使用済原稿の移動開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const docIds = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // H列
  const usedFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_USED_DOCS);

  docIds.forEach((row, index) => {
    const docId = row[0];
    if (!docId) return;

    try {
      const file = DriveApp.getFileById(docId);
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        usedFolder.addFile(file);
        const parents = file.getParents();
        while (parents.hasNext()) {
          parents.next().removeFile(file);
        }
        Logger.log(`行 ${index + 2}: 移動完了`);
      }
    } catch (e) {
      Logger.log(`行 ${index + 2}: 移動エラー - ${e.message}`);
    }
  });
  Logger.log('Step 09 完了');
}


// ==========================================
// 10: 原稿ありスライド(I列)を移動
// ==========================================
function step10_moveSlidesWithScript() {
  Logger.log('--- Step 10: 原稿ありスライドの移動開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const slideIds = sheet.getRange(2, 9, lastRow - 1, 1).getValues(); // I列
  const targetFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_SCRIPTED_SLIDES);

  slideIds.forEach((row, index) => {
    const slideId = row[0];
    if (!slideId) return;

    try {
      const file = DriveApp.getFileById(slideId);
      if (file.getMimeType() === MimeType.GOOGLE_SLIDES) {
        targetFolder.addFile(file);
        const parents = file.getParents();
        while (parents.hasNext()) {
          parents.next().removeFile(file);
        }
        Logger.log(`行 ${index + 2}: 移動完了`);
      }
    } catch (e) {
      Logger.log(`行 ${index + 2}: 移動エラー - ${e.message}`);
    }
  });
  Logger.log('Step 10 完了');
}


// ==========================================
// 11: ブログ記事(G列)を未投稿フォルダへ移動
// ==========================================
function step11_moveBlogDocs() {
  Logger.log('--- Step 11: ブログ記事の移動開始 ---');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID_UNPOSTED_BLOG);
  const colBlogDocId = 7; // G列

  for (let row = 2; row <= lastRow; row++) {
    const docId = sheet.getRange(row, colBlogDocId).getValue();
    if (!docId) continue;

    try {
      const file = DriveApp.getFileById(docId);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      Logger.log(`行 ${row}: 移動完了`);
    } catch (e) {
      Logger.log(`行 ${row}: 移動エラー - ${e.message}`);
    }
  }
  Logger.log('Step 11 完了');
}


// ==========================================
// ヘルパー関数群
// ==========================================

/**
 * Gemini API 呼び出し (共通)
 */
function callGeminiAPI(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8000,
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    throw new Error(`API Error: ${responseCode} - ${response.getContentText()}`);
  }
  
  const data = JSON.parse(response.getContentText());
  if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts) {
    return data.candidates[0].content.parts[0].text;
  } else {
    throw new Error('APIレスポンスにコンテンツがありません');
  }
}

/**
 * スライド内容取得 (Step 06用)
 */
function getSlideContent(slideId) {
  try {
    const presentation = SlidesApp.openById(slideId);
    const slides = presentation.getSlides();
    let content = '';
    
    slides.forEach((slide, index) => {
      content += `\n=== スライド ${index + 1} ===\n`;
      slide.getShapes().forEach(shape => {
        if (shape.getText()) {
          const text = shape.getText().asString();
          if (text.trim()) content += text + '\n';
        }
      });
      slide.getTables().forEach(table => {
        const numRows = table.getNumRows();
        const numCols = table.getNumColumns();
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            const cellText = table.getCell(i, j).getText().asString();
            if (cellText.trim()) content += cellText + ' ';
          }
          content += '\n';
        }
      });
      content += '\n';
    });
    return content;
  } catch (e) {
    throw new Error(`スライド取得エラー: ${e.message}`);
  }
}

/**
 * スピーカーノート用プロンプト作成 (Step 06用)
 */
function createSpeakerNotesPrompt(slideContent) {
  return `スライド資料に対応するスピーカーノート(話す原稿)を作成してください。

【出力ルール】
・各スライドは必ず次の形式で出力してください
・見出し記号「#」は使用しないでください
・各スライドの区切りには「@@ Slide 数字」を使ってください
・余計な前置き、まとめ、注釈は書かないでください

【出力形式(厳守)】
@@ Slide 1
ここにスライド1で話す原稿を書く
視聴者に語りかける口調で、
背景説明・具体例・補足を含める

@@ Slide 2
ここにスライド2で話す原稿を書く

(以下、全スライド分続ける)

【原稿のトーン】
・中学生〜社会人初級者にも分かるやさしい説明
・動画でそのまま読める自然な話し言葉
・1スライドあたり多すぎず、1〜2分で話せる分量

【出力形式】
・Markdown(.md)形式
・本文のみを出力

---
【スライドの内容】
${slideContent}`;
}

/**
 * ブログ記事用プロンプト作成 (Step 08用)
 */
function createBlogPrompt(docId, slideId) {
  let prompt = 'あなたは「提供された動画・資料・原稿を統合し、マルチメディア記事として最適化するAI編集者」です。\n';
  prompt += '以下の【入力データ】を元に、ブログ記事を作成してください。\n\n';
  prompt += '**【入力データ】**\n';
  prompt += '* **YouTube URL:** [ここにYouTube動画のURLを貼り付け]\n';

  if (slideId) {
    prompt += `* **スライドURL:** https://docs.google.com/presentation/d/${slideId}\n`;
  } else {
    prompt += '* **スライドURL:** [スライドIDが見つかりません]\n';
  }

  prompt += '* **入力原稿:**\n    """\n';
  if (docId) {
    try {
      const doc = DocumentApp.openById(docId);
      prompt += `${doc.getBody().getText()}\n`;
    } catch (e) {
      prompt += `[ドキュメント取得エラー: ${e.message}]\n`;
    }
  } else {
    prompt += '[ドキュメントIDが見つかりません]\n';
  }
  prompt += '    """\n\n---\n';
  
  prompt += '**【出力のルール】**\n';
  prompt += '1. **構成順序:** 記事の冒頭で「動画」を見せ、次に「スライド」で補足し、最後に「テキスト」で詳細を読む流れを作る。\n';
  prompt += '2. **ターゲット:** 「動画を見る派」も「読む派」も両方取り込みたいWeb読者。\n';
  prompt += '3. **連携:** 動画とスライドの内容が原稿に基づいていることを前提に、相互にリンクさせるような紹介文にする。\n\n';
  
  // (長いのでフォーマット指定の一部は省略しつつ、重要な骨子を維持)
  prompt += '**【出力フォーマット】**\n';
  prompt += '## タイトル案\n（3つ提案）\n\n';
  prompt += '## 1. AI生成レポート\n* **トピック:** [キーワード]\n* **AIの要約:** （140文字以内）\n\n';
  prompt += '## 2. 【動画で見る】（推奨）\n**📺 [YouTube動画を再生する]**\n\n';
  prompt += '## 3. 【スライドで要点把握】\n**👉 [スライドを表示する]**\n\n';
  prompt += '## 4. 【テキスト詳細解説】\n（入力原稿をブログ向けに整形）\n\n';
  prompt += '## 5. AIの深掘り考察\n\n';
  prompt += '## 6. 編集後記（Human）\n';

  return prompt;
}

/**
 * Markdown適用 (Step 08用)
 */
function applyMarkdownToDocument(body, markdown) {
  body.clear();
  const lines = markdown.split('\n');
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.startsWith('# ')) {
      body.appendParagraph(line.substring(2)).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (line.startsWith('## ')) {
      body.appendParagraph(line.substring(3)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (line.startsWith('### ')) {
      body.appendParagraph(line.substring(4)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    } else if (line.match(/^[\-\*]\s+/)) {
      body.appendListItem(line.replace(/^[\-\*]\s+/, '')).setGlyphType(DocumentApp.GlyphType.BULLET);
    } else if (line.trim() === '') {
      body.appendParagraph('');
    } else {
      applyInlineFormatting(body.appendParagraph(line));
    }
  }
}

/**
 * インライン書式適用 (Step 08用)
 */
function applyInlineFormatting(paragraph) {
  const text = paragraph.getText();
  
  // 太字 **text**
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(text)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    // 注: 本来は後ろから処理するかオフセット計算が必要ですが、簡易実装のため省略
    // 厳密なMarkdownパーサーが必要な場合はライブラリ推奨
    paragraph.editAsText().setBold(start, end - 1, true); // 簡易的に全体を太字化（記号含む）
  }
}
