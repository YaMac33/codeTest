/**
 * 統合自動化スクリプト改善版
 * 主な改善点:
 * - 設定の検証機能追加
 * - エラーハンドリング強化
 * - 列定義の明示化
 * - 処理状態管理
 * - ログの構造化
 */

// ===== 列定義 (スプレッドシートの構造を明示) =====
const COLUMNS = {
  TIMESTAMP: 1,        // A列: タイムスタンプ
  SLIDE_TITLE: 2,      // B列: スライドタイトル
  SCRIPT_DOC_URL: 3,   // C列: 原稿ドキュメントURL
  SLIDE_URL: 4,        // D列: スライドURL
  STATUS: 5,           // E列: 処理状態
  ERROR_LOG: 6,        // F列: エラーログ
  BLOG_DOC_ID: 7,      // G列: ブログドキュメントID
  SCRIPT_DOC_ID: 8,    // H列: 原稿ドキュメントID
  SLIDE_ID: 9          // I列: スライドID
};

// ===== 処理状態の定義 =====
const STATUS = {
  PENDING: '未処理',
  PROCESSING: '処理中',
  COMPLETED: '完了',
  ERROR: 'エラー',
  SKIPPED: 'スキップ'
};

// ===== 設定 =====
const CONFIG = {
  SPREADSHEET_ID: '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU',
  SHEET_NAME: 'フォームの回答 1',
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('YOUR_GEMINI_API_KEY'),

  // フォルダID
  FOLDER_ID_SLIDE_STORAGE: '1vhp1a-6eUkWPTZ01jnpzCrO4fGbh0WHb',
  FOLDER_ID_USED_DOCS: '1sHXWVlKHVZzyEM8gikOhulVOH2-aRBnf',
  FOLDER_ID_SCRIPTED_SLIDES: '1KqLbSMfTPFC4oiXV4iuWvc6gECj36Yna',
  FOLDER_ID_UNPOSTED_BLOG: '1KYutXxQUaMuOZVs34Apt94FLYMI7VoBh',
  FOLDER_ID_OUTPUT_ROOT: '',

  // Gemini設定
  GEMINI_MODEL: 'gemini-2.0-flash-exp',
  
  // 処理設定
  RETRY_COUNT: 3,           // リトライ回数
  RETRY_DELAY_MS: 2000,     // リトライ間隔
  API_CALL_DELAY_MS: 2000,  // API呼び出し間隔
  BATCH_SIZE: 10            // バッチ処理サイズ
};

// ===== 設定検証 =====
class ConfigValidator {
  static validate() {
    const errors = [];
    
    // APIキーチェック
    if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY') {
      errors.push('Gemini APIキーが設定されていません');
    }
    
    // スプレッドシートチェック
    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {
      errors.push(`スプレッドシートにアクセスできません: ${CONFIG.SPREADSHEET_ID}`);
    }
    
    // フォルダチェック
    const folderIds = [
      'FOLDER_ID_SLIDE_STORAGE',
      'FOLDER_ID_USED_DOCS',
      'FOLDER_ID_SCRIPTED_SLIDES',
      'FOLDER_ID_UNPOSTED_BLOG'
    ];
    
    folderIds.forEach(key => {
      if (CONFIG[key]) {
        try {
          DriveApp.getFolderById(CONFIG[key]);
        } catch (e) {
          errors.push(`フォルダにアクセスできません (${key}): ${CONFIG[key]}`);
        }
      }
    });
    
    if (errors.length > 0) {
      throw new Error(`設定エラー:\n${errors.join('\n')}`);
    }
    
    Logger.log('設定検証: OK');
    return true;
  }
}

// ===== 実行結果管理 =====
class ExecutionResult {
  constructor(stepName) {
    this.stepName = stepName;
    this.successCount = 0;
    this.errorCount = 0;
    this.skippedCount = 0;
    this.errors = [];
    this.startTime = new Date();
  }
  
  addSuccess() {
    this.successCount++;
  }
  
  addError(rowNum, message) {
    this.errorCount++;
    this.errors.push({ row: rowNum, message });
  }
  
  addSkipped() {
    this.skippedCount++;
  }
  
  getSummary() {
    const duration = ((new Date() - this.startTime) / 1000).toFixed(1);
    return {
      step: this.stepName,
      duration: `${duration}秒`,
      success: this.successCount,
      error: this.errorCount,
      skipped: this.skippedCount,
      total: this.successCount + this.errorCount + this.skippedCount
    };
  }
  
  logSummary() {
    const summary = this.getSummary();
    Logger.log(`\n=== ${summary.step} 完了 ===`);
    Logger.log(`処理時間: ${summary.duration}`);
    Logger.log(`成功: ${summary.success}, エラー: ${summary.error}, スキップ: ${summary.skipped}, 合計: ${summary.total}`);
    
    if (this.errors.length > 0) {
      Logger.log('\n【エラー詳細】');
      this.errors.forEach(err => {
        Logger.log(`  行${err.row}: ${err.message}`);
      });
    }
  }
}

// ===== スプレッドシート操作ヘルパー =====
class SheetHelper {
  constructor(spreadsheetId, sheetName) {
    this.sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!this.sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }
  }
  
  getLastRow() {
    return this.sheet.getLastRow();
  }
  
  getValue(row, col) {
    return this.sheet.getRange(row, col).getValue();
  }
  
  setValue(row, col, value) {
    this.sheet.getRange(row, col).setValue(value);
  }
  
  getValues(startRow, col, numRows) {
    return this.sheet.getRange(startRow, col, numRows, 1).getValues();
  }
  
  setValues(startRow, col, values) {
    this.sheet.getRange(startRow, col, values.length, 1).setValues(values);
  }
  
  updateStatus(row, status, errorMessage = '') {
    this.setValue(row, COLUMNS.STATUS, status);
    if (errorMessage) {
      this.setValue(row, COLUMNS.ERROR_LOG, errorMessage);
    }
  }
  
  getDataRange(startRow = 2) {
    const lastRow = this.getLastRow();
    if (lastRow < startRow) return [];
    
    const numRows = lastRow - startRow + 1;
    return this.sheet.getRange(startRow, 1, numRows, 9).getValues();
  }
}

// ===== リトライ機能付きAPI呼び出し =====
class GeminiClient {
  static call(prompt, retryCount = CONFIG.RETRY_COUNT) {
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
    
    for (let attempt = 1; attempt <= retryCount; attempt++) {
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          const data = JSON.parse(response.getContentText());
          if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts) {
            return data.candidates[0].content.parts[0].text;
          }
        }
        
        // リトライ可能なエラーコード
        if ([429, 500, 503].includes(responseCode) && attempt < retryCount) {
          Logger.log(`API呼び出し失敗 (試行${attempt}/${retryCount}): ${responseCode}`);
          Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
          continue;
        }
        
        throw new Error(`API Error: ${responseCode} - ${response.getContentText()}`);
        
      } catch (e) {
        if (attempt === retryCount) {
          throw e;
        }
        Logger.log(`API呼び出しエラー (試行${attempt}/${retryCount}): ${e.message}`);
        Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
      }
    }
  }
}

// ===== ファイル操作ヘルパー =====
class FileHelper {
  static moveFile(fileId, targetFolderId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const targetFolder = DriveApp.getFolderById(targetFolderId);
      
      targetFolder.addFile(file);
      
      // 元のフォルダから削除
      const parents = file.getParents();
      while (parents.hasNext()) {
        const parent = parents.next();
        if (parent.getId() !== targetFolderId) {
          parent.removeFile(file);
        }
      }
      
      return true;
    } catch (e) {
      throw new Error(`ファイル移動エラー: ${e.message}`);
    }
  }
  
  static createDocument(title, content, folderId = null) {
    const doc = DocumentApp.create(title);
    doc.getBody().setText(content);
    
    const file = DriveApp.getFileById(doc.getId());
    
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      file.moveTo(folder);
    }
    
    return {
      id: doc.getId(),
      url: doc.getUrl()
    };
  }
}

// ===== メイン実行関数 =====
function runAllSteps() {
  Logger.log('========== 全工程の処理を開始 ==========');
  
  const allResults = [];
  
  try {
    // 設定検証
    ConfigValidator.validate();
    
    // 各ステップを実行
    const steps = [
      { name: 'Step 04: スライド移動', func: step04_moveRootSlides },
      { name: 'Step 05: URL書き出し', func: step05_exportSlideUrls },
      { name: 'Step 06: 原稿生成', func: step06_generateSpeakerNotes },
      { name: 'Step 07: 原稿転記', func: step07_importSpeakerNotes },
      { name: 'Step 08: ブログ生成', func: step08_generateBlogArticles },
      { name: 'Step 09: 原稿移動', func: step09_moveUsedScriptDocs },
      { name: 'Step 10: スライド移動', func: step10_moveSlidesWithScript },
      { name: 'Step 11: ブログ移動', func: step11_moveBlogDocs }
    ];
    
    steps.forEach(step => {
      Logger.log(`\n--- ${step.name} 開始 ---`);
      const result = step.func();
      if (result) {
        result.logSummary();
        allResults.push(result.getSummary());
      }
      Utilities.sleep(1000);
    });
    
    // 全体サマリー
    Logger.log('\n========== 全工程完了サマリー ==========');
    allResults.forEach(result => {
      Logger.log(`${result.step}: 成功${result.success} / エラー${result.error} / スキップ${result.skipped}`);
    });
    Logger.log('========================================');
    
  } catch (e) {
    Logger.log(`\n!!! 致命的エラー: ${e.message}`);
    Logger.log(e.stack);
    throw e;
  }
}

// ===== Step 06: 原稿生成 (改善版) =====
function step06_generateSpeakerNotes() {
  const result = new ExecutionResult('Step 06: 原稿生成');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const data = sheet.getDataRange();
  const outputFolder = CONFIG.FOLDER_ID_OUTPUT_ROOT 
    ? DriveApp.getFolderById(CONFIG.FOLDER_ID_OUTPUT_ROOT) 
    : DriveApp.getRootFolder();
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const slideId = row[COLUMNS.SLIDE_ID - 1];
    const existingDocId = row[COLUMNS.SCRIPT_DOC_ID - 1];
    
    // スライドIDがない、または既に原稿がある場合はスキップ
    if (!slideId) {
      result.addSkipped();
      return;
    }
    
    if (existingDocId) {
      Logger.log(`行${rowNum}: 既に原稿が存在するためスキップ`);
      result.addSkipped();
      return;
    }
    
    try {
      sheet.updateStatus(rowNum, STATUS.PROCESSING);
      
      // スライド内容取得
      const slideContent = getSlideContent(slideId);
      
      // プロンプト作成とAPI呼び出し
      const prompt = createSpeakerNotesPrompt(slideContent);
      const speakerNotes = GeminiClient.call(prompt);
      
      // ドキュメント作成
      const fileName = `スピーカーノート_${rowNum}_${new Date().getTime()}`;
      const docInfo = FileHelper.createDocument(
        fileName, 
        speakerNotes, 
        outputFolder.getId()
      );
      
      // シート更新
      sheet.setValue(rowNum, COLUMNS.SCRIPT_DOC_URL, docInfo.url);
      sheet.setValue(rowNum, COLUMNS.SCRIPT_DOC_ID, docInfo.id);
      sheet.updateStatus(rowNum, STATUS.COMPLETED);
      
      result.addSuccess();
      Logger.log(`行${rowNum}: 原稿生成完了`);
      
      Utilities.sleep(CONFIG.API_CALL_DELAY_MS);
      
    } catch (e) {
      const errorMsg = `原稿生成エラー: ${e.message}`;
      sheet.updateStatus(rowNum, STATUS.ERROR, errorMsg);
      result.addError(rowNum, errorMsg);
      Logger.log(`行${rowNum}: ${errorMsg}`);
    }
  });
  
  return result;
}

// ===== 以下、他のステップも同様のパターンで改善 =====
// (スペースの都合上、主要な改善パターンを示しました)

// 既存の補助関数群 (getSlideContent, createSpeakerNotesPrompt等) は
// そのまま使用可能ですが、必要に応じてエラーハンドリングを強化します

function getSlideContent(slideId) {
  try {
    const presentation = SlidesApp.openById(slideId);
    const slides = presentation.getSlides();
    let content = '';
    
    slides.forEach((slide, index) => {
      content += `\n=== スライド ${index + 1} ===\n`;
      
      // テキストシェイプ処理
      slide.getShapes().forEach(shape => {
        try {
          if (shape.getText()) {
            const text = shape.getText().asString();
            if (text.trim()) content += text + '\n';
          }
        } catch (e) {
          Logger.log(`スライド${index + 1}のシェイプ読み込みエラー: ${e.message}`);
        }
      });
      
      // テーブル処理
      slide.getTables().forEach(table => {
        try {
          const numRows = table.getNumRows();
          const numCols = table.getNumColumns();
          for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
              const cellText = table.getCell(i, j).getText().asString();
              if (cellText.trim()) content += cellText + ' ';
            }
            content += '\n';
          }
        } catch (e) {
          Logger.log(`スライド${index + 1}のテーブル読み込みエラー: ${e.message}`);
        }
      });
      
      content += '\n';
    });
    
    return content;
  } catch (e) {
    throw new Error(`スライド取得エラー (ID: ${slideId}): ${e.message}`);
  }
}

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

// ===== Step 04: ルートスライド移動 =====
function step04_moveRootSlides() {
  const result = new ExecutionResult('Step 04: スライド移動');
  
  try {
    const targetFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_SLIDE_STORAGE);
    const rootFolder = DriveApp.getRootFolder();
    const files = rootFolder.getFilesByType(MimeType.GOOGLE_SLIDES);
    
    while (files.hasNext()) {
      const file = files.next();
      try {
        targetFolder.addFile(file);
        rootFolder.removeFile(file);
        result.addSuccess();
        Logger.log(`移動完了: ${file.getName()}`);
      } catch (e) {
        result.addError(file.getName(), e.message);
      }
    }
  } catch (e) {
    Logger.log(`Step 04 エラー: ${e.message}`);
    throw e;
  }
  
  return result;
}

// ===== Step 05: スライドURL書き出し =====
function step05_exportSlideUrls() {
  const result = new ExecutionResult('Step 05: URL書き出し');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID_SLIDE_STORAGE);
    const files = folder.getFiles();
    
    const titles = [];
    const urls = [];
    
    while (files.hasNext()) {
      const file = files.next();
      titles.push([file.getName()]);
      urls.push([file.getUrl()]);
      result.addSuccess();
    }
    
    if (titles.length > 0) {
      sheet.setValues(2, COLUMNS.SLIDE_TITLE, titles);
      sheet.setValues(2, COLUMNS.SLIDE_URL, urls);
      Logger.log(`書き込み完了: ${titles.length} 件`);
    } else {
      Logger.log('書き込むファイルがありません');
    }
  } catch (e) {
    Logger.log(`Step 05 エラー: ${e.message}`);
    throw e;
  }
  
  return result;
}

// ===== Step 07: 原稿転記 =====
function step07_importSpeakerNotes() {
  const result = new ExecutionResult('Step 07: 原稿転記');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const data = sheet.getDataRange();
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const docId = String(row[COLUMNS.SCRIPT_DOC_ID - 1]).trim();
    const slideId = String(row[COLUMNS.SLIDE_ID - 1]).trim();
    
    if (!docId || !slideId) {
      result.addSkipped();
      return;
    }
    
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
        
        result.addSuccess();
        Logger.log(`行${rowNum}: 転記完了`);
      } else {
        result.addSkipped();
      }
    } catch (e) {
      result.addError(rowNum, e.message);
      Logger.log(`行${rowNum}: エラー - ${e.message}`);
    }
  });
  
  return result;
}

// ===== Step 08: ブログ記事生成 =====
function step08_generateBlogArticles() {
  const result = new ExecutionResult('Step 08: ブログ生成');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const data = sheet.getDataRange();
  const outputFolder = CONFIG.FOLDER_ID_OUTPUT_ROOT 
    ? DriveApp.getFolderById(CONFIG.FOLDER_ID_OUTPUT_ROOT) 
    : DriveApp.getRootFolder();
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const docId = row[COLUMNS.SCRIPT_DOC_ID - 1];
    const slideId = row[COLUMNS.SLIDE_ID - 1];
    const existingBlogId = row[COLUMNS.BLOG_DOC_ID - 1];
    
    if (!docId && !slideId) {
      result.addSkipped();
      return;
    }
    
    // if (existingBlogId) {
    //  Logger.log(`行${rowNum}: 既にブログが存在するためスキップ`);
    //  result.addSkipped();
    //  return;
    // }
    
    try {
      sheet.updateStatus(rowNum, STATUS.PROCESSING);
      
      const prompt = createBlogPrompt(docId, slideId);
      const blogArticle = GeminiClient.call(prompt);
      
      const fileName = `noteブログ記事_${rowNum}_${new Date().getTime()}`;
      const doc = DocumentApp.create(fileName);
      applyMarkdownToDocument(doc.getBody(), blogArticle);
      
      const blogDocId = doc.getId();
      DriveApp.getFileById(blogDocId).moveTo(outputFolder);
      
      sheet.setValue(rowNum, COLUMNS.BLOG_DOC_ID, blogDocId);
      sheet.updateStatus(rowNum, STATUS.COMPLETED);
      
      result.addSuccess();
      Logger.log(`行${rowNum}: ブログ生成完了`);
      
      Utilities.sleep(CONFIG.API_CALL_DELAY_MS);
      
    } catch (e) {
      const errorMsg = `ブログ生成エラー: ${e.message}`;
      sheet.updateStatus(rowNum, STATUS.ERROR, errorMsg);
      result.addError(rowNum, errorMsg);
    }
  });
  
  return result;
}

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

function applyInlineFormatting(paragraph) {
  const text = paragraph.getText();
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(text)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    paragraph.editAsText().setBold(start, end - 1, true);
  }
}

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

// ===== Step 09: 使用済原稿移動 =====
function step09_moveUsedScriptDocs() {
  const result = new ExecutionResult('Step 09: 原稿移動');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const data = sheet.getDataRange();
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const docId = row[COLUMNS.SCRIPT_DOC_ID - 1];
    
    if (!docId) {
      result.addSkipped();
      return;
    }
    
    try {
      FileHelper.moveFile(docId, CONFIG.FOLDER_ID_USED_DOCS);
      result.addSuccess();
      Logger.log(`行${rowNum}: 原稿移動完了`);
    } catch (e) {
      result.addError(rowNum, `原稿移動エラー: ${e.message}`);
    }
  });
  
  return result;
}

// ===== Step 10: 原稿ありスライド移動 =====
function step10_moveSlidesWithScript() {
  const result = new ExecutionResult('Step 10: スライド移動');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const data = sheet.getDataRange();
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const slideId = row[COLUMNS.SLIDE_ID - 1];
    
    if (!slideId) {
      result.addSkipped();
      return;
    }
    
    try {
      FileHelper.moveFile(slideId, CONFIG.FOLDER_ID_SCRIPTED_SLIDES);
      result.addSuccess();
      Logger.log(`行${rowNum}: スライド移動完了`);
    } catch (e) {
      result.addError(rowNum, `スライド移動エラー: ${e.message}`);
    }
  });
  
  return result;
}

// ===== Step 11: ブログ記事移動 =====
function step11_moveBlogDocs() {
  const result = new ExecutionResult('Step 11: ブログ移動');
  const sheet = new SheetHelper(CONFIG.SPREADSHEET_ID, CONFIG.SHEET_NAME);
  
  const lastRow = sheet.getLastRow();
  
  for (let row = 2; row <= lastRow; row++) {
    const docId = sheet.getValue(row, COLUMNS.BLOG_DOC_ID);
    
    if (!docId) {
      result.addSkipped();
      continue;
    }
    
    try {
      FileHelper.moveFile(docId, CONFIG.FOLDER_ID_UNPOSTED_BLOG);
      result.addSuccess();
      Logger.log(`行${row}: ブログ移動完了`);
    } catch (e) {
      result.addError(row, `ブログ移動エラー: ${e.message}`);
    }
  }
  
  return result;
}
