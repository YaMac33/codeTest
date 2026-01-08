/**
 * スプレッドシートのI列からスライドIDを取得し、
 * Gemini APIでスピーカーノート（話す原稿）を生成してGoogleドキュメントとして保存する
 */

// ★設定項目
const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU'; // スプレッドシートID
const SHEET_NAME = ''; // シート名(空の場合は最初のシート)
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('YOUR_GEMINI_API_KEY');
const OUTPUT_FOLDER_ID = ''; // 出力先フォルダID(空の場合はマイドライブ)

/**
 * メイン関数:スピーカーノートを生成
 */
function generateSpeakerNotes() {
  // APIキーの確認
  if (!GEMINI_API_KEY || GEMINI_API_KEY === '〇〇') {
    Logger.log('エラー: GEMINI_API_KEYを設定してください');
    return;
  }
  
  // スプレッドシートを取得
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = SHEET_NAME ? spreadsheet.getSheetByName(SHEET_NAME) : spreadsheet.getSheets()[0];
  
  if (!sheet) {
    Logger.log(`エラー: シートが見つかりません`);
    return;
  }
  
  // データ範囲を取得(ヘッダー行を除く)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データが見つかりません');
    return;
  }
  
  // I列(9)のデータを取得
  const data = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  
  // 出力先フォルダを指定
  const outputFolder = OUTPUT_FOLDER_ID ? 
    DriveApp.getFolderById(OUTPUT_FOLDER_ID) : 
    DriveApp.getRootFolder();
  
  // 各行を処理
  data.forEach((row, index) => {
    const slideId = row[0]; // I列: GoogleスライドID
    
    // 空行をスキップ
    if (!slideId) {
      Logger.log(`行 ${index + 2}: スキップ(スライドIDなし)`);
      return;
    }
    
    try {
      Logger.log(`行 ${index + 2}: 処理開始...`);
      
      // スライドの内容を取得
      const slideContent = getSlideContent(slideId);
      
      // プロンプトを生成
      const prompt = createSpeakerNotesPrompt(slideContent);
      
      // Gemini APIでスピーカーノートを生成
      const speakerNotes = callGeminiAPI(prompt);
      
      // Googleドキュメントを作成
      const fileName = `スピーカーノート_${index + 2}_${new Date().getTime()}`;
      const doc = DocumentApp.create(fileName);
      const body = doc.getBody();
      
      // 生成された原稿を書き込み
      body.setText(speakerNotes);
      
      // ファイルを指定フォルダに移動
      const file = DriveApp.getFileById(doc.getId());
      file.moveTo(outputFolder);
      
      Logger.log(`行 ${index + 2}: 成功 - ${fileName}`);
      Logger.log(`ドキュメントURL: ${doc.getUrl()}`);
      sheet.getRange(index + 2, 3).setValue(doc.getUrl());
      
    } catch (error) {
      Logger.log(`行 ${index + 2}: エラー - ${error.message}`);
    }
    
    // APIレート制限対策(少し待機)
    Utilities.sleep(2000);
  });
  
  Logger.log('処理完了');
}

/**
 * スライドの内容を取得
 */
function getSlideContent(slideId) {
  try {
    const presentation = SlidesApp.openById(slideId);
    const slides = presentation.getSlides();
    
    let content = '';
    
    slides.forEach((slide, index) => {
      content += `\n=== スライド ${index + 1} ===\n`;
      
      // スライド内のすべての要素を取得
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        if (shape.getText()) {
          const text = shape.getText().asString();
          if (text.trim()) {
            content += text + '\n';
          }
        }
      });
      
      // テーブルの内容も取得
      const tables = slide.getTables();
      tables.forEach(table => {
        const numRows = table.getNumRows();
        const numCols = table.getNumColumns();
        for (let i = 0; i < numRows; i++) {
          for (let j = 0; j < numCols; j++) {
            const cellText = table.getCell(i, j).getText().asString();
            if (cellText.trim()) {
              content += cellText + ' ';
            }
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
 * スピーカーノート生成用プロンプトを作成
 */
function createSpeakerNotesPrompt(slideContent) {
  let prompt = '';
  
  prompt += 'スライド資料に対応するスピーカーノート(話す原稿)を作成してください。\n\n';
  
  prompt += '【出力ルール】\n';
  prompt += '・各スライドは必ず次の形式で出力してください\n';
  prompt += '・見出し記号「#」は使用しないでください\n';
  prompt += '・各スライドの区切りには「@@ Slide 数字」を使ってください\n';
  prompt += '・余計な前置き、まとめ、注釈は書かないでください\n\n';
  
  prompt += '【出力形式(厳守)】\n\n';
  prompt += '@@ Slide 1\n';
  prompt += 'ここにスライド1で話す原稿を書く\n';
  prompt += '視聴者に語りかける口調で、\n';
  prompt += '背景説明・具体例・補足を含める\n\n';
  prompt += '@@ Slide 2\n';
  prompt += 'ここにスライド2で話す原稿を書く\n\n';
  prompt += '(以下、全スライド分続ける)\n\n';
  
  prompt += '【原稿のトーン】\n';
  prompt += '・中学生〜社会人初級者にも分かるやさしい説明\n';
  prompt += '・動画でそのまま読める自然な話し言葉\n';
  prompt += '・1スライドあたり多すぎず、1〜2分で話せる分量\n\n';
  
  prompt += '【出力形式】\n';
  prompt += '・Markdown(.md)形式\n';
  prompt += '・本文のみを出力\n\n';
  
  prompt += '---\n\n';
  prompt += '【スライドの内容】\n';
  prompt += slideContent;
  
  return prompt;
}

/**
 * Gemini APIを呼び出してスピーカーノートを生成
 */
function callGeminiAPI(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${GEMINI_API_KEY}`;
  
  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
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
  
  // レスポンスからテキストを抽出
  if (data.candidates && data.candidates.length > 0 && 
      data.candidates[0].content && 
      data.candidates[0].content.parts && 
      data.candidates[0].content.parts.length > 0) {
    return data.candidates[0].content.parts[0].text;
  } else {
    throw new Error('APIレスポンスにコンテンツがありません');
  }
}

/**
 * 利用可能なシート名を表示
 */
function listSheetNames() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('利用可能なシート名:');
  spreadsheet.getSheets().forEach(s => Logger.log(`  - ${s.getName()}`));
}
