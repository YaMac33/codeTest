/**
 * スプレッドシートからドキュメントIDとスライドIDを取得し、
 * Gemini APIでnoteブログ記事を生成してGoogleドキュメントとして保存する
 */

// ★設定項目
const SPREADSHEET_ID = '1_w7tG6QF2iQ4hRMRCAIsTuXP7Tg-Rus4X2erGtf2VmU'; // スプレッドシートID
const SHEET_NAME = ''; // シート名(空の場合は最初のシート)
const GEMINI_API_KEY = '○○'; // Gemini APIキー
const OUTPUT_FOLDER_ID = ''; // 出力先フォルダID(空の場合はマイドライブ)

/**
 * メイン関数:noteブログ記事を生成
 */
function generateNoteBlogArticles() {
  // APIキーの確認
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
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
  
  // H列(8)とI列(9)のデータを取得
  const data = sheet.getRange(2, 8, lastRow - 1, 2).getValues();
  
  // 出力先フォルダを指定
  const outputFolder = OUTPUT_FOLDER_ID ? 
    DriveApp.getFolderById(OUTPUT_FOLDER_ID) : 
    DriveApp.getRootFolder();
  
  // 各行を処理
  data.forEach((row, index) => {
    const docId = row[0]; // H列: GoogleドキュメントID
    const slideId = row[1]; // I列: GoogleスライドID
    
    // 空行をスキップ
    if (!docId && !slideId) {
      Logger.log(`行 ${index + 2}: スキップ(空行)`);
      return;
    }
    
    try {
      Logger.log(`行 ${index + 2}: 処理開始...`);
      
      // プロンプトを生成
      const prompt = createPrompt(docId, slideId);
      
      // Gemini APIでブログ記事を生成
      const blogArticle = callGeminiAPI(prompt);
      
      // Googleドキュメントを作成
      const fileName = `noteブログ記事_${index + 2}_${new Date().getTime()}`;
      const doc = DocumentApp.create(fileName);
      const body = doc.getBody();
      
      // マークダウンを適用してドキュメントに書き込み
      applyMarkdownToDocument(body, blogArticle);
      
      // ファイルを指定フォルダに移動
      const file = DriveApp.getFileById(doc.getId());
      file.moveTo(outputFolder);
      
      Logger.log(`行 ${index + 2}: 成功 - ${fileName}`);
      Logger.log(`ドキュメントURL: ${doc.getUrl()}`);
      
    } catch (error) {
      Logger.log(`行 ${index + 2}: エラー - ${error.message}`);
    }
    
    // APIレート制限対策(少し待機)
    Utilities.sleep(1000);
  });
  
  Logger.log('処理完了');
}

/**
 * プロンプトを生成（記事作成プロンプト：新仕様）
 */
function createPrompt(docId, slideId) {
  let prompt = '';
  prompt += 'あなたは「提供された動画・資料・原稿を統合し、マルチメディア記事として最適化するAI編集者」です。\n';
  prompt += '以下の【入力データ】を元に、ブログ記事を作成してください。\n\n';

  prompt += '**【入力データ】**\n';
  prompt += '* **YouTube URL:** [ここにYouTube動画のURLを貼り付け]\n';

  // スライドURL
  if (slideId) {
    const slideUrl = `https://docs.google.com/presentation/d/${slideId}`;
    prompt += `* **スライドURL:** ${slideUrl}\n`;
  } else {
    prompt += '* **スライドURL:** [スライドIDが見つかりません]\n';
  }

  prompt += '* **入力原稿:**\n';
  prompt += '    """\n';

  // ドキュメント本文
  if (docId) {
    try {
      const doc = DocumentApp.openById(docId);
      const docContent = doc.getBody().getText();
      prompt += `${docContent}\n`;
    } catch (e) {
      const docUrl = `https://docs.google.com/document/d/${docId}`;
      prompt += `${docUrl}\n`;
      prompt += `[ドキュメント取得エラー: ${e.message}]\n`;
    }
  } else {
    prompt += '[ドキュメントIDが見つかりません]\n';
  }

  prompt += '    """\n\n';

  prompt += '---\n\n';

  prompt += '**【出力のルール】**\n';
  prompt += '1. **構成順序:** 記事の冒頭で「動画」を見せ、次に「スライド」で補足し、最後に「テキスト」で詳細を読む流れを作る。\n';
  prompt += '2. **ターゲット:** 「動画を見る派」も「読む派」も両方取り込みたいWeb読者。\n';
  prompt += '3. **連携:** 動画とスライドの内容が原稿に基づいていることを前提に、相互にリンクさせるような紹介文にする。\n\n';

  prompt += '---\n\n';

  prompt += '**【出力フォーマット】**\n\n';

  prompt += '## タイトル案\n';
  prompt += '（動画やスライドの内容も含め、クリックしたくなるタイトルを3つ提案）\n\n';

  prompt += '## 1. AI生成レポート（記事の冒頭）\n';
  prompt += 'この記事は、動画・スライド・原稿を元にAI（Gemini）が構成しました。\n';
  prompt += '* **トピック:** [キーワードを抽出]\n';
  prompt += '* **コンテンツ構成:** 動画あり / スライドあり / テキスト解説あり\n';
  prompt += '* **AIの要約:** （コンテンツ全体の要点を140文字以内でズバリ要約）\n\n';

  prompt += '---\n\n';

  prompt += '## 2. 【動画で見る】（推奨）\n';
  prompt += 'まずは、実際の解説動画をご覧ください。雰囲気や詳細なニュアンスが伝わります。\n\n';
  prompt += '**📺 [YouTube動画を再生する]（ {入力されたYouTube URL} ）**\n\n';
  prompt += '**▼動画のポイント**\n';
  prompt += '* （動画を見るべき理由や、動画ならではの見どころを原稿から推測して1行で）\n\n';

  prompt += '---\n\n';

  prompt += '## 3. 【スライドで要点把握】\n';
  prompt += '時間がない方は、こちらのスライドで要点だけサクッと確認できます。\n\n';
  prompt += '**👉 [スライドを表示する]（ {入力されたスライドURL} ）**\n\n';
  prompt += '**▼スライドのハイライト**\n';
  prompt += '* （スライドで特に注目すべきポイントを3つ箇条書き）\n\n';

  prompt += '---\n\n';

  prompt += '## 4. 【テキスト詳細解説】\n';
  prompt += '（入力原稿をブログ向けに見やすく整形・リライトしてください）\n\n';
  prompt += '### [見出し1]\n';
  prompt += '（原稿の内容を整理）\n\n';
  prompt += '### [見出し2]\n';
  prompt += '（原稿の内容を整理）\n\n';
  prompt += '...\n\n';

  prompt += '---\n\n';

  prompt += '## 5. AIの深掘り考察\n';
  prompt += '（今回のテーマについて、AIならではの補足情報や、動画・スライドでは触れられていないかもしれない一歩踏み込んだ視点を追記）\n\n';

  prompt += '---\n\n';

  prompt += '## 6. 編集後記（Human）\n';
  prompt += '（※ここは私が書くので「ここに一言書いてください」と出力）\n';

  return prompt;
}


/**
 * Gemini APIを呼び出してブログ記事を生成
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
 * マークダウンをGoogleドキュメントに適用
 */
function applyMarkdownToDocument(body, markdown) {
  // 既存のコンテンツをクリア
  body.clear();
  
  // 行ごとに処理
  const lines = markdown.split('\n');
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // 見出し1 (# )
    if (line.startsWith('# ')) {
      const text = line.substring(2);
      const paragraph = body.appendParagraph(text);
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    }
    // 見出し2 (## )
    else if (line.startsWith('## ')) {
      const text = line.substring(3);
      const paragraph = body.appendParagraph(text);
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    }
    // 見出し3 (### )
    else if (line.startsWith('### ')) {
      const text = line.substring(4);
      const paragraph = body.appendParagraph(text);
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    }
    // 見出し4 (#### )
    else if (line.startsWith('#### ')) {
      const text = line.substring(5);
      const paragraph = body.appendParagraph(text);
      paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING4);
    }
    // 箇条書き (- または * )
    else if (line.match(/^[\-\*]\s+/)) {
      const text = line.replace(/^[\-\*]\s+/, '');
      const listItem = body.appendListItem(text);
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    }
    // 番号付きリスト (1. 2. など)
    else if (line.match(/^\d+\.\s+/)) {
      const text = line.replace(/^\d+\.\s+/, '');
      const listItem = body.appendListItem(text);
      listItem.setGlyphType(DocumentApp.GlyphType.NUMBER);
    }
    // 空行
    else if (line.trim() === '') {
      body.appendParagraph('');
    }
    // 通常のテキスト
    else {
      const paragraph = body.appendParagraph(line);
      
      // 太字とイタリックを適用
      applyInlineFormatting(paragraph);
    }
  }
}

/**
 * インライン書式(太字・イタリック)を適用
 */
function applyInlineFormatting(paragraph) {
  const text = paragraph.getText();
  
  // **太字** を適用
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(text)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    const contentStart = start + 2;
    const contentEnd = end - 2;
    
    // マークダウン記号を削除して太字を適用
    paragraph.editAsText().deleteText(contentEnd, contentEnd + 1);
    paragraph.editAsText().deleteText(contentStart - 2, contentStart - 1);
    paragraph.editAsText().setBold(contentStart - 2, contentEnd - 3, true);
  }
  
  // *イタリック* を適用
  const italicRegex = /\*(.+?)\*/g;
  const currentText = paragraph.getText();
  while ((match = italicRegex.exec(currentText)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    const contentStart = start + 1;
    const contentEnd = end - 1;
    
    // マークダウン記号を削除してイタリックを適用
    paragraph.editAsText().deleteText(contentEnd, contentEnd);
    paragraph.editAsText().deleteText(contentStart - 1, contentStart - 1);
    paragraph.editAsText().setItalic(contentStart - 1, contentEnd - 2, true);
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
