function importSpeakerNotesFromSpreadsheet() {
  // ===== 設定 =====
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = 2;
  
  // ===== H列とI列のデータを取得 =====
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log('エラー: データが見つかりません');
    return;
  }
  
  const docIds = sheet.getRange(startRow, 8, lastRow - startRow + 1, 1).getValues();
  const slideIds = sheet.getRange(startRow, 9, lastRow - startRow + 1, 1).getValues();
  
  Logger.log(`=== 処理対象: ${lastRow - startRow + 1}行 ===`);
  
  // ===== 各行を処理 =====
  for (let i = 0; i < docIds.length; i++) {
    const rowNum = startRow + i;
    const docId = String(docIds[i][0]).trim();
    const slideId = String(slideIds[i][0]).trim();
    
    if (!docId || !slideId) continue;
    
    try {
      const doc = DocumentApp.openById(docId);
      const text = doc.getBody().getText();
      
      // ★ ここが変更点（@@ Slide X）
      const slideMatches = [...text.matchAll(
        /@@ Slide (\d+)[^\n]*\n([\s\S]*?)(?=@@ Slide \d+|$)/g
      )];
      
      if (slideMatches.length === 0) {
        Logger.log(`警告: 行${rowNum}にスライド情報なし`);
        continue;
      }
      
      const presentation = SlidesApp.openById(slideId);
      const slides = presentation.getSlides();
      
      let successCount = 0;
      
      slideMatches.forEach(match => {
        const slideNum = Number(match[1]);
        const noteText = match[2].trim();
        const slideIndex = slideNum - 1;
        
        if (slideIndex >= slides.length || !noteText) return;
        
        const notesPage = slides[slideIndex].getNotesPage();
        const textShape = notesPage.getSpeakerNotesShape();
        textShape.getText().setText(noteText);
        successCount++;
      });
      
      Logger.log(`行${rowNum}完了: ${successCount}/${slideMatches.length}`);
      
    } catch (e) {
      Logger.log(`エラー: 行${rowNum} - ${e.message}`);
    }
  }
  
  Logger.log('========== 全処理完了 ==========');
}
