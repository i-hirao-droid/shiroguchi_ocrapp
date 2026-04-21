/**
 * Webアプリのエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('請求書OCRアプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/**
 * 1ページ分のJPEG画像をGemini APIでOCR処理する
 * @param {Object} pageData - {data: "data:image/jpeg;base64,...", mimeType: "image/jpeg"}
 * @param {number} pageNum - 現在のページ番号
 * @param {number} totalPages - 全ページ数
 * @return {string} JSON文字列（抽出結果）
 */
function processOnePage(pageData, pageNum, totalPages) {
  try {
    var result = extractOnePageWithGemini(pageData, pageNum, totalPages);
    return JSON.stringify(result);
  } catch (e) {
    Logger.log('processOnePage エラー p.' + pageNum + ': ' + e.message);
    return JSON.stringify({ line_items: [], error: e.message });
  }
}


/**
 * 全ページの結果を統合してスプレッドシートに書き込む
 * @param {Object} mergedData - クライアント側で統合済みの請求書データ
 * @param {string} fileName - ファイル名
 * @return {Object} 処理結果
 */
function saveToSpreadsheet(mergedData, fileName) {
  try {
    var spreadsheetUrl = writeToSpreadsheet(mergedData, fileName);
    return {
      success: true,
      spreadsheetUrl: spreadsheetUrl
    };
  } catch (e) {
    Logger.log('スプレッドシート書き込みエラー: ' + e.message);
    throw e;
  }
}

