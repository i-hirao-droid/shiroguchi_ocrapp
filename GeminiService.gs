/**
 * 1ページ分の画像をGemini APIで解析し、請求書データを抽出する
 */
function extractOnePageWithGemini(pageData, pageNum, totalPages) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY がScript Propertiesに設定されていません。');
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=' + apiKey;
  var base64 = pageData.data.split(',')[1];

  var prompt = 'これは請求書の' + pageNum + 'ページ目（全' + totalPages + 'ページ）の画像です。\n' +
    'このページに記載されている情報をJSON形式で抽出してください。\n\n' +
    '- invoice_number: 請求書番号（記載があれば）\n' +
    '- issue_date: 発行日 (YYYY-MM-DD形式、記載があれば)\n' +
    '- due_date: 支払期限 (YYYY-MM-DD形式、記載があれば)\n' +
    '- sender_name: 請求元（発行者）の会社名（記載があれば）\n' +
    '- recipient_name: 請求先（宛先）の会社名（記載があれば）\n' +
    '- delivery_destination: 納品先名（記載があれば）\n' +
    '- subtotal: 小計（税抜合計、商品合計、記載があれば）\n' +
    '- tax_rate: 税率（%、記載があれば）\n' +
    '- tax_amount: 消費税額（記載があれば）\n' +
    '- total_amount: 合計金額（税込、総合計、記載があれば）\n' +
    '- line_items: このページに記載されている明細行の配列。各行に以下を含む:\n' +
    '  - delivery_date: 納品日 (YYYY-MM-DD形式)\n' +
    '  - item_name: 品目名（商品名）\n' +
    '  - quantity: 数量\n' +
    '  - unit_price: 単価\n' +
    '  - jitsuyo_kigo: 実予記号（「実予」という専用列がある場合はその列のアルファベット1文字。専用列がない場合は「単価」と「金額」の境界線上付近に手書きされているアルファベット1文字。ない場合は空文字""）\n' +
    '  - amount: 金額\n\n' +
    '【重要なルール】\n' +
    '1. 実予記号は、「実予」列がある場合はそこから抽出し、ない場合は「単価」と「金額」を区切る縦線上付近から抽出してください（罫線に重なって手書きされていても確実に見つけてください）。\n' +
    '2. 実予記号は黒色だけでなく、赤色（ピンク色）のペンで書かれていることもあります（Kなど）。色が変わっても実予記号として抽出してください。\n' +
    '3. 実予記号は必ず「アルファベット（A-Z）」です。もし「丁」と認識した場合は「J」としてください。\n' +
    '4. 実予記号が「｜」（縦の手書き線）や「↓」で表現されている行は、「上の行と同じ」という意味なので、空文字("")として抽出してください。\n' +
    '5. 品目名（「7F」や「Uボルト」など）の一部を実予記号として誤認しないように注意してください。\n' +
    '6. 金額(amount)から実予記号のアルファベットは除外し、カンマを除去した数値のみにしてください。\n\n' +
    '金額は数値（number型）で返してください。このページに記載がない項目はnullとしてください。\n' +
    '明細行がこのページにない場合はline_itemsを空配列にしてください。';

  var payload = {
    contents: [{
      parts: [
        { inline_data: { mime_type: pageData.mimeType, data: base64 } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          invoice_number: { type: 'STRING', nullable: true },
          issue_date: { type: 'STRING', nullable: true },
          due_date: { type: 'STRING', nullable: true },
          sender_name: { type: 'STRING', nullable: true },
          recipient_name: { type: 'STRING', nullable: true },
          delivery_destination: { type: 'STRING', nullable: true },
          subtotal: { type: 'NUMBER', nullable: true },
          tax_rate: { type: 'NUMBER', nullable: true },
          tax_amount: { type: 'NUMBER', nullable: true },
          total_amount: { type: 'NUMBER', nullable: true },
          line_items: {
            type: 'ARRAY',
            items: {
              type: 'OBJECT',
              properties: {
                delivery_date: { type: 'STRING', nullable: true },
                item_name: { type: 'STRING' },
                quantity: { type: 'NUMBER', nullable: true },
                unit_price: { type: 'NUMBER', nullable: true },
                jitsuyo_kigo: { type: 'STRING', nullable: true },
                amount: { type: 'NUMBER' }
              },
              required: ['item_name', 'amount']
            }
          }
        },
        required: ['line_items']
      }
    }
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var statusCode = response.getResponseCode();

  if (statusCode === 429) {
    Utilities.sleep(2000);
    response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    statusCode = response.getResponseCode();
  }

  if (statusCode !== 200) {
    var errorBody = response.getContentText();
    throw new Error('Gemini APIエラー (HTTP ' + statusCode + '): ' + errorBody.substring(0, 200));
  }

  var result = JSON.parse(response.getContentText());
  var text = result.candidates[0].content.parts
    .filter(function(p) { return p.text; })
    .map(function(p) { return p.text; })
    .join('');

  try {
    return propagateJitsuyoKigo_(JSON.parse(text));
  } catch (e) {
    Logger.log('Gemini応答のJSONパースまたは処理に失敗: ' + e.message);
    throw new Error('データの解析処理中にエラーが発生しました。');
  }
}

/**
 * 実予記号の単一値を正規化する。
 * 丁→J、全角→半角に補正したうえで、A-Zの1文字でなければ空文字（＝継承対象）を返す。
 */
function normalizeJitsuyoKigo_(raw) {
  var kigo = raw != null ? String(raw).trim() : '';
  if (kigo === '') return '';

  // 全角英字→半角、既知の誤認識補正
  kigo = kigo.replace(/丁/g, 'J');
  kigo = kigo.replace(/[Ａ-Ｚ]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
  kigo = kigo.toUpperCase();

  // A-Zの1文字のみ採用。それ以外（｜ I l 1 ↓ - 等）は継承対象として空文字を返す
  return /^[A-Z]$/.test(kigo) ? kigo : '';
}

/**
 * 実予記号が空欄（＝矢印・縦線・同上マーカー）の場合、
 * 上の行のアルファベットを継承する。
 * 次の有効なアルファベットが現れるまでは同じ記号が連続する業務慣習に対応。
 */
function propagateJitsuyoKigo_(data) {
  if (!data.line_items || data.line_items.length === 0) return data;
  var lastSymbol = null;
  data.line_items.forEach(function(item) {
    var kigo = normalizeJitsuyoKigo_(item.jitsuyo_kigo);
    if (kigo !== '') {
      lastSymbol = kigo;
      item.jitsuyo_kigo = kigo;
    } else if (lastSymbol) {
      item.jitsuyo_kigo = lastSymbol;
    } else {
      item.jitsuyo_kigo = '';
    }
  });
  return data;
}

