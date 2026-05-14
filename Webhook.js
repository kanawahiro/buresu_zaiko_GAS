/**
 * Webhook.js - doPost / doGet 入口
 *
 * Vercel API ラッパーから fetch される。
 * リクエストボディ: { action: string, secret: string, payload?: object }
 * レスポンス:       { success: bool, ...result } or { success: false, error: string }
 */

/**
 * 動作確認用 GET エンドポイント。ブラウザで開いて疎通確認に使う。
 * 実装上の処理は何もしない（読み取り副作用ゼロ）。
 */
function doGet(e) {
  return jsonResponse_({
    success: true,
    message: 'aidu-bress-idou webhook is alive',
    timestamp: new Date().toISOString(),
  });
}

/**
 * メイン入口。secret 検証 → action 分岐 → JSON 返却。
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // SECRET 検証（スクリプトプロパティに WEBHOOK_SECRET を設定しておく必要あり）
    const expected = PropertiesService.getScriptProperties().getProperty(CONFIG.SECRET_PROPERTY);
    if (!expected) {
      return jsonResponse_({
        success: false,
        error: 'WEBHOOK_SECRET がスクリプトプロパティに設定されていません',
      });
    }
    if (body.secret !== expected) {
      return jsonResponse_({ success: false, error: 'unauthorized' });
    }

    const action = body.action;
    const payload = body.payload || {};

    let result;
    switch (action) {
      case 'load_inventory':
        result = loadInventory_();
        break;
      case 'list_instructions':
        result = listInstructions_(payload);
        break;
      case 'create_instruction':
        result = createInstruction_(payload);
        break;
      case 'complete_instruction':
        result = completeInstruction_(payload);
        break;
      case 'cancel_instruction':
        result = cancelInstruction_(payload);
        break;
      case 'update_instruction':
        result = updateInstruction_(payload);
        break;
      case 'set_box_count':
        result = setBoxCount_(payload);
        break;
      default:
        result = { success: false, error: 'unknown action: ' + action };
    }
    return jsonResponse_(result);
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: (err && err.message) || String(err),
      stack: (err && err.stack) || '',
    });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
