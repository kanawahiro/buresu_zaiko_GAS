/**
 * Inventory.js - load_inventory アクション（読み取り専用）
 *
 * ブレスシートに登録されている MSKU を一覧の対象とし、
 * 商品名・色・バリ・サイズは会津シート側 (L/O/P/Q) を真として表示する。
 * 在庫数は両シートからそれぞれ取得 (ブレス H / 会津 R)。
 *
 * このファイルからは書き込みを行わない。
 */

function loadInventory_() {
  const buresu = SpreadsheetApp.openById(CONFIG.BURESU_SS_ID).getSheetByName(CONFIG.BURESU_SHEET);
  if (!buresu) {
    return { success: false, error: `ブレスシート '${CONFIG.BURESU_SHEET}' が見つかりません` };
  }
  const aidu = SpreadsheetApp.openById(CONFIG.AIDU_SS_ID).getSheetByName(CONFIG.AIDU_SHEET);
  if (!aidu) {
    return { success: false, error: `会津シート '${CONFIG.AIDU_SHEET}' が見つかりません` };
  }

  const buresuRows = readSheetRows_(buresu, CONFIG.BURESU_COL.CHECK_DATE);
  const aiduRows = readSheetRows_(aidu, CONFIG.AIDU_COL.CHECK_DATE);

  // 会津側を MSKU でインデックス化
  const aiduMap = {};
  for (const row of aiduRows) {
    const msku = String(row[CONFIG.AIDU_COL.MSKU - 1] || '').trim();
    if (!msku) continue;
    aiduMap[msku] = {
      name: String(row[CONFIG.AIDU_COL.NAME - 1] || ''),
      color: String(row[CONFIG.AIDU_COL.COLOR - 1] || ''),
      vari: String(row[CONFIG.AIDU_COL.VARI - 1] || ''),
      size: String(row[CONFIG.AIDU_COL.SIZE - 1] || ''),
      stock: toNumberSafe_(row[CONFIG.AIDU_COL.STOCK - 1]),
    };
  }

  // ブレス側をベースに、会津情報をマージ
  const items = [];
  for (const row of buresuRows) {
    const msku = String(row[CONFIG.BURESU_COL.MSKU - 1] || '').trim();
    if (!msku) continue;
    const buresuStock = toNumberSafe_(row[CONFIG.BURESU_COL.STOCK - 1]);
    const aiduInfo = aiduMap[msku];
    items.push({
      msku: msku,
      name: aiduInfo ? aiduInfo.name : '',
      color: aiduInfo ? aiduInfo.color : '',
      vari: aiduInfo ? aiduInfo.vari : '',
      size: aiduInfo ? aiduInfo.size : '',
      buresu_stock: buresuStock,
      aidu_stock: aiduInfo ? aiduInfo.stock : 0,
      missing_in_aidu: !aiduInfo,
    });
  }

  return { success: true, items: items };
}

/**
 * シートのデータ行（2行目以降）を、必要列幅まで読んで二次元配列で返す
 */
function readSheetRows_(sheet, lastCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];
  return sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - 1, lastCol).getValues();
}

function toNumberSafe_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}
