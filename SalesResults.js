/**
 * SalesResults.js - load_sales_results アクション（読み取り専用）
 *
 * 「販売実績」シートを読み、MSKU ごとの週次販売数を返す。
 * ヘッダー: MSKU, productId, sku, <週ラベル...>（4列目以降が月〜日の週次バケット、左が最新）。
 * 販売数は会津＋施設の合算（参考値）。このファイルからは書き込みを行わない。
 */

function loadSalesResults_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SALES_SS_ID).getSheetByName(CONFIG.SALES_SHEET);
  if (!sheet) {
    return { success: false, error: `販売実績シート '${CONFIG.SALES_SHEET}' が見つかりません` };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { success: true, weekLabels: [], rows: [] };
  }

  // 先頭3列（MSKU / productId / sku）の後ろが週ラベル
  const header = values[0];
  const weekLabels = header.slice(3).map(function (v) { return String(v || '').trim(); });

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const msku = String(row[0] || '').trim();
    if (!msku) continue;
    const weeks = row.slice(3).map(toNumberSafe_);
    rows.push({ msku: msku, weeks: weeks });
  }

  return { success: true, weekLabels: weekLabels, rows: rows };
}
