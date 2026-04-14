// ===== 設定 =====
const MASTER_FILE_ID = '1zold3yYTASu6d5SE-jSOseC89uhM1BUkSiJ27o5QmEc';
const MASTER_SHEET_NAME = '輸入';
const MASTER_MSKU_COL = 13;      // N列（0始まり）

const FACILITY_SHEET_NAME = 'MSKU輸入商品';
const FACILITY_MSKU_COL = 0;     // A列

// 大元から取得する列（0始まり）と、施設シートの書き込み先列（1始まり）
const COL_MAP = [
  { masterCol: 0,  facilityCol: 2 },  // A → B
  { masterCol: 11, facilityCol: 3 },  // L → C
  { masterCol: 12, facilityCol: 4 },  // M → D
  { masterCol: 14, facilityCol: 5 },  // O → E
  { masterCol: 15, facilityCol: 6 },  // P → F
  { masterCol: 16, facilityCol: 7 },  // Q → G
];
// ================

function syncMasterToFacility() {
  // 大元シートをMapに変換
  let masterSS;
  try {
    masterSS = SpreadsheetApp.openById(MASTER_FILE_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '大元スプレッドシートを開けませんでした。\n\n' +
      '確認事項:\n' +
      '・大元シートへのアクセス権限があるか\n' +
      '・ファイルIDが正しいか\n\n' +
      '詳細: ' + e.message
    );
    return;
  }
  const masterSheet = masterSS.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();

  const masterMap = {};
  for (let i = 1; i < masterData.length; i++) {
    const msku = masterData[i][MASTER_MSKU_COL];
    if (msku) masterMap[String(msku).trim()] = masterData[i];
  }

  // 施設シートを読み込み
  const facilitySS = SpreadsheetApp.getActiveSpreadsheet();
  const facilitySheet = facilitySS.getSheetByName(FACILITY_SHEET_NAME);
  const facilityData = facilitySheet.getDataRange().getValues();

  let updated = 0;
  let skipped = 0;
  let notFound = [];

  for (let i = 1; i < facilityData.length; i++) {
    const msku = String(facilityData[i][FACILITY_MSKU_COL]).trim();
    if (!msku) continue;

    // B列（facilityCol=2、0始まりだとindex=1）に値があればスキップ
    const hasData = COL_MAP.some(({ facilityCol }) => {
      const val = facilityData[i][facilityCol - 1];
      return val !== '' && val !== null && val !== undefined;
    });
    if (hasData) {
      skipped++;
      continue;
    }

    if (masterMap[msku]) {
      const masterRow = masterMap[msku];
      // B〜G列（facilityCol 2〜7）をまとめて1回で書き込む
      const rowValues = COL_MAP.map(({ masterCol }) => masterRow[masterCol]);
      facilitySheet.getRange(i + 1, 2, 1, rowValues.length).setValues([rowValues]);
      updated++;
    } else {
      notFound.push(msku);
    }
  }

  // 完了メッセージ
  let msg = `同期完了：${updated}件を更新しました。`;
  if (skipped > 0) {
    msg += `\n既にデータあり（スキップ）：${skipped}件`;
  }
  if (notFound.length > 0) {
    msg += `\n\n大元シートに見つからなかったMSKU（${notFound.length}件）:\n${notFound.join('\n')}`;
  }
  SpreadsheetApp.getUi().alert(msg);
}

// メニューに追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫同期')
    .addItem('商品情報を同期（大元→施設）', 'syncMasterToFacility')
    .addToUi();
}
