/**
 * Config.js - aidu-bress-idou GAS Webhook 設定
 *
 * このコードは新規プロジェクトです。既存の `aidu_zaiko_kanri_herasu_etc`
 * （会津在庫減算GAS）と `buresu_zaiko_GAS`（ブレスMSKU同期GAS）には
 * 一切手を入れません。同じスプシを共有するため、
 * 触る列は以下に限定し、それ以外には絶対に書き込まない:
 *   - 会津スプシ「輸入」シートの R列（在庫）, S列（確認日付）
 *   - ブレススプシ「MSKU輸入商品」シートの H列（在庫）, I列（確認日付）
 *   - ブレススプシ「指示ログ」シート（新規・このツール専用）
 */

const CONFIG = {
  // ブレス（施設）スプシ — 在庫を +N する
  BURESU_SS_ID: '1pBktMC6MltAhFv93CWxSewlsNXxWvH2RPLweFqPArvU',
  BURESU_SHEET: 'MSKU輸入商品',
  BURESU_COL: {
    MSKU: 1,        // A列
    STOCK: 8,       // H列
    CHECK_DATE: 9,  // I列
  },

  // 会津スプシ — 在庫を -N する
  AIDU_SS_ID: '1zold3yYTASu6d5SE-jSOseC89uhM1BUkSiJ27o5QmEc',
  AIDU_SHEET: '輸入',
  AIDU_COL: {
    NAME: 12,       // L列  商品名
    MSKU: 14,       // N列  キー
    COLOR: 15,      // O列  色
    VARI: 16,       // P列  バリ
    SIZE: 17,       // Q列  サイズ
    STOCK: 18,      // R列  会津在庫
    CHECK_DATE: 19, // S列  確認日付
  },

  // 販売実績シート（読み取り専用 — 在庫表に前回/前々回イベント販売数を表示するため）
  SALES_SS_ID: '1zj6FEWTAXexP2V0_DFyolvWPyT8V2Vb1A2Ve2pXHmHM',
  SALES_SHEET: '販売実績',

  // 受注シートの条件付き書式（Q列）から読み取る出荷元表示。読み取り専用。
  ORDER_SS_ID: '1-e2UQanmr2nY6fkkDQzG507vh0fvPJmC-K0zQ0H-VtU',
  ORDER_SHEET: '送付先単位',
  ORDER_PRODUCT_COLUMN: 17, // Q列「備考 品名」
  SHIPPING_ORIGIN_COLORS: {
    '#ffffff': '会津出荷',
    '#d9ead3': '施設出荷',
    '#fff2cc': '施設出荷',
  },
  SHIPPING_ORIGIN_MIN_TERM_LENGTH: 3,
  // 受注の品名表記と会津マスターの表記が異なるものだけを、確認済みMSKUへ明示対応する。
  // キーは ShippingOrigins.js の shippingManualMappingKey_() で正規化する。
  SHIPPING_ORIGIN_MANUAL_MSKU_MAP: {
    // 受注側の「フリル通常品」は、在庫マスターのエリザベスカラー通常品と同一扱い。
    'フリル通常品ブルー': 'MSKU0001',
    'フリル通常品グリーン': 'MSKU0002',
    'フリル通常品ピンク': 'MSKU0003',
    'フリル通常品イエロー': 'MSKU0004',
    'フリル通常品グレー': 'MSKU0005',
    'pシート骨柄ホワイトs': 'MSKU0035',
    'pシート骨柄ホワイトm': 'MSKU0036',
    'pシート骨柄ホワイトl': 'MSKU0037',
    'pシートホワイトs': 'MSKU0042',
    'pシートホワイトm': 'MSKU0043',
    'pシートホワイトl': 'MSKU0044',
    'pシートシルバーs': 'MSKU0021',
    'pシートシルバーm': 'MSKU0022',
    'pシートシルバーl': 'MSKU0023',
    'カラスネット1.2m×1.2m': 'MSKU0162',
    'カラスネット1.5m×1.5m': 'MSKU0163',
    '洗車タオル中': 'MSKU0178',
    '洗車タオル大': 'MSKU0179',
    '洗車タオル特大': 'MSKU0180',
  },

  // 指示ログ（ブレススプシに新規追加するシート）
  LOG_SHEET: '指示ログ',
  LOG_HEADER: ['バスケットID', '作成日時', 'ステータス', '完了/キャンセル日時', 'items_json', '在庫変動', '想定箱数'],
  LOG_COL: {
    ID: 1,
    CREATED_AT: 2,
    STATUS: 3,
    RESOLVED_AT: 4,
    ITEMS_JSON: 5,
    STOCK_CHANGES: 6,  // 完了時のみ書き込み: [{msku, qty, aidu_before, aidu_after, buresu_before, buresu_after}, ...]
    BOX_COUNT: 7,      // 想定する段ボール箱数（未確定なら 0）
  },

  STATUS: {
    PENDING: '未発送',
    DONE: '完了',
    CANCELLED: 'キャンセル',
    REVOKED: '取消',     // 完了済を在庫巻き戻ししたもの
  },

  // データ行は 2行目から（1行目はヘッダー）
  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  // PropertiesService.getScriptProperties() に保存する secret のキー名
  SECRET_PROPERTY: 'WEBHOOK_SECRET',

  // LockService の待機時間（ms）
  LOCK_TIMEOUT_MS: 10000,
};
