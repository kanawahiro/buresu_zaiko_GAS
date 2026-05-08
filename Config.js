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

  // 指示ログ（ブレススプシに新規追加するシート）
  LOG_SHEET: '指示ログ',
  LOG_HEADER: ['バスケットID', '作成日時', 'ステータス', '完了/キャンセル日時', 'items_json'],
  LOG_COL: {
    ID: 1,
    CREATED_AT: 2,
    STATUS: 3,
    RESOLVED_AT: 4,
    ITEMS_JSON: 5,
  },

  STATUS: {
    PENDING: '未発送',
    DONE: '完了',
    CANCELLED: 'キャンセル',
  },

  // データ行は 2行目から（1行目はヘッダー）
  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  // PropertiesService.getScriptProperties() に保存する secret のキー名
  SECRET_PROPERTY: 'WEBHOOK_SECRET',

  // LockService の待機時間（ms）
  LOCK_TIMEOUT_MS: 10000,
};
