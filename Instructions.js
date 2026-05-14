/**
 * Instructions.js - 指示バスケットの管理
 *
 * - getOrCreateLogSheet_  : 「指示ログ」シートを取得（無ければ新規作成、既存なら触らない）
 * - listInstructions_     : 未発送のみ／全件のリスト返却
 * - createInstruction_    : 1指示バスケットを 1行として追加（在庫は触らない）
 * - completeInstruction_  : 在庫を 会津-N / ブレス+N、確認日付更新、ステータス完了
 * - cancelInstruction_    : ステータスのみキャンセル（在庫は触らない）
 *
 * 共有スプシで触る列を厳格に限定:
 *   会津 R 列, S 列  /  ブレス H 列, I 列  /  ブレス「指示ログ」シート
 * これら以外は読むだけで書かない。
 */

// ---------------- list ----------------

function listInstructions_(payload) {
  const includeAll = !!(payload && payload.includeAll);
  const sheet = getOrCreateLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return { success: true, instructions: [] };
  }
  const rows = sheet.getRange(
    CONFIG.DATA_START_ROW, 1, lastRow - 1, CONFIG.LOG_HEADER.length
  ).getValues();

  const result = [];
  for (const row of rows) {
    const id = row[CONFIG.LOG_COL.ID - 1];
    if (!id) continue;
    const status = row[CONFIG.LOG_COL.STATUS - 1];
    if (!includeAll && status !== CONFIG.STATUS.PENDING) continue;
    let items = [];
    try {
      items = JSON.parse(row[CONFIG.LOG_COL.ITEMS_JSON - 1] || '[]');
    } catch (_) { /* 破損行は items 空で返す */ }
    let stockChanges = [];
    try {
      stockChanges = JSON.parse(row[CONFIG.LOG_COL.STOCK_CHANGES - 1] || '[]');
    } catch (_) { /* 破損行は空で返す */ }
    const rawBox = row[CONFIG.LOG_COL.BOX_COUNT - 1];
    const boxCount = (rawBox === '' || rawBox === null || rawBox === undefined) ? 0 : Number(rawBox);
    result.push({
      id: id,
      created_at: toIsoString_(row[CONFIG.LOG_COL.CREATED_AT - 1]),
      status: status,
      resolved_at: toIsoString_(row[CONFIG.LOG_COL.RESOLVED_AT - 1]),
      items: items,
      stock_changes: stockChanges,
      box_count: Number.isFinite(boxCount) ? boxCount : 0,
    });
  }
  // 新しい順
  result.sort(function (a, b) {
    return (b.created_at || '').localeCompare(a.created_at || '');
  });
  return { success: true, instructions: result };
}

// ---------------- create ----------------

function createInstruction_(payload) {
  const items = (payload && payload.items) || [];
  if (!Array.isArray(items) || items.length === 0) {
    return { success: false, error: 'items は 1件以上の配列で指定してください' };
  }
  const okInt = function (n) { return Number.isFinite(n) && Number.isInteger(n); };
  const hasField = function (it, k) { return (k in it) && it[k] !== null && it[k] !== undefined; };
  for (const it of items) {
    if (!it || !it.msku) return { success: false, error: 'item に msku がありません' };
    const hasMin = hasField(it, 'qty_min');
    const hasMax = hasField(it, 'qty_max');
    const hasNewFields = ('qty_min' in it) || ('qty_max' in it);
    if (hasNewFields) {
      // 少なくとも片方は必要
      if (!hasMin && !hasMax) {
        return { success: false, error: 'qty_min/qty_max のどちらかは必要: ' + it.msku };
      }
      let min = null, max = null;
      if (hasMin) {
        min = Number(it.qty_min);
        if (!okInt(min) || min < 1) {
          return { success: false, error: 'qty_min が不正: ' + it.msku + ' = ' + it.qty_min };
        }
      }
      if (hasMax) {
        max = Number(it.qty_max);
        if (!okInt(max) || max < 1) {
          return { success: false, error: 'qty_max が不正: ' + it.msku + ' = ' + it.qty_max };
        }
      }
      if (hasMin && hasMax && max < min) {
        return { success: false, error: 'qty_max < qty_min: ' + it.msku + ' (' + min + '〜' + max + ')' };
      }
      // qty は未確定(0)でも OK。確定値が入っている場合のみ範囲チェック。
      if (it.qty !== undefined && it.qty !== null && Number(it.qty) !== 0) {
        const q = Number(it.qty);
        if (!okInt(q) || q < 1) {
          return { success: false, error: 'qty が不正: ' + it.msku + ' = ' + it.qty };
        }
        if (hasMin && q < min) {
          return { success: false, error: 'qty が下限未満: ' + it.msku + ' (q=' + q + ' / min=' + min + ')' };
        }
        if (hasMax && q > max) {
          return { success: false, error: 'qty が上限超過: ' + it.msku + ' (q=' + q + ' / max=' + max + ')' };
        }
      }
    } else {
      // 旧クライアント互換（qty のみ）
      const qty = Number(it.qty);
      if (!okInt(qty) || qty <= 0) {
        return { success: false, error: 'qty が不正: ' + it.msku + ' = ' + it.qty };
      }
    }
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    return { success: false, error: 'lock 取得に失敗（しばらく待ってから再試行してください）' };
  }
  try {
    const sheet = getOrCreateLogSheet_();
    const id = generateInstructionId_();
    const now = new Date();
    sheet.appendRow([id, now, CONFIG.STATUS.PENDING, '', JSON.stringify(items), '', 0]);
    return {
      success: true,
      instruction: {
        id: id,
        created_at: now.toISOString(),
        status: CONFIG.STATUS.PENDING,
        resolved_at: null,
        items: items,
        box_count: 0,
      },
    };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- complete ----------------

function completeInstruction_(payload) {
  const id = payload && payload.id;
  if (!id) return { success: false, error: 'id が必要です' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    return { success: false, error: 'lock 取得に失敗（しばらく待ってから再試行してください）' };
  }
  try {
    const logSheet = getOrCreateLogSheet_();
    const rowIdx = findInstructionRow_(logSheet, id);
    if (rowIdx < 0) return { success: false, error: '指示が見つかりません: ' + id };

    const status = logSheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).getValue();
    if (status !== CONFIG.STATUS.PENDING) {
      return { success: false, error: 'すでに ' + status + ' の指示です' };
    }

    const itemsJson = logSheet.getRange(rowIdx, CONFIG.LOG_COL.ITEMS_JSON).getValue();
    let items;
    try {
      items = JSON.parse(itemsJson || '[]');
    } catch (_) {
      return { success: false, error: 'items_json のパースに失敗' };
    }
    if (!items.length) return { success: false, error: '指示の items が空です' };

    // 防御的検証: 全 item の qty が確定済み（> 0、整数）かつ範囲内であること
    const okInt = function (n) { return Number.isFinite(n) && Number.isInteger(n); };
    const hasField = function (it, k) { return (k in it) && it[k] !== null && it[k] !== undefined; };
    for (const it of items) {
      const q = Number(it.qty);
      if (!okInt(q) || q <= 0) {
        return { success: false, error: '確定数が未入力の品目があります: ' + (it && it.msku) };
      }
      if (hasField(it, 'qty_min')) {
        const min = Number(it.qty_min);
        if (okInt(min) && q < min) {
          return { success: false, error: '確定数が下限未満: ' + it.msku + ' (q=' + q + ' / min=' + min + ')' };
        }
      }
      if (hasField(it, 'qty_max')) {
        const max = Number(it.qty_max);
        if (okInt(max) && q > max) {
          return { success: false, error: '確定数が上限超過: ' + it.msku + ' (q=' + q + ' / max=' + max + ')' };
        }
      }
    }

    const buresu = SpreadsheetApp.openById(CONFIG.BURESU_SS_ID).getSheetByName(CONFIG.BURESU_SHEET);
    const aidu = SpreadsheetApp.openById(CONFIG.AIDU_SS_ID).getSheetByName(CONFIG.AIDU_SHEET);
    if (!buresu || !aidu) return { success: false, error: '在庫シートが見つかりません' };

    const buresuRowMap = buildMskuRowMap_(buresu, CONFIG.BURESU_COL.MSKU);
    const aiduRowMap = buildMskuRowMap_(aidu, CONFIG.AIDU_COL.MSKU);

    // 事前検証: 全 MSKU が両シートに存在し、会津在庫が足りるか
    const errors = [];
    const plan = []; // { msku, qty, bRow, aRow, bStock, aStock }
    for (const it of items) {
      const qty = Number(it.qty);
      const bRow = buresuRowMap[it.msku];
      const aRow = aiduRowMap[it.msku];
      if (!bRow) { errors.push(it.msku + ': ブレスシートに該当行なし'); continue; }
      if (!aRow) { errors.push(it.msku + ': 会津シートに該当行なし'); continue; }
      const bStock = toNumberSafe_(buresu.getRange(bRow, CONFIG.BURESU_COL.STOCK).getValue());
      const aStock = toNumberSafe_(aidu.getRange(aRow, CONFIG.AIDU_COL.STOCK).getValue());
      if (aStock < qty) {
        errors.push(it.msku + ': 会津在庫 ' + aStock + ' < 送る数 ' + qty);
        continue;
      }
      plan.push({ msku: it.msku, qty: qty, bRow: bRow, aRow: aRow, bStock: bStock, aStock: aStock });
    }
    if (errors.length > 0) {
      // 部分書き込みを起こさないよう、エラー時は何も書かない
      return { success: false, error: '在庫不整合のため中止: ' + errors.join('; ') };
    }

    // 在庫更新（会津 -N、ブレス +N、両方の確認日付を当日に）
    const now = new Date();
    const stockChanges = [];
    for (const p of plan) {
      buresu.getRange(p.bRow, CONFIG.BURESU_COL.STOCK).setValue(p.bStock + p.qty);
      buresu.getRange(p.bRow, CONFIG.BURESU_COL.CHECK_DATE).setValue(now);
      aidu.getRange(p.aRow, CONFIG.AIDU_COL.STOCK).setValue(p.aStock - p.qty);
      aidu.getRange(p.aRow, CONFIG.AIDU_COL.CHECK_DATE).setValue(now);
      stockChanges.push({
        msku: p.msku,
        qty: p.qty,
        aidu_before: p.aStock,
        aidu_after: p.aStock - p.qty,
        buresu_before: p.bStock,
        buresu_after: p.bStock + p.qty,
      });
    }

    // 指示ログ更新
    logSheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).setValue(CONFIG.STATUS.DONE);
    logSheet.getRange(rowIdx, CONFIG.LOG_COL.RESOLVED_AT).setValue(now);
    logSheet.getRange(rowIdx, CONFIG.LOG_COL.STOCK_CHANGES).setValue(JSON.stringify(stockChanges));

    return { success: true, id: id, resolved_at: now.toISOString(), stock_changes: stockChanges };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- cancel ----------------

function cancelInstruction_(payload) {
  const id = payload && payload.id;
  if (!id) return { success: false, error: 'id が必要です' };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    return { success: false, error: 'lock 取得に失敗（しばらく待ってから再試行してください）' };
  }
  try {
    const sheet = getOrCreateLogSheet_();
    const rowIdx = findInstructionRow_(sheet, id);
    if (rowIdx < 0) return { success: false, error: '指示が見つかりません: ' + id };

    const status = sheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).getValue();
    if (status !== CONFIG.STATUS.PENDING) {
      return { success: false, error: 'すでに ' + status + ' の指示です' };
    }
    const now = new Date();
    sheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).setValue(CONFIG.STATUS.CANCELLED);
    sheet.getRange(rowIdx, CONFIG.LOG_COL.RESOLVED_AT).setValue(now);
    return { success: true, id: id, resolved_at: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- set_box_count (想定箱数の更新) ----------------

function setBoxCount_(payload) {
  const id = payload && payload.id;
  if (!id) return { success: false, error: 'id が必要です' };
  const boxCount = Number(payload && payload.box_count);
  if (!Number.isFinite(boxCount) || !Number.isInteger(boxCount) || boxCount < 0) {
    return { success: false, error: 'box_count が不正: ' + (payload && payload.box_count) };
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    return { success: false, error: 'lock 取得に失敗（しばらく待ってから再試行してください）' };
  }
  try {
    const sheet = getOrCreateLogSheet_();
    const rowIdx = findInstructionRow_(sheet, id);
    if (rowIdx < 0) return { success: false, error: '指示が見つかりません: ' + id };

    const status = sheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).getValue();
    if (status !== CONFIG.STATUS.PENDING) {
      return { success: false, error: 'すでに ' + status + ' の指示は編集できません' };
    }
    sheet.getRange(rowIdx, CONFIG.LOG_COL.BOX_COUNT).setValue(boxCount);
    return { success: true, id: id, box_count: boxCount };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- update (qty 編集) ----------------

function updateInstruction_(payload) {
  const id = payload && payload.id;
  const items = payload && payload.items;
  if (!id) return { success: false, error: 'id が必要です' };
  if (!Array.isArray(items) || items.length === 0) {
    return { success: false, error: 'items は 1件以上の配列で指定してください' };
  }
  const okInt = function (n) { return Number.isFinite(n) && Number.isInteger(n); };
  const hasField = function (it, k) { return (k in it) && it[k] !== null && it[k] !== undefined; };
  for (const it of items) {
    if (!it || !it.msku) return { success: false, error: 'item に msku がありません' };
    const q = Number(it.qty);
    // 確定数は 0（未確定）でも保存OK。負数・小数・NaN は不可。
    if (!okInt(q) || q < 0) {
      return { success: false, error: 'qty が不正: ' + it.msku + ' = ' + it.qty };
    }
    const hasMin = hasField(it, 'qty_min');
    const hasMax = hasField(it, 'qty_max');
    const hasNewFields = ('qty_min' in it) || ('qty_max' in it);
    if (hasNewFields) {
      let min = null, max = null;
      if (hasMin) {
        min = Number(it.qty_min);
        if (!okInt(min) || min < 1) {
          return { success: false, error: 'qty_min が不正: ' + it.msku };
        }
      }
      if (hasMax) {
        max = Number(it.qty_max);
        if (!okInt(max) || max < 1) {
          return { success: false, error: 'qty_max が不正: ' + it.msku };
        }
      }
      if (hasMin && hasMax && max < min) {
        return { success: false, error: 'qty_max < qty_min: ' + it.msku };
      }
      if (q > 0) {
        if (hasMin && q < min) {
          return { success: false, error: 'qty が下限未満: ' + it.msku + ' (q=' + q + ' / min=' + min + ')' };
        }
        if (hasMax && q > max) {
          return { success: false, error: 'qty が上限超過: ' + it.msku + ' (q=' + q + ' / max=' + max + ')' };
        }
      }
    } else if (q === 0) {
      // 旧クライアント互換: qty_min/max 無しで qty=0 は無効
      return { success: false, error: 'qty が不正: ' + it.msku + ' = 0' };
    }
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
    return { success: false, error: 'lock 取得に失敗（しばらく待ってから再試行してください）' };
  }
  try {
    const sheet = getOrCreateLogSheet_();
    const rowIdx = findInstructionRow_(sheet, id);
    if (rowIdx < 0) return { success: false, error: '指示が見つかりません: ' + id };

    const status = sheet.getRange(rowIdx, CONFIG.LOG_COL.STATUS).getValue();
    if (status !== CONFIG.STATUS.PENDING) {
      return { success: false, error: 'すでに ' + status + ' の指示は編集できません' };
    }
    sheet.getRange(rowIdx, CONFIG.LOG_COL.ITEMS_JSON).setValue(JSON.stringify(items));
    return { success: true, id: id, items: items };
  } finally {
    lock.releaseLock();
  }
}

// ---------------- helpers ----------------

/**
 * ブレススプシに「指示ログ」シートが無ければ新規作成してヘッダーを書く。
 * 既存シートには基本手を加えないが、列が増えた場合のみヘッダーを末尾に追記する。
 */
function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.BURESU_SS_ID);
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET);
    sheet.getRange(1, 1, 1, CONFIG.LOG_HEADER.length).setValues([CONFIG.LOG_HEADER]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // 既存シートに列が足りない場合のみヘッダー末尾を補充（既存ヘッダーは上書きしない）
  const currentLastCol = sheet.getLastColumn();
  if (currentLastCol < CONFIG.LOG_HEADER.length) {
    const startCol = Math.max(currentLastCol + 1, 1);
    const additional = CONFIG.LOG_HEADER.slice(startCol - 1);
    sheet.getRange(1, startCol, 1, additional.length).setValues([additional]);
  }
  return sheet;
}

function findInstructionRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return -1;
  const ids = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.LOG_COL.ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return CONFIG.DATA_START_ROW + i;
  }
  return -1;
}

function buildMskuRowMap_(sheet, mskuCol) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < CONFIG.DATA_START_ROW) return map;
  const values = sheet.getRange(CONFIG.DATA_START_ROW, mskuCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const m = String(values[i][0] || '').trim();
    if (m) map[m] = CONFIG.DATA_START_ROW + i;
  }
  return map;
}

function generateInstructionId_() {
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const ymd = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const hms = Utilities.formatDate(new Date(), tz, 'HHmmss');
  return 'I-' + ymd + '-' + hms;
}

function toIsoString_(v) {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString();
  return String(v);
}
