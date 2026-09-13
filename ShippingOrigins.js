/**
 * 受注「送付先単位」Q列の条件付き書式から、MSKU別の出荷元表示を生成する。
 *
 * 条件付き書式の定義だけを根拠にする。通常のセル背景や注文履歴は参照しない。
 * 未対応の条件式・対応付け不能な条件は warnings に残し、推測して表示しない。
 */
function loadShippingOrigins_() {
  const orderSheet = SpreadsheetApp.openById(CONFIG.ORDER_SS_ID).getSheetByName(CONFIG.ORDER_SHEET);
  const aiduSheet = SpreadsheetApp.openById(CONFIG.AIDU_SS_ID).getSheetByName(CONFIG.AIDU_SHEET);
  if (!orderSheet) return { success: false, error: `受注シート '${CONFIG.ORDER_SHEET}' が見つかりません` };
  if (!aiduSheet) return { success: false, error: `会津シート '${CONFIG.AIDU_SHEET}' が見つかりません` };

  const candidates = shippingOriginCandidates_(aiduSheet);
  const warnings = [];
  const matches = {};
  let targetRuleCount = 0;

  orderSheet.getConditionalFormatRules().forEach(function(rule, index) {
    if (!shippingRuleIntersectsProductColumn_(rule)) return;
    const origin = CONFIG.SHIPPING_ORIGIN_COLORS[shippingRuleBackground_(rule, orderSheet)];
    if (!origin) return;
    targetRuleCount++;

    const condition = shippingRuleCondition_(rule);
    if (!condition.supported) {
      warnings.push({ rule: index + 1, type: 'unsupported_rule', message: condition.message });
      return;
    }
    const ruleMatches = candidates.filter(function(item) { return condition.matches(item.searchText); });
    if (!ruleMatches.length) {
      warnings.push({ rule: index + 1, type: 'unmatched_rule', message: 'MSKU対応を確認できません: ' + condition.label });
      return;
    }
    // 商品管理番号・販売SKUの完全対応表をこのGASで参照できないため、会津マスターの
    // 表記照合が一意に決まる場合だけ確定表示する。複数候補を同じ出荷元と推測しない。
    if (ruleMatches.length !== 1) {
      warnings.push({ rule: index + 1, type: 'ambiguous_mapping', message: 'MSKU対応が一意に決まりません: ' + condition.label + ' (' + ruleMatches.length + '件)' });
      return;
    }
    const item = ruleMatches[0];
    if (!matches[item.msku]) matches[item.msku] = [];
    matches[item.msku].push({ origin: origin, condition: condition.label });
  });

  const origins = Object.keys(matches).sort().map(function(msku) {
    const reasons = matches[msku];
    const kinds = {};
    reasons.forEach(function(reason) { kinds[reason.origin] = true; });
    const originKinds = Object.keys(kinds);
    return {
      msku: msku,
      origin: originKinds.length === 1 ? originKinds[0] : '条件による',
      conditions: reasons.map(function(reason) { return reason.condition; }),
    };
  });

  return {
    success: true,
    origins: origins,
    fetched_at: new Date().toISOString(),
    target_rule_count: targetRuleCount,
    unmatched_count: warnings.length,
    warnings: warnings,
  };
}

function shippingOriginCandidates_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];
  const width = CONFIG.AIDU_COL.SIZE;
  return sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - 1, width).getDisplayValues()
    .map(function(row) {
      const msku = String(row[CONFIG.AIDU_COL.MSKU - 1] || '').trim();
      if (!msku) return null;
      // 会津マスターの品名・色・バリ・サイズだけで対応付ける。画面表示名や注文履歴は使わない。
      const parts = [
        row[CONFIG.AIDU_COL.NAME - 1], row[CONFIG.AIDU_COL.COLOR - 1],
        row[CONFIG.AIDU_COL.VARI - 1], row[CONFIG.AIDU_COL.SIZE - 1],
      ].map(function(value) { return String(value || '').trim(); }).filter(Boolean);
      return { msku: msku, searchText: shippingNormalize_(parts.join(' ')) };
    }).filter(Boolean);
}

function shippingRuleIntersectsProductColumn_(rule) {
  return rule.getRanges().some(function(range) {
    const first = range.getColumn();
    return first <= CONFIG.ORDER_PRODUCT_COLUMN && first + range.getNumColumns() - 1 >= CONFIG.ORDER_PRODUCT_COLUMN;
  });
}

function shippingRuleBackground_(rule, sheet) {
  // ConditionalFormatRule 自身には背景色 getter がない。真偽条件の書式から読む。
  const condition = rule.getBooleanCondition();
  if (!condition) return '';
  if (condition.getBackgroundObject) {
    const color = condition.getBackgroundObject();
    if (color && color.asRgbColor) {
      try {
        return color.asRgbColor().asHexString().toLowerCase();
      } catch (_) {
        // テーマ色は asRgbColor() では取得できないため、下で具体色へ解決する。
      }
    }
    if (color && color.asThemeColor && sheet) {
      try {
        const themeType = color.asThemeColor().getThemeColorType();
        return sheet.getParent().getSpreadsheetTheme().getConcreteColor(themeType)
          .asRgbColor().asHexString().toLowerCase();
      } catch (_) {
        // テーマを解決できない色は対象外として扱う。
      }
    }
  }
  return condition.getBackground ? String(condition.getBackground() || '').toLowerCase() : '';
}

function shippingRuleCondition_(rule) {
  const condition = rule.getBooleanCondition();
  if (!condition) return { supported: false, message: '真偽条件ではない条件付き書式です' };
  const type = String(condition.getCriteriaType());
  const values = condition.getCriteriaValues().map(function(value) { return String(value || ''); });
  if (type === 'TEXT_CONTAINS') return shippingContainsCondition_(values[0]);
  if (type === 'CUSTOM_FORMULA') return shippingCustomFormulaCondition_(values[0]);
  return { supported: false, message: '未対応の条件式: ' + type };
}

function shippingContainsCondition_(term) {
  const normalized = shippingNormalize_(term);
  if (normalized.length < CONFIG.SHIPPING_ORIGIN_MIN_TERM_LENGTH) {
    return { supported: false, message: '照合語が短すぎます: ' + term };
  }
  return {
    supported: true,
    label: '品名に「' + term + '」を含む',
    matches: function(text) { return text.indexOf(normalized) !== -1; },
  };
}

function shippingCustomFormulaCondition_(formula) {
  // Q列だけを参照する単一の SEARCH / REGEXMATCH、または SEARCH の OR 結合だけを許可する。
  // AND、他列参照、注文全体参照などは商品単位へ縮約せず、未対応として残す。
  const body = String(formula || '').trim().replace(/^=/, '').trim();
  const ref = '\\$?Q\\$?\\d+';
  const quoted = '"((?:[^"\\\\]|\\\\.|"")*)"';
  const searchAtom = '(?:ISNUMBER\\s*\\(\\s*)?SEARCH\\(\\s*' + quoted + '\\s*,\\s*' + ref + '\\s*\\)\\s*\\)?';
  const directSearch = new RegExp('^' + searchAtom + '$', 'i');
  const directMatch = directSearch.exec(body);
  if (directMatch) return shippingSearchTermsCondition_([directMatch[1].replace(/""/g, '"')]);

  const orMatch = /^OR\((.*)\)$/i.exec(body);
  if (orMatch) {
    const atom = new RegExp(searchAtom, 'gi');
    const terms = [];
    let found;
    while ((found = atom.exec(orMatch[1]))) terms.push(found[1].replace(/""/g, '"'));
    // 原子式を除いた残りが区切りのカンマと空白だけなら、他の条件を含まない安全な OR。
    const remainder = orMatch[1].replace(new RegExp(searchAtom, 'gi'), '').replace(/[\s,]/g, '');
    if (terms.length && !remainder) return shippingSearchTermsCondition_(terms);
  }

  const regex = new RegExp('^REGEXMATCH\\(\\s*' + ref + '\\s*,\\s*' + quoted + '\\s*\\)$', 'i').exec(body);
  if (regex) {
    try {
      const expression = new RegExp(regex[1].replace(/""/g, '"'), 'i');
      return { supported: true, label: '品名の正規表現: ' + regex[1], matches: function(text) { return expression.test(text); } };
    } catch (err) {
      return { supported: false, message: '無効な正規表現: ' + regex[1] };
    }
  }
  return { supported: false, message: '商品単位へ安全に解釈できないカスタム数式: ' + formula };
}

function shippingSearchTermsCondition_(terms) {
  const conditions = terms.map(shippingContainsCondition_);
  const unsupported = conditions.find(function(item) { return !item.supported; });
  if (unsupported) return unsupported;
  return {
    supported: true,
    label: conditions.map(function(item) { return item.label; }).join(' / '),
    matches: function(text) { return conditions.some(function(item) { return item.matches(text); }); },
  };
}

function shippingNormalize_(value) {
  return String(value || '').normalize('NFKC').toLowerCase().replace(/[\s\-‐‑‒–—―_・\/／]/g, '');
}
