# 在庫同期システム 実装計画書

## 概要

メイン在庫シート（大元）と施設シートの間で、MSKUをキーとしたデータ連携を行う。

| 項目 | 内容 |
|------|------|
| 連携キー | MSKU（バリエーション単位のユニーク番号） |
| ファイル構成 | 別々のスプレッドシート（2ファイル） |
| 機能数 | 2つ（GAS × 1、関数 × 1） |

---

## ファイル情報

### 大元シート（メイン在庫）

| 項目 | 値 |
|------|------|
| ファイルID | `1zold3yYTASu6d5SE-jSOseC89uhM1BUkSiJ27o5QmEc` |
| URL | https://docs.google.com/spreadsheets/d/1zold3yYTASu6d5SE-jSOseC89uhM1BUkSiJ27o5QmEc/edit |
| シート名 | `輸入` |
| MSKUの列 | N列 |
| 施設在庫数を表示する列 | V列（機能②のXLOOKUP設置場所） |

### 施設シート

| 項目 | 値 |
|------|------|
| ファイルID | `1pBktMC6MltAhFv93CWxSewlsNXxWvH2RPLweFqPArvU` |
| URL | https://docs.google.com/spreadsheets/d/1pBktMC6MltAhFv93CWxSewlsNXxWvH2RPLweFqPArvU/edit |
| シート名 | `MSKU輸入商品` |
| MSKUの列 | A列（手入力） |
| 在庫数の入力列 | H列 |

---

## 機能① 商品情報の同期（GAS）

### 概要

大元シートから施設シートへ、MSKUをキーに商品情報をコピーする。  
商品の追加・変更があったときのみ手動実行する運用。

### データフロー

```
大元シート（輸入）
  N列: MSKU ──照合──▶ 施設シート（MSKU輸入商品）
                         A列: MSKU（入力済み）
                         B列 ◀── A列（大元）
                         C列 ◀── L列（大元）
                         D列 ◀── M列（大元）
                         E列 ◀── O列（大元）
                         F列 ◀── P列（大元）
                         G列 ◀── Q列（大元）
```

### 列マッピング

| 大元シートの列 | 施設シートの列 |
|--------------|--------------|
| A列 | B列 |
| L列 | C列 |
| M列 | D列 |
| O列 | E列 |
| P列 | F列 |
| Q列 | G列 |

### 運用フロー

1. 施設シートのA列にMSKUを入力（手作業）
2. 施設スプレッドシートのメニュー「在庫同期」→「商品情報を同期（大元→施設）」を実行
3. B〜G列に値として書き込まれる（数式なし・静的データ）
4. 完了ダイアログで更新件数と未マッチのMSKUを確認

### GASスクリプト

設置場所：施設スプレッドシート の Apps Script

```javascript
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
  const masterSS = SpreadsheetApp.openById(MASTER_FILE_ID);
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
  let notFound = [];

  for (let i = 1; i < facilityData.length; i++) {
    const msku = String(facilityData[i][FACILITY_MSKU_COL]).trim();
    if (!msku) continue;

    if (masterMap[msku]) {
      const masterRow = masterMap[msku];
      COL_MAP.forEach(({ masterCol, facilityCol }) => {
        facilitySheet.getRange(i + 1, facilityCol).setValue(masterRow[masterCol]);
      });
      updated++;
    } else {
      notFound.push(msku);
    }
  }

  // 完了メッセージ
  let msg = `同期完了：${updated}件を更新しました。`;
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
```

### セットアップ手順

1. 施設スプレッドシートを開く
2. メニュー：拡張機能 → Apps Script
3. 上記コードを貼り付けて保存（Ctrl+S）
4. スプレッドシートをリロード
5. メニューに「在庫同期」が追加されていることを確認

---

## 機能② 施設在庫数の表示（XLOOKUP関数）

### 概要

施設シートのH列に入力された在庫数を、大元シートのV列にリアルタイムで表示する。  
GAS不要。施設側が在庫数を更新するたびに自動反映。

### データフロー

```
施設シート（MSKU輸入商品）
  A列: MSKU
  H列: 在庫数（手入力）
        │
        └──XLOOKUP照合──▶ 大元シート（輸入）
                             N列: MSKU
                             V列: 施設在庫数（表示）
```

### 設置する数式

大元シートの **V2セル** に以下を入力し、下方向にコピー：

```
=IFERROR(XLOOKUP(N2, IMPORTRANGE("1pBktMC6MltAhFv93CWxSewlsNXxWvH2RPLweFqPArvU","MSKU輸入商品!A:A"), IMPORTRANGE("1pBktMC6MltAhFv93CWxSewlsNXxWvH2RPLweFqPArvU","MSKU輸入商品!H:H"), ""),"")
```

### セットアップ手順

1. 大元スプレッドシートのV2セルに上記数式を入力
2. 初回のみ「アクセスを許可」ダイアログが表示されるので許可
3. V2セルを選択し、必要な行数分下にコピー

---

## 今後の拡張メモ

- 施設シートに表示させる列（B〜G）の列名・用途が確定したら本計画書に追記する
- 施設が複数になる場合は、シート名を変数化してGASを拡張対応できる構造になっている
- 在庫数以外に施設→大元へ書き戻したい項目が増えた場合は、機能①と同様のGASで逆方向に対応可能
