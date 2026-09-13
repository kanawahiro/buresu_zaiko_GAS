/**
 * 会津「輸入」シートのブレス在庫・全在庫の配列数式を復旧する。
 * 配列数式の展開先には直接値を書き込まない。
 */
function repairAizuBlessStockFormulas() {
  const sheet = SpreadsheetApp.openById(CONFIG.AIDU_SS_ID).getSheetByName(CONFIG.AIDU_SHEET);
  if (!sheet) throw new Error('会津「輸入」シートが見つかりません。');

  const lastRow = sheet.getMaxRows();
  sheet.getRange('V1').clearContent();
  sheet.getRange('Z1').clearContent();
  sheet.getRange('AP1').clearContent();
  SpreadsheetApp.flush();
  sheet.getRange('V2:V' + lastRow).clearContent();
  sheet.getRange('Z2:Z' + lastRow).clearContent();
  sheet.getRange('AP2:AP' + lastRow).clearContent();
  SpreadsheetApp.flush();

  const blockers = ['V8', 'Z8', 'AP8'].filter(function(a1) {
    return sheet.getRange(a1).getValue() !== '';
  });
  if (blockers.length) {
    throw new Error('配列数式の展開先を空にできませんでした: ' + blockers.join(', '));
  }

  const buresuId = CONFIG.BURESU_SS_ID;
  sheet.getRange('V1').setFormula(
    '={"ブレス在庫";ARRAYFORMULA(IF(N2:N="","",IFERROR(XLOOKUP(N2:N,IMPORTRANGE("' + buresuId + '","MSKU輸入商品!A:A"),IMPORTRANGE("' + buresuId + '","MSKU輸入商品!H:H"),""),"")))}'
  );
  sheet.getRange('Z1').setFormula(
    '=ARRAYFORMULA(IF(ROW(AP:AP)=1,"⬅︎全在庫",IFERROR(R:R+V:V+CY:CY+CQ:CQ+CA:CA+BO:BO+BP:BP+CB:CB+CR:CR+CZ:CZ+CU:CU+CV:CV+CE:CE+CF:CF+BS:BS+BT:BT+CM:CM+CN:CN+BW:BW+BX:BX+CI:CI+CJ:CJ,"-")))'
  );
  sheet.getRange('AP1').setFormula('=ARRAYFORMULA(IF(ROW(AP:AP)=1,"全在庫",IFERROR(Z:Z,"-")))');
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('会津のブレス在庫・全在庫の参照式を復旧しました。');
}
