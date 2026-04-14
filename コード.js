/**
 * 「輸入 特典 資材」シートのF列が編集された時に変更履歴を記録する関数
 * @param {Object} e 編集イベントオブジェクト
 */
function onEdit(e) {
  // 編集されたシート名、範囲、行、列を取得
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const column = range.getColumn();
  
  // 「輸入 特典 資材」シートのF列(6列目)の編集のみを対象とする
  if (sheetName !== '輸入 特典 資材' || column !== 6) {
    return;
  }
  
  // 変更前の値と変更後の値を取得
  const oldValue = e.oldValue !== undefined ? e.oldValue : "";
  const newValue = e.value !== undefined ? e.value : "";

  // 同一行の必要な列の値を取得（A列、C列、D列、F列）
  const dataRange = sheet.getRange(row, 1, 1, 6);
  const rowData = dataRange.getValues()[0];
  const aValue = rowData[0]; // A列の値
  const cValue = rowData[2]; // C列の値
  const dValue = rowData[3]; // D列の値

  // 現在の日時とユーザーメールアドレスを取得

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail() || "";

  // 「輸入 特典 資材変更履歴」シートを取得
  const logSheet = e.source.getSheetByName('輸入 特典 資材変更履歴');
  
  // 変更履歴を記録
  const logData = [
    aValue,      // A列: 輸入 特典 資材シートのA列の値
    cValue,      // B列: 輸入 特典 資材シートのC列の値
    dValue,      // C列: 輸入 特典 資材シートのD列の値
    oldValue,    // D列: 変更前のF列の値
    newValue,    // E列: 変更後のF列の値
    formattedDate, // F列: 変更日時
    userEmail    // G列: 変更者のメールアドレス
  ];

  // 変更履歴シートの2行目に新しい行を挿入
  logSheet.insertRowBefore(2);

  // 変更履歴シートの2行目に記録
  logSheet.getRange(2, 1, 1, 7).setValues([logData]);
}