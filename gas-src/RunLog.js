/**
 * clasp run が使えない状況でも実行結果を外部(Drive経由)から確認できるように、
 * 実行のたびにログシートへ1行追記する。
 */
function appendRunLog_(event, detail) {
  try {
    var ss = getOrCreateRunLogSheet_();
    var sheet = ss.getSheets()[0];
    sheet.appendRow([new Date(), event, detail ? String(detail).substring(0, 5000) : '']);
  } catch (e) {
    // ログ機構自体の失敗でジョブを止めない
  }
}

function getOrCreateRunLogSheet_() {
  var props = PropertiesService.getScriptProperties();
  var key = 'RUN_LOG_SHEET_ID';
  var existing = props.getProperty(key);
  if (existing) {
    try { return SpreadsheetApp.openById(existing); } catch (e) { /* 作り直す */ }
  }
  var ss = SpreadsheetApp.create('【実行ログ】eNPSレポート自動発行');
  var file = DriveApp.getFileById(ss.getId());
  getWorkFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  ss.getSheets()[0].appendRow(['timestamp', 'event', 'detail']);
  props.setProperty(key, ss.getId());
  return ss;
}
