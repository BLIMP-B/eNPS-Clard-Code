/**
 * 一度だけ実行するセットアップ用関数群。
 */

/**
 * 集計ロジック.xlsx(変換済みシート)の「設定」シートにある
 * ROOT_FOLDER_ID等のフォルダ系項目を、今回実際に作成したフォルダIDへ
 * 書き換える。係数・閾値等の業務ロジック値は変更しない。
 */
function updateConfigFolderIds_() {
  var sheetId = getOrCreateConvertedConfigSheet_();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('設定');
  var values = sheet.getDataRange().getValues();

  var headerRowIdx = -1, keyCol = -1, valueCol = -1;
  for (var r = 0; r < values.length; r++) {
    if (values[r].indexOf('キー') >= 0 && values[r].indexOf('値') >= 0) {
      headerRowIdx = r;
      keyCol = values[r].indexOf('キー');
      valueCol = values[r].indexOf('値');
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('設定シートのヘッダーが見つかりません');

  var updates = {
    ROOT_FOLDER_ID: BOOTSTRAP.ROOT_FOLDER_ID,
    INPUT_FOLDER_ID: getInputFolder_().getId(),
    SUMMARY_FOLDER_ID: getSummaryFolder_().getId(),
    STORE_REPORT_FOLDER_ID: getStoreReportFolder_().getId(),
    AM_REPORT_FOLDER_ID: getAmReportFolder_().getId(),
    B_REPORT_FOLDER_ID: getBReportFolder_().getId(),
    GAS_PROJECT_ID: ScriptApp.getScriptId()
  };

  var changed = 0;
  for (var r2 = headerRowIdx + 1; r2 < values.length; r2++) {
    var key = values[r2][keyCol];
    if (updates.hasOwnProperty(key)) {
      sheet.getRange(r2 + 1, valueCol + 1).setValue(updates[key]);
      changed++;
    }
  }
  Logger.log('設定シートのフォルダ項目を%s件更新しました', changed);
  CacheService.getScriptCache().remove(CONFIG_CACHE_KEY);
  return changed;
}

/**
 * 列マッピング修正後、今回すでに出力した成果物と中間データを全て削除し、
 * ジョブ状態を初期化して最初から発行テストをやり直せるようにする。
 * (入力ファイル配置用と集計ロジック.xlsxはそのまま残す)
 */
function resetOutputsAndJob() {
  var deleted = 0;
  [getSummaryFolder_(), getStoreReportFolder_(), getAmReportFolder_(), getBReportFolder_()].forEach(function (folder) {
    deleted += trashAllFilesRecursive_(folder);
  });

  var props = PropertiesService.getScriptProperties();
  ['STAGING_SHEET_ID', 'RESOLVED_STORES_SHEET_ID', 'SUMMARY_SHEET_ID', 'ZERO_RESPONSE_STORES',
    'SURVEY_HEADERS', 'RENDER_SHEET_ID'].forEach(function (key) {
    var id = props.getProperty(key);
    if (id) {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 既に削除済み等は無視 */ }
    }
    props.deleteProperty(key);
  });

  resetJob();
  appendRunLog_('RESET_OUTPUTS_AND_JOB', 'deletedFiles=' + deleted);
  Logger.log('出力と中間データを削除し、ジョブを初期化しました(%s件のファイルを削除)', deleted);
  return deleted;
}

function trashAllFilesRecursive_(folder) {
  var count = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    files.next().setTrashed(true);
    count++;
  }
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    count += trashAllFilesRecursive_(subFolders.next());
  }
  return count;
}

/**
 * 稼働数バグ修正後、店舗確定とサマリだけを作り直すためのパッチ関数。
 * (ジョブ全体を再実行せずに済むよう、実行中のジョブとは独立して動かす)
 */
function regenerateSummaryOnly() {
  var config = loadConfig_();
  var stores = resolveStoreMaster_(config);
  saveResolvedStores_(stores);
  var deadline = Date.now() + 25 * 60 * 1000;
  var state = { cursor: {} };
  var result = buildSummaryPhase_(state, config, deadline);
  appendRunLog_('PATCH_REGENERATE_SUMMARY', 'done=' + (result && result.done));
  return result;
}
