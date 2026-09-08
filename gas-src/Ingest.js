/**
 * 大容量CSVアンケートファイルの分割取得・正規化。
 *
 * - Drive APIのRangeリクエストでファイルをCSV_CHUNK_BYTES単位に分割取得し、
 *   一度に全量をメモリへ載せない。
 * - 取得済みバイトオフセットと未処理の行バッファをジョブ状態(Script Properties)へ
 *   保存し、実行時間予算内で処理できなかった分は次回実行時に続きから再開する。
 * - 正規化後の行はROW_BATCH_SIZEごとに中間スプレッドシート(_GAS作業用)へ追記し、
 *   後続の集計フェーズが逐次読み出せるようにする(全件をメモリ保持しない)。
 */

var SURVEY_SOURCES = [
  { key: 'shain_yakushoku', nameContains: '社員・役職Ptn', role: '社員・役職ptn' },
  { key: 'part_arbeit', nameContains: 'パート・アルバイト', role: 'パート・アルバイト' },
  { key: 'fc', nameContains: 'FC★従業員アンケート', role: 'FC' },
  { key: 'tencho', nameContains: '店長★従業員アンケート', role: '店長' }
];

function ingestSurveysPhase_(state, config, deadline) {
  if (!state.cursor.sourceIndex) state.cursor.sourceIndex = 0;
  if (!state.cursor.byteOffset) state.cursor.byteOffset = 0;
  if (!state.cursor.headerBySource) state.cursor.headerBySource = {};
  if (!state.cursor.pendingTail) state.cursor.pendingTail = {};

  var inputFolder = getInputFolder_();
  var chunkBytes = getConfigNumber_(config, 'CSV_CHUNK_BYTES', 524288);
  var rowBatchSize = getConfigNumber_(config, 'ROW_BATCH_SIZE', 500);

  while (state.cursor.sourceIndex < SURVEY_SOURCES.length) {
    if (Date.now() >= deadline) return { done: false };

    var src = SURVEY_SOURCES[state.cursor.sourceIndex];
    var file = findFileByNameContains_(inputFolder, src.nameContains);
    if (!file) throw new Error('入力ファイルが見つかりません: ' + src.nameContains);

    var finished = ingestOneCsv_(file, src, state, chunkBytes, rowBatchSize, deadline);
    if (!finished) return { done: false }; // 時間切れ、同じソースを次回も継続

    state.cursor.sourceIndex++;
    state.cursor.byteOffset = 0;
    saveJobState_(state);
  }

  return { done: true };
}

/**
 * 1ファイル分をチャンク単位で読み進める。全チャンクを読み切ればtrueを返す。
 * 時間切れの場合はfalseを返し、byteOffset等のカーソルは呼び出し元のstateに保持される。
 */
function ingestOneCsv_(file, src, state, chunkBytes, rowBatchSize, deadline) {
  var fileId = file.getId();
  var totalSize = Number(file.getSize());
  var sheet = getOrCreateStagingSheet_();
  var tailKey = src.key;
  var rowBuffer = [];

  while (state.cursor.byteOffset < totalSize) {
    if (Date.now() >= deadline) {
      flushRowBuffer_(sheet, rowBuffer);
      return false;
    }

    var start = state.cursor.byteOffset;
    var end = Math.min(start + chunkBytes - 1, totalSize - 1);
    var chunkText = fetchDriveByteRange_(fileId, start, end);

    var tail = state.cursor.pendingTail[tailKey] || '';
    var combined = tail + chunkText;
    var isLastChunk = end >= totalSize - 1;

    var lastNewline = combined.lastIndexOf('\n');
    var completePart, newTail;
    if (isLastChunk) {
      completePart = combined;
      newTail = '';
    } else if (lastNewline >= 0) {
      completePart = combined.substring(0, lastNewline);
      newTail = combined.substring(lastNewline + 1);
    } else {
      completePart = '';
      newTail = combined;
    }
    state.cursor.pendingTail[tailKey] = newTail;

    if (completePart) {
      var parsedRows = Utilities.parseCsv(completePart);
      for (var i = 0; i < parsedRows.length; i++) {
        var row = parsedRows[i];
        if (!state.cursor.headerBySource[tailKey]) {
          state.cursor.headerBySource[tailKey] = row;
          continue; // ヘッダー行は保存のみ
        }
        rowBuffer.push({ role: src.role, sourceKey: src.key, row: row });
        if (rowBuffer.length >= rowBatchSize) {
          flushRowBuffer_(sheet, rowBuffer);
          rowBuffer = [];
        }
      }
    }

    state.cursor.byteOffset = end + 1;
    saveJobState_(state);
  }

  flushRowBuffer_(sheet, rowBuffer);
  return true;
}

/**
 * Drive APIでファイルの指定バイト範囲だけを取得する(大容量ファイルの分割取得)。
 */
function fetchDriveByteRange_(fileId, startByte, endByte) {
  var url = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media';
  var resp = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      Range: 'bytes=' + startByte + '-' + endByte
    },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code !== 206 && code !== 200) {
    throw new Error('Drive分割取得に失敗しました (HTTP ' + code + '): ' + resp.getContentText().substring(0, 200));
  }
  return resp.getContentText('UTF-8');
}

function getOrCreateStagingSheet_() {
  var props = PropertiesService.getScriptProperties();
  var key = 'STAGING_SHEET_ID';
  var existing = props.getProperty(key);
  if (existing) {
    try {
      return SpreadsheetApp.openById(existing);
    } catch (e) { /* 作り直す */ }
  }
  var ss = SpreadsheetApp.create('【中間データ】eNPS集計ステージング');
  var file = DriveApp.getFileById(ss.getId());
  getWorkFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  props.setProperty(key, ss.getId());

  var sheet = ss.getSheets()[0];
  sheet.setName('raw_rows');
  sheet.appendRow(['role', 'sourceKey', 'rowJson']);
  return ss;
}

function flushRowBuffer_(ss, rowBuffer) {
  if (rowBuffer.length === 0) return;
  var sheet = ss.getSheetByName('raw_rows');
  var data = rowBuffer.map(function (r) {
    return [r.role, r.sourceKey, JSON.stringify(r.row)];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, 3).setValues(data);
}
