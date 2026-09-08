/**
 * 実施店舗の確定ロジック。
 *
 * 1. ファイル名に「実施店舗一覧」を含むファイルから、今回対象の店舗コード集合を確定する。
 * 2. ファイル名が「yyyymmdd_マスタ」形式(例: 20260807_マスタ)のファイルから、
 *    AM名・B長名・ブロック・事業部など不足情報を補完する。
 * どちらも固定IDに依存せず、入力フォルダ内をファイル名パターンで毎回検索する
 * (回によって対象店舗・担当者が変動するため)。
 */

function resolveStoreMaster_(config) {
  var inputFolder = getInputFolder_();

  var rosterFile = findFileByNameContains_(inputFolder, '実施店舗一覧');
  if (!rosterFile) throw new Error('「実施店舗一覧」を含む入力ファイルが見つかりません');

  var masterFile = findFileByNamePattern_(inputFolder, /^\d{8}_マスタ/);
  if (!masterFile) throw new Error('「yyyymmdd_マスタ」形式の入力ファイルが見つかりません');

  var roster = readRosterStores_(rosterFile);
  var masterIndex = readMasterIndex_(masterFile);

  var workforceFile = findFileByNameContains_(getInputFolder_(), '稼働数一覧');
  var workforceIndex = workforceFile ? readWorkforceIndex_(workforceFile) : {};

  var invalidValues = getConfigString_(config, 'INVALID_STORE_SELECTION_VALUES', 'undefined')
    .split(',')
    .map(function (s) { return s.trim(); });

  var stores = [];
  var unmatched = [];
  for (var i = 0; i < roster.length; i++) {
    var r = roster[i];
    if (invalidValues.indexOf(String(r.storeCode)) >= 0) continue;
    var m = masterIndex[String(r.storeCode)];
    if (!m) {
      unmatched.push(r.storeCode);
      continue;
    }
    stores.push({
      storeCode: String(r.storeCode),
      storeName: r.storeName || m.storeName,
      directOrFc: r.directOrFc || m.directOrFc,
      openDate: r.openDate || m.openDate,
      operatingCompany: r.operatingCompany || m.operatingCompany,
      brand: m.brand,
      businessDivision: m.businessDivision,
      block: m.block,
      amName: m.amName,
      amEmployeeId: m.amEmployeeId,
      managerName: m.managerName,
      bLeaderName: m.bLeaderName,
      workforce: workforceIndex[String(r.storeCode)] !== undefined ? workforceIndex[String(r.storeCode)] : m.workforce
    });
  }

  if (unmatched.length > 0) {
    throw new Error('対象店舗一覧の店舗コードがマスタに一致しません: ' + unmatched.join(', '));
  }

  return stores;
}

function saveResolvedStores_(stores) {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.create('【中間データ】確定店舗一覧');
  var file = DriveApp.getFileById(ss.getId());
  getWorkFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  var sheet = ss.getSheets()[0];
  sheet.setName('resolved_stores');
  sheet.appendRow(['storeCode', 'storeJson']);
  var rows = stores.map(function (s) { return [s.storeCode, JSON.stringify(s)]; });
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  props.setProperty('RESOLVED_STORES_SHEET_ID', ss.getId());
}

function loadResolvedStores_() {
  var id = PropertiesService.getScriptProperties().getProperty('RESOLVED_STORES_SHEET_ID');
  if (!id) throw new Error('確定店舗一覧が見つかりません(RESOLVE_STORE_MASTERフェーズ未完了)');
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName('resolved_stores');
  var values = sheet.getDataRange().getValues();
  var stores = [];
  for (var r = 1; r < values.length; r++) {
    stores.push(JSON.parse(values[r][1]));
  }
  return stores;
}

function findFileByNameContains_(folder, needle) {
  var it = folder.getFiles();
  var matches = [];
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(needle) >= 0) matches.push(f);
  }
  if (matches.length === 0) return null;
  if (matches.length > 1) {
    // 最終更新が最も新しいものを正本とする
    matches.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });
  }
  return matches[0];
}

function findFileByNamePattern_(folder, regex) {
  var it = folder.getFiles();
  var matches = [];
  while (it.hasNext()) {
    var f = it.next();
    if (regex.test(f.getName())) matches.push(f);
  }
  if (matches.length === 0) return null;
  matches.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });
  return matches[0];
}

/**
 * 実施店舗一覧xlsxを読み込み、店舗コード・店舗名・直FC・オープン日・経営企業名の
 * 一覧を返す(シート上は2ブロック横並びのレイアウトのため両方を走査する)。
 */
function readRosterStores_(file) {
  var sheetId = convertXlsxToSheet_(file, 'roster');
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();

  var results = [];
  var headerRowIdx = -1;
  for (var r = 0; r < values.length; r++) {
    if (values[r].indexOf('店舗コード') >= 0) { headerRowIdx = r; break; }
  }
  if (headerRowIdx < 0) throw new Error('実施店舗一覧のヘッダー行(店舗コード)が見つかりません');

  var header = values[headerRowIdx];
  var blocks = [];
  for (var c = 0; c < header.length; c++) {
    if (header[c] === '店舗コード') {
      blocks.push({
        status: c - 1 >= 0 ? c - 1 : c,
        code: c,
        name: c + 1,
        open: c + 2,
        directFc: c + 3,
        company: c + 4
      });
    }
  }

  for (var r2 = headerRowIdx + 1; r2 < values.length; r2++) {
    var row = values[r2];
    for (var b = 0; b < blocks.length; b++) {
      var blk = blocks[b];
      var code = row[blk.code];
      if (code === '' || code === null || code === undefined) continue;
      results.push({
        storeCode: code,
        storeName: row[blk.name],
        openDate: row[blk.open],
        directOrFc: row[blk.directFc],
        operatingCompany: row[blk.company]
      });
    }
  }
  return results;
}

/**
 * yyyymmdd_マスタxlsxを読み込み、店舗コードをキーとした店舗情報マップを返す。
 */
function readMasterIndex_(file) {
  var sheetId = convertXlsxToSheet_(file, 'master');
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();

  var headerRowIdx = -1;
  for (var r = 0; r < values.length; r++) {
    if (values[r].indexOf('CODE') >= 0 || values[r].indexOf('コード') >= 0) { headerRowIdx = r; break; }
  }
  if (headerRowIdx < 0) throw new Error('マスタのヘッダー行が見つかりません');

  var header = values[headerRowIdx];
  var colIdx = headerColumnIndexByFuzzyName_(header);

  var index = {};
  for (var r2 = headerRowIdx + 1; r2 < values.length; r2++) {
    var row = values[r2];
    var code = row[colIdx.code];
    if (code === '' || code === null || code === undefined) continue;
    index[String(code)] = {
      storeName: colIdx.storeName >= 0 ? row[colIdx.storeName] : '',
      directOrFc: colIdx.directOrFc >= 0 ? row[colIdx.directOrFc] : '',
      openDate: colIdx.openDate >= 0 ? row[colIdx.openDate] : '',
      operatingCompany: colIdx.operatingCompany >= 0 ? row[colIdx.operatingCompany] : '',
      brand: colIdx.brand >= 0 ? row[colIdx.brand] : '',
      businessDivision: colIdx.businessDivision >= 0 ? row[colIdx.businessDivision] : '',
      block: colIdx.block >= 0 ? row[colIdx.block] : '',
      amName: colIdx.amName >= 0 ? row[colIdx.amName] : '',
      amEmployeeId: colIdx.amEmployeeId >= 0 ? row[colIdx.amEmployeeId] : '',
      managerName: colIdx.managerName >= 0 ? row[colIdx.managerName] : '',
      bLeaderName: colIdx.bLeaderName >= 0 ? row[colIdx.bLeaderName] : '',
      workforce: colIdx.workforce >= 0 ? row[colIdx.workforce] : ''
    };
  }
  return index;
}

/**
 * 「第n回eNPS稼働数一覧.xlsx」から店舗コード→稼働数のマップを読み込む。
 * (マスタファイルには稼働数列が無いため、専用ファイルから取得する)
 */
function readWorkforceIndex_(file) {
  var sheetId = convertXlsxToSheet_(file, 'workforce');
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();

  var headerRowIdx = -1, codeCol = -1, workforceCol = -1;
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var cIdx = row.indexOf('CODE');
    if (cIdx >= 0) {
      headerRowIdx = r;
      codeCol = cIdx;
      workforceCol = row.indexOf('稼働数');
      break;
    }
  }
  if (headerRowIdx < 0 || workforceCol < 0) return {};

  var index = {};
  for (var r2 = headerRowIdx + 1; r2 < values.length; r2++) {
    var code = values[r2][codeCol];
    if (code === '' || code === null || code === undefined) continue;
    index[String(code)] = values[r2][workforceCol];
  }
  return index;
}

/**
 * マスタのヘッダー名は回によって表記揺れがあるため、部分一致で列位置を解決する。
 */
function headerColumnIndexByFuzzyName_(header) {
  function find(candidates) {
    for (var i = 0; i < header.length; i++) {
      var h = String(header[i] || '');
      for (var j = 0; j < candidates.length; j++) {
        if (h.indexOf(candidates[j]) >= 0) return i;
      }
    }
    return -1;
  }
  return {
    code: find(['CODE', 'コード']),
    storeName: find(['屋号', '店舗名']),
    directOrFc: find(['直FC', '直営']),
    openDate: find(['オープン日', '開店日']),
    operatingCompany: find(['経営企業名']),
    brand: find(['業態']),
    businessDivision: find(['事業部']),
    block: find(['ブロック']),
    amName: find(['AM']),
    amEmployeeId: find(['AM社員']),
    managerName: find(['店長名']),
    bLeaderName: find(['ブロック長', 'B長']),
    workforce: find(['稼働数'])
  };
}

/**
 * xlsx入力ファイルをGoogleスプレッドシートへ変換した作業用コピーを返す(キャッシュ付き)。
 */
function convertXlsxToSheet_(xlsxFile, cacheKeySuffix) {
  var props = PropertiesService.getScriptProperties();
  var propKey = 'CONVERTED_' + cacheKeySuffix + '_' + xlsxFile.getId();
  var existing = props.getProperty(propKey);
  if (existing) {
    try {
      SpreadsheetApp.openById(existing);
      return existing;
    } catch (e) { /* 作り直す */ }
  }
  var resource = {
    name: '【変換済み】' + xlsxFile.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [getWorkFolder_().getId()]
  };
  var converted = Drive.Files.copy(resource, xlsxFile.getId());
  props.setProperty(propKey, converted.id);
  return converted.id;
}
