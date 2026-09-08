/**
 * ブートストラップ定数(この2つだけはコードに直接持つ。他の全設定・係数・閾値は
 * 集計ロジック.xlsx の「設定」シートから読み込む)。
 */
var BOOTSTRAP = {
  ROOT_FOLDER_ID: '1KmdbhbCIBQbktbBDtLu6MAVw_gUlsOFr', // 第32回_eNPS (検証用ラウンド)
  CONFIG_XLSX_NAME: '集計ロジック.xlsx'
};

var CONFIG_CACHE_KEY = 'ENPS_CONFIG_V1';
var CONFIG_SHEET_ID_PROP = 'CONFIG_SHEET_ID';

/**
 * 集計ロジック.xlsx を(初回のみ)Googleスプレッドシートへ変換し、
 * 「設定」シートをキー・値のマップとして読み込む。
 * 2回目以降は変換済みシートIDをScript Propertiesから再利用する。
 */
function loadConfig_() {
  var cached = CacheService.getScriptCache().get(CONFIG_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  var sheetId = getOrCreateConvertedConfigSheet_();
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('設定');
  if (!sheet) throw new Error('集計ロジック.xlsx に「設定」シートが見つかりません');

  var values = sheet.getDataRange().getValues();
  var header = null;
  var colIdx = {};
  var config = { _raw: [] };

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    if (!header && row.indexOf('キー') >= 0 && row.indexOf('値') >= 0) {
      header = row;
      for (var c = 0; c < row.length; c++) colIdx[row[c]] = c;
      continue;
    }
    if (!header) continue;
    var key = row[colIdx['キー']];
    var val = row[colIdx['値']];
    if (!key) continue;
    config[String(key).trim()] = val;
    config._raw.push({ key: key, value: val });
  }

  if (Object.keys(config).length <= 1) {
    throw new Error('設定シートからキー・値を読み取れませんでした。ヘッダー行(キー/値)を確認してください。');
  }

  CacheService.getScriptCache().put(CONFIG_CACHE_KEY, JSON.stringify(config), 300);
  return config;
}

function getOrCreateConvertedConfigSheet_() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(CONFIG_SHEET_ID_PROP);
  if (existing) {
    try {
      SpreadsheetApp.openById(existing);
      return existing;
    } catch (e) {
      // 変換済みシートが見つからない場合は作り直す
    }
  }

  var root = DriveApp.getFolderById(BOOTSTRAP.ROOT_FOLDER_ID);
  var it = root.getFilesByName(BOOTSTRAP.CONFIG_XLSX_NAME);
  if (!it.hasNext()) throw new Error(BOOTSTRAP.CONFIG_XLSX_NAME + ' がルートフォルダに見つかりません');
  var xlsxFile = it.next();

  // Drive Advanced Service で xlsx -> Googleスプレッドシートへ変換コピーを作成
  var resource = {
    name: '【変換済み】' + BOOTSTRAP.CONFIG_XLSX_NAME,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [getWorkFolder_().getId()]
  };
  var converted = Drive.Files.copy(resource, xlsxFile.getId());
  props.setProperty(CONFIG_SHEET_ID_PROP, converted.id);
  return converted.id;
}

/**
 * ルート直下のフォルダを名称で解決する(固定IDに依存しない)。
 * 見つからなければ作成する。
 */
function getOrCreateSubFolder_(name) {
  var root = DriveApp.getFolderById(BOOTSTRAP.ROOT_FOLDER_ID);
  var it = root.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return root.createFolder(name);
}

function getInputFolder_() { return getOrCreateSubFolder_('入力ファイル配置用'); }
function getSummaryFolder_() { return getOrCreateSubFolder_('サマリ'); }
function getStoreReportFolder_() { return getOrCreateSubFolder_('店舗レポート'); }
function getAmReportFolder_() { return getOrCreateSubFolder_('AMレポート'); }
function getBReportFolder_() { return getOrCreateSubFolder_('B長レポート'); }
function getWorkFolder_() { return getOrCreateSubFolder_('_GAS作業用'); }

function getConfigNumber_(config, key, defaultValue) {
  var v = config[key];
  if (v === undefined || v === null || v === '') return defaultValue;
  var n = Number(v);
  return isNaN(n) ? defaultValue : n;
}

function getConfigString_(config, key, defaultValue) {
  var v = config[key];
  return (v === undefined || v === null || v === '') ? defaultValue : String(v);
}
