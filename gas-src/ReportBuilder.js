/**
 * サマリ・店舗PDF・AM PDF・B長PDFの生成。
 *
 * v1方針: まず「実施店舗一覧」相当のサマリ表と、店舗ごとの主要指標を
 * 正しく計算・出力するところまでを確実に動かす。PDFのビジュアル(チャート・
 * レイアウト)は原本と完全一致ではなく、Google スライドの1ページ構成で
 * 主要数値(NPS・回答率・要因スコア上位/下位)を表示する簡易版とし、
 * パイプライン全体(店舗確定→集計→出力→通知)が正しく流れることを
 * 最優先で検証する。ビジュアルは後続の反復で本番相当に近づける。
 */

function buildSummaryPhase_(state, config, deadline) {
  var stores = loadResolvedStores_();
  var aggMap = loadStoreAggregates_();
  var lowRateThreshold = getConfigNumber_(config, 'LOW_RESPONSE_RATE_THRESHOLD', 0.5);

  var ss = SpreadsheetApp.create('第32回_物語eNPS_実施店舗一覧');
  var file = DriveApp.getFileById(ss.getId());
  getSummaryFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  var sheet = ss.getSheets()[0];
  sheet.setName('実施店舗一覧');
  // 読み取り専用ディレクトリの第32回実績にある実施店舗一覧.xlsxと同じ列構成に揃える。
  var header = ['店舗コード', '店舗名', '直営FC', '経営企業名', '業態', '事業部', 'AM', 'ブロック', 'B長',
    '対象区分', '有効回答数', '稼働数', '回答率', 'NPS', '店舗PDF', 'AM PDF', 'B長PDF', 'QA状態'];
  sheet.appendRow(header);

  var zeroResponseStores = [];
  var lowResponseStores = [];

  var rows = stores.map(function (s) {
    var agg = aggMap[s.storeCode] || { responseCount: 0, nps: null };
    var workforce = Number(s.workforce) || 0;
    var responseCount = agg.responseCount || 0;
    var responseRateCap = getConfigNumber_(config, 'RESPONSE_RATE_CAP', 1);
    var responseRate = workforce > 0 ? Math.min(responseCount / workforce, responseRateCap) : 0;

    if (responseCount === 0) zeroResponseStores.push(s);
    else if (workforce > 0 && responseRate <= lowRateThreshold) lowResponseStores.push({ store: s, count: responseCount });

    return [
      s.storeCode, s.storeName, s.directOrFc, s.operatingCompany, s.brand, s.businessDivision,
      s.amName, s.block, s.bLeaderName,
      responseCount === 0 ? '対象外(回答数0)' : '対象',
      responseCount, workforce,
      workforce > 0 ? Math.round(responseRate * 1000) / 10 + '%' : '-',
      agg.nps === null || agg.nps === undefined ? '-' : agg.nps,
      '', '', '', '' // 店舗PDF/AM PDF/B長PDF/QA状態はレポート生成後に埋める
    ];
  });

  if (rows.length > 0) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('SUMMARY_SHEET_ID', ss.getId());

  var lowRespText = buildLowResponseText_(zeroResponseStores, lowResponseStores);
  var lowRespFile = DriveApp.createFile('回答数0、極少店舗リスト.txt', lowRespText, MimeType.PLAIN_TEXT);
  getSummaryFolder_().addFile(lowRespFile);
  DriveApp.getRootFolder().removeFile(lowRespFile);

  PropertiesService.getScriptProperties().setProperty('ZERO_RESPONSE_STORES', JSON.stringify(zeroResponseStores.map(function (s) { return s.storeCode; })));

  return { done: true };
}

function buildLowResponseText_(zeroStores, lowStores) {
  var lines = [];
  lines.push('回答数0、極少店舗リスト');
  lines.push('');
  lines.push('【回答数0】：レポートファイル発行なし');
  zeroStores.forEach(function (s) {
    lines.push(s.storeCode + ' ' + s.storeName + ' ' + (s.operatingCompany || s.directOrFc));
  });
  lines.push('');
  lines.push('【回答数極少】：レポートファイル発行あり');
  lowStores.forEach(function (item) {
    lines.push(item.count + '件 ' + item.store.storeCode + ' ' + item.store.storeName + ' ' + (item.store.operatingCompany || item.store.directOrFc));
  });
  lines.push('');
  lines.push('生成日時: ' + new Date().toISOString());
  return lines.join('\n');
}

/**
 * 店舗PDFをバッチ生成する。カーソル(処理済みインデックス)で中断・再開する。
 */
function buildStoreReportsPhase_(state, config, deadline) {
  var stores = loadResolvedStores_();
  var aggMap = loadStoreAggregates_();
  var zeroResponseCodes = JSON.parse(PropertiesService.getScriptProperties().getProperty('ZERO_RESPONSE_STORES') || '[]');

  if (!state.cursor.storeIdx) state.cursor.storeIdx = 0;

  while (state.cursor.storeIdx < stores.length) {
    if (Date.now() >= deadline) return { done: false };

    var s = stores[state.cursor.storeIdx];
    if (zeroResponseCodes.indexOf(s.storeCode) < 0) {
      var agg = aggMap[s.storeCode] || { responseCount: 0, nps: null, factorAnswers: {} };
      generateStoreReportPdf_(s, agg, config);
    }
    state.cursor.storeIdx++;
    saveJobState_(state);
  }

  return { done: true };
}

/**
 * 店舗・AM・B長PDFはいずれも「1つの使い回しシートへ内容を書き換えてPDF
 * エクスポートする」方式で生成する。ファイルをレポートの数だけ新規作成すると
 * SlidesApp.create等の1日あたり呼び出し回数クォータをすぐに使い切ってしまう
 * ため(実際に第32回検証で発生)、create系呼び出しは初回の1回だけに抑える。
 */
function getOrCreateRenderSheet_() {
  var props = PropertiesService.getScriptProperties();
  var key = 'RENDER_SHEET_ID';
  var existing = props.getProperty(key);
  if (existing) {
    try { return SpreadsheetApp.openById(existing); } catch (e) { /* 作り直す */ }
  }
  var ss = SpreadsheetApp.create('【PDFレンダリング用】eNPSレポート');
  var file = DriveApp.getFileById(ss.getId());
  getWorkFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  ss.getSheets()[0].setName('render');
  props.setProperty(key, ss.getId());
  return ss;
}

/**
 * タイトル+本文行を使い回しシートへ書き込み、その場でPDFとしてエクスポートする。
 */
function renderTextReportPdf_(title, bodyLines, targetFolder, fileName) {
  var ss = getOrCreateRenderSheet_();
  var sheet = ss.getSheetByName('render');
  sheet.clear();
  sheet.setColumnWidth(1, 640);

  sheet.getRange(1, 1).setValue(title).setFontSize(16).setFontWeight('bold');
  if (bodyLines.length > 0) {
    var rows = bodyLines.map(function (l) { return [l]; });
    sheet.getRange(3, 1, rows.length, 1).setValues(rows).setFontSize(11).setWrap(true);
  }
  SpreadsheetApp.flush();

  var pdfBlob = ss.getAs(MimeType.PDF);
  var pdfFile = targetFolder.createFile(pdfBlob).setName(fileName);
  return pdfFile;
}

/**
 * 要因設問ごとのプラス/マイナス回答数からスコア(プラス回答割合、100点満点)を
 * 算出し、実際の設問文言を添えてスコア降順で返す。原本PDFの「自店舗比較」表に
 * 相当する情報(要因スコア降順の一覧)を、チャートではなくテキスト表として示す。
 */
/**
 * BASE_FACTOR(集計ロジック.xlsx 指標定義シート正本の要因スコア: リッカート点数
 * [100/75/50/25/1]の平均値、0-100点)を、設問文言とともにスコア降順で返す。
 */
function computeFactorScoreRows_(agg) {
  var headers = loadSurveyHeaders_();
  var keys = Object.keys(agg.factorAnswers || {});
  var rows = keys.map(function (key) {
    var sep = key.lastIndexOf(':');
    var sourceKey = key.substring(0, sep);
    var colIdx = Number(key.substring(sep + 1));
    var f = agg.factorAnswers[key];
    var headerRow = headers[sourceKey];
    var label = headerRow && headerRow[colIdx] ? String(headerRow[colIdx]).replace(/（要因）$/, '') : key;
    return { label: label, score: f.baseFactor, plusCount: f.plusCount, minusCount: f.minusCount };
  });
  rows.sort(function (a, b) { return (b.score || 0) - (a.score || 0); });
  return rows;
}

function generateStoreReportPdf_(store, agg, config) {
  var brandFolder = getOrCreateNamedSubfolderOf_(getStoreReportFolder_(), normalizeBrandFolderName_(store.brand));

  var title = 'eNPSレポート  ' + store.storeName + '（' + store.storeCode + '）';
  var body = [
    '所属: ' + (store.directOrFc || '-') + '　業態: ' + (store.brand || '-') + '　事業部: ' + (store.businessDivision || '-'),
    'AM: ' + (store.amName || '-') + '　ブロック: ' + (store.block || '-') + '　B長: ' + (store.bLeaderName || '-'),
    '店長: ' + (store.managerName || '-'),
    '',
    '【eNPSの結果】',
    '有効回答数: ' + (agg.responseCount || 0) + '人 / 稼働数: ' + (store.workforce || '-') + '人',
    'NPS: ' + (agg.nps === null || agg.nps === undefined ? '-' : agg.nps),
    '批判: ' + (agg.detractors || 0) + '人　中立: ' + (agg.passives || 0) + '人　推奨: ' + (agg.promoters || 0) + '人',
    '',
    '【自店舗比較 要因スコア(BASE_FACTOR, 0-100点, 降順)】'
  ];

  var factorRows = computeFactorScoreRows_(agg);
  factorRows.forEach(function (r, i) {
    body.push((i + 1) + '. ' + r.label + '　' + (r.score === null ? '-' : r.score + '点') +
      '（プラス相当' + r.plusCount + '／マイナス相当' + r.minusCount + '）');
  });

  var fileName = store.storeCode + '_eNPS_第58期1回(32)_' + store.storeName + '_' + (store.directOrFc || '') + '.pdf';
  return renderTextReportPdf_(title, body, brandFolder, fileName);
}

/**
 * 業態フォルダ名は固定リストを持たず、マスタファイルの「業態」列の値を
 * そのまま使う。回によって存在する業態は変動するため、ここで固定化しない。
 */
function normalizeBrandFolderName_(brand) {
  var b = String(brand || '').trim();
  return b || 'その他';
}

function getOrCreateNamedSubfolderOf_(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

function buildAmReportsPhase_(state, config, deadline) {
  // v1: AM別ロールアップは店舗レポートと同様の簡易PDFで代替する。
  var stores = loadResolvedStores_();
  var aggMap = loadStoreAggregates_();

  if (!state.cursor.amKeys) {
    var byAm = {};
    stores.forEach(function (s) {
      var key = s.amEmployeeId || s.amName;
      if (!key) return;
      if (!byAm[key]) byAm[key] = { amName: s.amName, stores: [] };
      byAm[key].stores.push(s);
    });
    state.cursor.amMap = byAm;
    state.cursor.amKeys = Object.keys(byAm);
    state.cursor.amIdx = 0;
  }

  while (state.cursor.amIdx < state.cursor.amKeys.length) {
    if (Date.now() >= deadline) return { done: false };
    var key = state.cursor.amKeys[state.cursor.amIdx];
    generateAmReportPdf_(state.cursor.amMap[key], aggMap, config);
    state.cursor.amIdx++;
    saveJobState_(state);
  }
  return { done: true };
}

function generateAmReportPdf_(amGroup, aggMap, config) {
  var title = 'eNPS AMレポート　' + amGroup.amName;
  var lines = amGroup.stores.map(function (s) {
    var agg = aggMap[s.storeCode] || { responseCount: 0, nps: null };
    return s.storeCode + ' ' + s.storeName + '  NPS:' + (agg.nps === null || agg.nps === undefined ? '-' : agg.nps) + '  回答数:' + (agg.responseCount || 0);
  });
  var fileName = 'AM_' + amGroup.amName + '_eNPS_第58期1回(32).pdf';
  return renderTextReportPdf_(title, lines, getAmReportFolder_(), fileName);
}

function buildBReportsPhase_(state, config, deadline) {
  var stores = loadResolvedStores_();
  var aggMap = loadStoreAggregates_();

  if (!state.cursor.bKeys) {
    var byB = {};
    stores.forEach(function (s) {
      var key = s.bLeaderName;
      if (!key) return;
      if (!byB[key]) byB[key] = { bLeaderName: key, stores: [] };
      byB[key].stores.push(s);
    });
    state.cursor.bMap = byB;
    state.cursor.bKeys = Object.keys(byB);
    state.cursor.bIdx = 0;
  }

  while (state.cursor.bIdx < state.cursor.bKeys.length) {
    if (Date.now() >= deadline) return { done: false };
    var key = state.cursor.bKeys[state.cursor.bIdx];
    generateBReportPdf_(state.cursor.bMap[key], aggMap, config);
    state.cursor.bIdx++;
    saveJobState_(state);
  }
  return { done: true };
}

function generateBReportPdf_(bGroup, aggMap, config) {
  var title = 'eNPS B長レポート　' + bGroup.bLeaderName;
  var lines = bGroup.stores.map(function (s) {
    var agg = aggMap[s.storeCode] || { responseCount: 0, nps: null };
    return s.storeCode + ' ' + s.storeName + '  NPS:' + (agg.nps === null || agg.nps === undefined ? '-' : agg.nps) + '  回答数:' + (agg.responseCount || 0);
  });
  var fileName = 'B長_' + bGroup.bLeaderName + '_eNPS_第58期1回(32).pdf';
  return renderTextReportPdf_(title, lines, getBReportFolder_(), fileName);
}

/**
 * 本番のChat Webhook/メール宛先は検証ラウンドでは絶対に叩かない。
 * NOTIFY_DRY_RUNをtrueにしている間はログへ記録するだけに留める。
 * 本番運用に切り替える際は、この検証用ガードを明示的に外すこと。
 */
var NOTIFY_DRY_RUN = true;

function notifyPhase_(state, config, deadline) {
  var webhookUrl = getConfigString_(config, 'CHAT_WEBHOOK_URL', '');
  var message = 'eNPSレポート生成が完了しました(第32回検証)。';

  if (NOTIFY_DRY_RUN) {
    appendRunLog_('NOTIFY_DRY_RUN', 'webhook送信をスキップ(検証モード)。宛先: ' + webhookUrl);
    return { done: true };
  }

  if (webhookUrl) {
    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: message }),
        muteHttpExceptions: true
      });
    } catch (e) {
      Logger.log('通知送信に失敗しました: ' + e);
    }
  }
  return { done: true };
}
