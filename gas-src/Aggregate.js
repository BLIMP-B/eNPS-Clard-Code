/**
 * 正規化済み回答(ステージングシート raw_rows)を店舗単位に集計する。
 *
 * - NPS: 推奨度0-10の回答を 批判(0-6) / 中立(7-8) / 推奨(9-10) に分類し、
 *   NPS = 推奨割合 - 批判割合 (%ポイント) で算出する。
 * - 要因スコア: 各要因設問の「とても当てはまる/当てはまる」等をプラス、
 *   「当てはまらない/全く当てはまらない」等をマイナスとしてカウントする。
 * - 回答率: 分母は稼働数(第32回eNPS稼働数一覧)、分子は対象店舗一覧に
 *   含まれる店舗への有効回答数(FACTOR_DENOMINATOR_MODE等の設定に従う)。
 *
 * このフェーズは正規化行を1店舗コードごとにグルーピングしながら
 * ROW_BATCH_SIZE件ずつストリーム処理し、集計結果は店舗コード単位の
 * オブジェクトとしてまとめてスクリプトプロパティではなく専用シートへ書き出す
 * (件数が多いためプロパティサイズ上限を避ける)。
 */

var POSITIVE_LABELS = ['とても当てはまる', '当てはまる', 'とてもできている', 'できている', 'よくしている'];
var NEGATIVE_LABELS = ['当てはまらない', '全く当てはまらない', 'できていない', '全くできていない'];

function aggregatePhase_(state, config, deadline) {
  var stagingSs = getStagingSheetOrThrow_();
  var rawSheet = stagingSs.getSheetByName('raw_rows');
  var lastRow = rawSheet.getLastRow();

  if (!state.cursor.aggRow) state.cursor.aggRow = 2; // 1行目はヘッダー
  if (!state.cursor.storeAgg) state.cursor.storeAgg = {}; // storeCode -> 集計中間値

  var batchSize = getConfigNumber_(config, 'ROW_BATCH_SIZE', 500);

  while (state.cursor.aggRow <= lastRow) {
    if (Date.now() >= deadline) return { done: false };

    var readCount = Math.min(batchSize, lastRow - state.cursor.aggRow + 1);
    var range = rawSheet.getRange(state.cursor.aggRow, 1, readCount, 3).getValues();

    for (var i = 0; i < range.length; i++) {
      var role = range[i][0];
      var row = JSON.parse(range[i][2]);
      accumulateRow_(state.cursor.storeAgg, role, row);
    }

    state.cursor.aggRow += readCount;
  }

  // 全行処理済み。店舗別の最終集計値を確定してステージングシートへ保存する。
  finalizeStoreAggregates_(state.cursor.storeAgg, config);
  saveAggregateResults_(state.cursor.storeAgg);

  return { done: true };
}

function getStagingSheetOrThrow_() {
  var id = PropertiesService.getScriptProperties().getProperty('STAGING_SHEET_ID');
  if (!id) throw new Error('ステージングデータが見つかりません(INGEST_SURVEYSフェーズ未完了)');
  return SpreadsheetApp.openById(id);
}

/**
 * 1行の回答を店舗別中間集計へ加算する。列位置は行データの構造上、
 * 店舗コードは各アンケートの最初の識別列(社員・パート・FC種別は
 * セグメント文字列の先頭4桁、店長は明示列)から抽出する。
 */
function accumulateRow_(storeAgg, role, row) {
  var storeCode = extractStoreCodeFromRow_(role, row);
  if (!storeCode) return;

  if (!storeAgg[storeCode]) {
    storeAgg[storeCode] = {
      npsScores: [],
      factorAnswers: {}, // 設問インデックス -> {plus, minus, total}
      responseCount: 0
    };
  }
  var agg = storeAgg[storeCode];
  agg.responseCount++;

  var npsValue = extractNpsScoreFromRow_(role, row);
  if (npsValue !== null && npsValue !== undefined && npsValue !== '') {
    agg.npsScores.push(Number(npsValue));
  }

  var factorStartIdx = role === '店長' ? 5 : (role === '店舗' ? 4 : 3);
  for (var c = factorStartIdx; c < row.length; c++) {
    var v = row[c];
    if (POSITIVE_LABELS.indexOf(v) < 0 && NEGATIVE_LABELS.indexOf(v) < 0) continue;
    if (!agg.factorAnswers[c]) agg.factorAnswers[c] = { plus: 0, minus: 0 };
    if (POSITIVE_LABELS.indexOf(v) >= 0) agg.factorAnswers[c].plus++;
    else agg.factorAnswers[c].minus++;
  }
}

function extractStoreCodeFromRow_(role, row) {
  // 社員・役職Ptn / パート・アルバイト / FC: 2列目が "コード_店舗名_直FC_ブランド" 形式
  // 店長: 2列目が店舗コード単独
  var raw = row[1];
  if (raw === undefined || raw === null || raw === '') return null;
  var s = String(raw);
  var underscoreIdx = s.indexOf('_');
  var codePart = underscoreIdx >= 0 ? s.substring(0, underscoreIdx) : s;
  var m = codePart.match(/\d+/);
  return m ? m[0] : null;
}

function extractNpsScoreFromRow_(role, row) {
  // 推奨度設問は各アンケートで3列目(店長のみ4列目)に位置する
  var idx = role === '店長' ? 4 : 2;
  return row[idx];
}

function finalizeStoreAggregates_(storeAgg, config) {
  var codes = Object.keys(storeAgg);
  for (var i = 0; i < codes.length; i++) {
    var agg = storeAgg[codes[i]];
    var promoters = 0, passives = 0, detractors = 0;
    for (var j = 0; j < agg.npsScores.length; j++) {
      var v = agg.npsScores[j];
      if (v >= 9) promoters++;
      else if (v >= 7) passives++;
      else detractors++;
    }
    var npsBase = agg.npsScores.length;
    agg.nps = npsBase > 0 ? Math.round(((promoters - detractors) / npsBase) * 1000) / 10 : null;
    agg.promoters = promoters;
    agg.passives = passives;
    agg.detractors = detractors;
  }
}

function saveAggregateResults_(storeAgg) {
  var ss = getStagingSheetOrThrow_();
  var sheetName = 'store_aggregates';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);
  sheet.appendRow(['storeCode', 'aggregateJson']);

  var codes = Object.keys(storeAgg);
  var rows = codes.map(function (code) {
    return [code, JSON.stringify(storeAgg[code])];
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function loadStoreAggregates_() {
  var ss = getStagingSheetOrThrow_();
  var sheet = ss.getSheetByName('store_aggregates');
  if (!sheet) return {};
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < values.length; r++) {
    map[values[r][0]] = JSON.parse(values[r][1]);
  }
  return map;
}
