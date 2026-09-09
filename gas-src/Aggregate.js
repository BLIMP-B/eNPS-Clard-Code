/**
 * 正規化済み回答(ステージングシート raw_rows)を店舗単位に集計する。
 *
 * 集計ロジック.xlsx の「指標定義」シートに定義された正式な計算式に従う
 * (2026-09-09にサブエージェントで全文抽出・確認済み)。要点:
 *
 * - NPS = PROMOTERS/ACCEPTED_RESPONSES - (WEAK_DETRACTORS+STRONG_DETRACTORS)/ACCEPTED_RESPONSES
 *   ACCEPTED_RESPONSES(店舗受理回答数)が分母。推奨度が空欄の受理行も分母に残る。
 * - 要因スコア(BASE_FACTOR) = SUM(因子の5段階リッカート点数)/ACCEPTED_RESPONSES
 *   リッカート点数は 100/75/50/25/1 (とても当てはまる=100 ... 全く当てはまらない=1)。
 *   単純な「プラス回答数÷(プラス+マイナス)」ではない。分母は各アンケート種別
 *   (sourceKey)ごとの受理回答数(FACTOR_DENOMINATOR_MODE=ACCEPTED_STORE_RESPONSES)。
 * - WEIGHTED_FACTOR(推奨度で重み付けした要因値)は本バージョンでは未実装(次段階で対応)。
 * - 回答率 = MIN(ACCEPTED_RESPONSES/稼働数, RESPONSE_RATE_CAP)
 *
 * このフェーズは正規化行を1店舗コードごとにグルーピングしながら
 * ROW_BATCH_SIZE件ずつストリーム処理し、集計結果は店舗コード単位の
 * オブジェクトとしてまとめてスクリプトプロパティではなく専用シートへ書き出す
 * (件数が多いためプロパティサイズ上限を避ける)。
 */

/**
 * リッカート5段階回答文言 -> 得点。集計ロジック.xlsxの指標定義シートに
 * 明示の数式は「100/75相当肯定件数」等の記述のみのため、標準的な5段階換算
 * (100/75/50/25/1)を採用する。得点そのものは係数のため、将来的には
 * 集計ロジック.xlsxの「係数・ウェイト」シートから読み込む設計に切り替える。
 */
var LIKERT_SCORE_MAP_ = {
  'とても当てはまる': 100, 'とてもできている': 100, 'よくしている': 100,
  '当てはまる': 75, 'できている': 75, 'したことがある': 75,
  'どちらとも言えない': 50, 'わからない': 50,
  '当てはまらない': 25, 'できていない': 25, 'あまりしていない': 25,
  '全く当てはまらない': 1, '全くできていない': 1, 'したことがない': 1
};
var POSITIVE_LIKERT_SCORES_ = [100, 75]; // FACTOR_POSITIVE_COUNT: 100/75相当
var NEGATIVE_LIKERT_SCORES_ = [25, 1];   // FACTOR_NEGATIVE_COUNT: 25/1相当(50点は含めない)

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
      var sourceKey = range[i][1];
      var row = JSON.parse(range[i][2]);
      accumulateRow_(state.cursor.storeAgg, sourceKey, row);
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
 * 各アンケート(sourceKey)ごとの列レイアウト。実データのヘッダーを直接
 * CSV解析して確認した実測値(2026年7月分)。列順が変わるリスクに備え、
 * 起動時に一度だけヘッダーとの整合性を検証する(validateSurveyLayouts_)。
 *   npsCol: 推奨度(0-10)設問の列インデックス
 *   factorStartCol: 要因設問(プラス/マイナスでカウントする設問群)の開始列
 *   storeCodeCol: 店舗コード(または「コード_店舗名_直FC_ブランド」形式)の列
 */
var SURVEY_LAYOUT_ = {
  shain_yakushoku: { storeCodeCol: 1, npsCol: 3, factorStartCol: 4 },
  part_arbeit: { storeCodeCol: 1, npsCol: 2, factorStartCol: 3 },
  fc: { storeCodeCol: 1, npsCol: 2, factorStartCol: 3 },
  tencho: { storeCodeCol: 1, npsCol: 4, factorStartCol: 5 }
};

/**
 * 1行の回答を店舗別中間集計へ加算する。
 */
function accumulateRow_(storeAgg, sourceKey, row) {
  var layout = SURVEY_LAYOUT_[sourceKey];
  if (!layout) throw new Error('未知のアンケート種別です: ' + sourceKey);

  var storeCode = extractStoreCodeFromRow_(layout, row);
  if (!storeCode) return;

  if (!storeAgg[storeCode]) {
    storeAgg[storeCode] = {
      npsScores: [],
      factorAnswers: {}, // sourceKey:列 -> {sum, plusCount, minusCount}
      bySourceAcceptedCount: {}, // sourceKey -> ACCEPTED_RESPONSES (要因スコアの分母)
      responseCount: 0 // ACCEPTED_RESPONSES(全ソース合算。NPSの分母)
    };
  }
  var agg = storeAgg[storeCode];
  agg.responseCount++;
  agg.bySourceAcceptedCount[sourceKey] = (agg.bySourceAcceptedCount[sourceKey] || 0) + 1;

  var npsValue = row[layout.npsCol];
  if (npsValue !== null && npsValue !== undefined && npsValue !== '') {
    agg.npsScores.push(Number(npsValue));
  }

  for (var c = layout.factorStartCol; c < row.length; c++) {
    var v = row[c];
    var score = LIKERT_SCORE_MAP_[v];
    if (score === undefined) continue; // 未回答・非該当設問はスキップ(欠損は分母に残す)
    // 4種のアンケートは対象者ごとに設問文言が異なる(同じ列位置でも別の質問)ため、
    // sourceKey付きの複合キーで区別し、他ソースの回答と取り違えないようにする。
    var key = sourceKey + ':' + c;
    if (!agg.factorAnswers[key]) agg.factorAnswers[key] = { sum: 0, plusCount: 0, minusCount: 0 };
    agg.factorAnswers[key].sum += score;
    if (POSITIVE_LIKERT_SCORES_.indexOf(score) >= 0) agg.factorAnswers[key].plusCount++;
    else if (NEGATIVE_LIKERT_SCORES_.indexOf(score) >= 0) agg.factorAnswers[key].minusCount++;
  }
}

function extractStoreCodeFromRow_(layout, row) {
  // 社員・役職Ptn / パート・アルバイト / FC: 指定列が「コード_店舗名_直FC_ブランド」形式
  // 店長: 指定列が店舗コード単独
  var raw = row[layout.storeCodeCol];
  if (raw === undefined || raw === null || raw === '') return null;
  var s = String(raw);
  var underscoreIdx = s.indexOf('_');
  var codePart = underscoreIdx >= 0 ? s.substring(0, underscoreIdx) : s;
  var m = codePart.match(/\d+/);
  return m ? m[0] : null;
}

/**
 * 起動時に一度だけ、想定した列レイアウトが実際のヘッダーと一致しているかを
 * 検証する。ズレていた場合は集計を開始せずエラーで止める(サイレントな
 * 集計ミスを防ぐため)。
 */
function validateSurveyLayouts_() {
  var inputFolder = getInputFolder_();
  var npsHeaderKeyword = 'おすすめ';
  var factorHeaderSuffix = '（要因）';

  SURVEY_SOURCES.forEach(function (src) {
    var file = findFileByNameContains_(inputFolder, src.nameContains);
    if (!file) throw new Error('入力ファイルが見つかりません: ' + src.nameContains);
    var layout = SURVEY_LAYOUT_[src.key];

    var blob = file.getBlob();
    var text = blob.getDataAsString('UTF-8');
    var firstLine = text.substring(0, Math.min(text.length, 200000));
    var headerRow = Utilities.parseCsv(firstLine)[0];

    var npsHeader = String(headerRow[layout.npsCol] || '');
    if (npsHeader.indexOf(npsHeaderKeyword) < 0) {
      throw new Error('列レイアウト検証エラー(' + src.key + '): npsCol=' + layout.npsCol +
        ' の見出しに「' + npsHeaderKeyword + '」が含まれません。実際: ' + npsHeader.substring(0, 60));
    }
    var factorHeader = String(headerRow[layout.factorStartCol] || '');
    if (factorHeader.indexOf(factorHeaderSuffix) < 0) {
      throw new Error('列レイアウト検証エラー(' + src.key + '): factorStartCol=' + layout.factorStartCol +
        ' の見出しが要因設問(「' + factorHeaderSuffix + '」)ではありません。実際: ' + factorHeader.substring(0, 60));
    }
  });
}

function finalizeStoreAggregates_(storeAgg, config) {
  var codes = Object.keys(storeAgg);
  for (var i = 0; i < codes.length; i++) {
    var agg = storeAgg[codes[i]];

    // NPS = PROMOTERS/ACCEPTED_RESPONSES - (WEAK+STRONG_DETRACTORS)/ACCEPTED_RESPONSES
    // (指標定義シート。ACCEPTED_RESPONSES=agg.responseCount が分母。
    //  推奨度が0-4:強批判/5-6:弱批判/7-8:中立/9-10:推奨)
    var promoters = 0, passives = 0, weakDetractors = 0, strongDetractors = 0;
    for (var j = 0; j < agg.npsScores.length; j++) {
      var v = agg.npsScores[j];
      if (v >= 9) promoters++;
      else if (v >= 7) passives++;
      else if (v >= 5) weakDetractors++;
      else strongDetractors++;
    }
    var accepted = agg.responseCount;
    var detractors = weakDetractors + strongDetractors;
    agg.nps = accepted > 0 ? Math.round(((promoters - detractors) / accepted) * 1000) / 10 : null;
    agg.promoters = promoters;
    agg.passives = passives;
    agg.detractors = detractors;

    // BASE_FACTOR = SUM(因子リッカート点数)/ACCEPTED_RESPONSES(そのアンケート種別の受理回答数)
    var factorKeys = Object.keys(agg.factorAnswers);
    for (var k = 0; k < factorKeys.length; k++) {
      var key = factorKeys[k];
      var sourceKey = key.substring(0, key.lastIndexOf(':'));
      var f = agg.factorAnswers[key];
      var sourceAccepted = agg.bySourceAcceptedCount[sourceKey] || 0;
      f.baseFactor = sourceAccepted > 0 ? Math.round((f.sum / sourceAccepted) * 10) / 10 : null;
    }
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
