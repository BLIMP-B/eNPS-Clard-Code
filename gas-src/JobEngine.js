/**
 * ジョブ制実行エンジン。
 * - 初回30分 / 再開6分の時間予算(集計ロジック.xlsxの設定を反映)を守り、
 *   安全余白を残して自動的に処理を中断・再開する。
 * - 進捗はScript Propertiesにフェーズ名とカーソル(処理済み件数/行)で保存する。
 * - いつでも停止(トリガー削除)・再開(startJob再実行)できる。
 */

var JOB_STATE_PROP = 'ENPS_JOB_STATE';
var JOB_TRIGGER_FUNCTION = 'runJob';

/**
 * ジョブの状態遷移フェーズ一覧。各フェーズはステップ関数を持ち、
 * 完了したら次のフェーズへ進む。
 */
var JOB_PHASES = [
  'RESOLVE_STORE_MASTER',
  'INGEST_SURVEYS',
  'AGGREGATE',
  'BUILD_SUMMARY',
  'BUILD_STORE_REPORTS',
  'BUILD_AM_REPORTS',
  'BUILD_B_REPORTS',
  'NOTIFY',
  'DONE'
];

/**
 * 各フェーズのハンドラ。すべて (state, config, deadline) を受け取り、
 * {done:true} で次フェーズへ、{done:false} で時間切れ・同一フェーズ継続。
 */
var PHASE_HANDLERS_ = {
  RESOLVE_STORE_MASTER: function (state, config, deadline) {
    var stores = resolveStoreMaster_(config);
    saveResolvedStores_(stores);
    return { done: true };
  },
  INGEST_SURVEYS: function (state, config, deadline) {
    return ingestSurveysPhase_(state, config, deadline);
  },
  AGGREGATE: function (state, config, deadline) {
    return aggregatePhase_(state, config, deadline);
  },
  BUILD_SUMMARY: function (state, config, deadline) {
    return buildSummaryPhase_(state, config, deadline);
  },
  BUILD_STORE_REPORTS: function (state, config, deadline) {
    return buildStoreReportsPhase_(state, config, deadline);
  },
  BUILD_AM_REPORTS: function (state, config, deadline) {
    return buildAmReportsPhase_(state, config, deadline);
  },
  BUILD_B_REPORTS: function (state, config, deadline) {
    return buildBReportsPhase_(state, config, deadline);
  },
  NOTIFY: function (state, config, deadline) {
    return notifyPhase_(state, config, deadline);
  }
};

function getJobState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(JOB_STATE_PROP);
  if (!raw) {
    return {
      phase: JOB_PHASES[0],
      cursor: {},
      startedAt: null,
      runCount: 0,
      errors: []
    };
  }
  return JSON.parse(raw);
}

function saveJobState_(state) {
  PropertiesService.getScriptProperties().setProperty(JOB_STATE_PROP, JSON.stringify(state));
}

function clearJobState_() {
  PropertiesService.getScriptProperties().deleteProperty(JOB_STATE_PROP);
}

/**
 * 外部から呼び出すエントリポイント。ジョブが存在しなければ新規開始、
 * 既にあれば続きから再開する。
 */
function startJob() {
  clearAllTriggers_(JOB_TRIGGER_FUNCTION);
  var state = getJobState_();
  state.runCount = 0;
  saveJobState_(state);
  runJob();
}

/**
 * トリガーからも手動からも呼ばれる、時間予算内でフェーズを進める本体。
 */
function runJob() {
  var config = loadConfig_();
  var state = getJobState_();
  state.runCount = (state.runCount || 0) + 1;

  var isFirstRun = state.runCount === 1 && !state.startedAt;
  var maxMinutes = isFirstRun
    ? getConfigNumber_(config, 'INITIAL_MAX_MINUTES', 30)
    : getConfigNumber_(config, 'RESUME_MAX_MINUTES', 6);
  var safetySeconds = isFirstRun
    ? getConfigNumber_(config, 'INITIAL_SAFETY_SECONDS', 45)
    : getConfigNumber_(config, 'RESUME_SAFETY_SECONDS', 30);

  var budgetMs = maxMinutes * 60 * 1000 - safetySeconds * 1000;
  var deadline = Date.now() + budgetMs;
  if (!state.startedAt) state.startedAt = new Date().toISOString();

  Logger.log('runJob phase=%s runCount=%s budgetMs=%s', state.phase, state.runCount, budgetMs);

  try {
    while (Date.now() < deadline && state.phase !== 'DONE') {
      var phaseFn = PHASE_HANDLERS_[state.phase];
      if (!phaseFn) throw new Error('未知のフェーズ: ' + state.phase);

      var result = phaseFn(state, config, deadline);
      // フェーズ関数は {done:true} を返せば次フェーズへ、
      // {done:false} を返せば同一フェーズを維持したまま時間切れで抜ける。
      if (result && result.done) {
        state.phase = nextPhase_(state.phase);
        state.cursor = {};
      }
      saveJobState_(state);
      if (result && result.done === false) break; // このフェーズは時間切れで中断
    }
  } catch (e) {
    state.errors.push({ at: new Date().toISOString(), phase: state.phase, message: String(e), stack: e.stack });
    saveJobState_(state);
    scheduleResume_(1); // エラー時も再試行できるよう次回トリガーは張る
    throw e;
  }

  if (state.phase === 'DONE') {
    clearAllTriggers_(JOB_TRIGGER_FUNCTION);
    Logger.log('ジョブ完了');
    return;
  }

  // 時間切れ・未完了 → 次回再開トリガーを設定して終了
  scheduleResume_(1);
}

function nextPhase_(current) {
  var idx = JOB_PHASES.indexOf(current);
  if (idx < 0 || idx === JOB_PHASES.length - 1) return 'DONE';
  return JOB_PHASES[idx + 1];
}

function scheduleResume_(minutesFromNow) {
  clearAllTriggers_(JOB_TRIGGER_FUNCTION);
  ScriptApp.newTrigger(JOB_TRIGGER_FUNCTION)
    .timeBased()
    .after(minutesFromNow * 60 * 1000)
    .create();
}

function clearAllTriggers_(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/** 明示的にジョブを停止する(トリガーのみ削除。進捗は保持され後で再開可能)。 */
function stopJob() {
  clearAllTriggers_(JOB_TRIGGER_FUNCTION);
}

/** 進捗を完全に破棄して最初からやり直す。 */
function resetJob() {
  clearAllTriggers_(JOB_TRIGGER_FUNCTION);
  clearJobState_();
}
