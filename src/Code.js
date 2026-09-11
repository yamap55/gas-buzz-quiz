/**
 * 早押し4択クイズ アプリ本体
 *
 * 設計上の要点:
 * - 進行状態は ScriptProperties を真実源とし、ポーリングで呼ばれる getState() はスプレッドシートに触らない
 *   （スプレッドシート読み取りは1回1〜2秒かかり、同時実行数の上限30に近づくため）
 * - 問題の表示は絶対時刻 startAt で全端末を同期させ、経過時間はクライアント計測値を採用する
 *   （ポーリング間隔と通信遅延を順位から排除するため）
 */

const SHEET = {
  QUESTIONS: '問題',
  PARTICIPANTS: '参加者',
  ANSWERS: '回答',
  STATE: '状態',
};

const PROP_STATE = 'state';
const PROP_ADMIN_KEY = 'adminKey';

// 出題操作から一斉表示までの猶予。全端末がポーリングで startAt を受け取る時間を確保する
const REVEAL_DELAY_MS = 8000;

const HEADER = {
  QUESTIONS: ['ID', '問題文', '種別', '選択肢1', '選択肢2', '選択肢3', '選択肢4', '正解', '制限時間(秒)', '最終出題時刻'],
  PARTICIPANTS: ['登録時刻', '名前'],
  ANSWERS: ['記録時刻', '問題ID', '名前', '回答', '経過時間ms', '正誤', '時間外', '経過時間ms(サーバー)'],
  STATE: ['項目', '値'],
};

// 問題シートの列番号（1始まり）
const COL_LAST_STARTED_AT = 10;

/* ========== エントリポイント ========== */

function doGet(e) {
  const params = (e && e.parameter) || {};
  const output = isAdminRequest_(params)
    ? adminPage_()
    : HtmlService.createHtmlOutputFromFile('Index');
  return output
    .setTitle('早押しクイズ大会')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 管理画面には合言葉を埋め込む。google.script.run の呼び出しに添えるため */
function adminPage_() {
  const template = HtmlService.createTemplateFromFile('Admin');
  template.adminKey = adminKey_();
  return template.evaluate();
}

/**
 * ウェブアプリ全体の公開設定は、参加者がGoogleログインなしで参加するために
 * ANYONE_ANONYMOUS にする必要がある。そのため管理画面は合言葉で切り分ける。
 * 一致しない場合は参加者画面を返し、管理画面の存在自体を示さない。
 */
function isAdminRequest_(params) {
  if (params.page !== 'admin') return false;
  const expected = PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_KEY);
  return Boolean(expected) && params.key === expected;
}

/** 合言葉を取得する。未設定なら生成して保存する */
function adminKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(PROP_ADMIN_KEY);
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty(PROP_ADMIN_KEY, key);
  }
  return key;
}

/**
 * 管理操作の実行を許可してよいか検証する。
 * 参加者画面からでも google.script.run で関数を呼べてしまうため、
 * 画面の出し分けだけでなくサーバー側の操作にも合言葉を要求する。
 */
function requireAdmin_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_KEY);
  if (!expected || key !== expected) throw new Error('権限がありません');
}

/** 管理画面のURLを表示する（エディタから実行し、実行ログで確認する） */
function showAdminUrl() {
  const url = ScriptApp.getService().getUrl() + '?page=admin&key=' + adminKey_();
  Logger.log(url);
  return url;
}

/* ========== 状態管理 ========== */

function defaultState_() {
  return { rev: 0, phase: 'waiting', startAt: 0, question: null, correct: null };
}

function rawState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_STATE);
  return raw ? JSON.parse(raw) : defaultState_();
}

function saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(PROP_STATE, JSON.stringify(state));
  writeStateMirror_(state);
}

/**
 * 参加者画面のポーリングから呼ばれる。スプレッドシートには触らない。
 * 正解は発表フェーズになるまでクライアントへ渡さない。
 */
function getState() {
  const state = rawState_();
  if (state.phase !== 'answer') {
    state.correct = null;
  }
  state.serverNow = Date.now();
  return state;
}

/** 人間が読む・手で直すための状態ミラーをシートに書く */
function writeStateMirror_(state) {
  const sh = book_().getSheetByName(SHEET.STATE);
  if (!sh) return; // setupSheets 実行前は何もしない
  const rows = [
    ['rev', state.rev],
    ['phase', state.phase],
    ['問題ID', state.question ? state.question.id : ''],
    ['startAt(epoch ms)', state.startAt || ''],
    ['startAt(表示用)', state.startAt ? formatTime_(new Date(state.startAt)) : ''],
    ['更新時刻', formatTime_(new Date())],
  ];
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/** 状態シートを手で直したあと、エディタから実行して ScriptProperties を作り直す（リカバリ用） */
function restoreStateFromSheet() {
  const sh = sheet_(SHEET.STATE);
  const map = {};
  sh.getRange(2, 1, 6, 2).getValues().forEach(function (row) {
    map[row[0]] = row[1];
  });

  const state = defaultState_();
  state.rev = Number(map['rev'] || 0) + 1;
  state.phase = String(map['phase'] || 'waiting');
  state.startAt = Number(map['startAt(epoch ms)'] || 0);

  const questionId = String(map['問題ID'] || '').trim();
  if (questionId) {
    const q = findQuestion_(questionId);
    if (q) {
      state.question = toClientQuestion_(q);
      state.correct = q.correct;
    }
  }
  saveState_(state);
  return state;
}

/* ========== 管理画面から呼ばれる操作 ========== */

function adminStartQuestion(key, questionId) {
  requireAdmin_(key);

  const q = findQuestion_(questionId);
  if (!q) throw new Error('問題が見つかりません: ' + questionId);

  const prev = rawState_();
  const state = {
    rev: prev.rev + 1,
    phase: 'question',
    startAt: Date.now() + REVEAL_DELAY_MS,
    question: toClientQuestion_(q),
    correct: q.correct,
  };
  saveState_(state);

  // 記録用に出題時刻を残す
  sheet_(SHEET.QUESTIONS).getRange(q.rowIndex, COL_LAST_STARTED_AT).setValue(new Date());
  return dashboard_();
}

function adminRevealAnswer(key) {
  requireAdmin_(key);

  const state = rawState_();
  if (!state.question) throw new Error('出題中の問題がありません');
  state.rev += 1;
  state.phase = 'answer';
  saveState_(state);
  return dashboard_();
}

function adminSetWaiting(key) {
  requireAdmin_(key);

  const prev = rawState_();
  const state = defaultState_();
  state.rev = prev.rev + 1;
  saveState_(state);
  return dashboard_();
}

/** 管理画面のポーリング用。参加者画面と違いスプレッドシートを読むが、接続は管理者1台だけ */
function adminGetDashboard(key) {
  requireAdmin_(key);
  return dashboard_();
}

function dashboard_() {
  const state = rawState_();
  return {
    state: state, // 管理画面には正解も渡す
    serverNow: Date.now(),
    questions: loadQuestions_().map(toClientQuestion_),
    participantCount: Math.max(0, sheet_(SHEET.PARTICIPANTS).getLastRow() - 1),
    answerCount: state.question ? countAnswers_(state.question.id) : 0,
  };
}

function countAnswers_(questionId) {
  const sh = sheet_(SHEET.ANSWERS);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  return sh.getRange(2, 2, last - 1, 1).getValues().filter(function (row) {
    return row[0] === questionId;
  }).length;
}

/* ========== 参加者画面から呼ばれる操作 ========== */

function registerParticipant(name) {
  const trimmed = String(name || '').trim();
  if (!trimmed) throw new Error('名前を入力してください');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet_(SHEET.PARTICIPANTS).appendRow([new Date(), trimmed]);
  } finally {
    lock.releaseLock();
  }
  return { name: trimmed };
}

/**
 * 回答を記録する。
 * 順位に使うのは payload.elapsedMs（クライアント計測値）で、サーバー側の経過時間は参考値として残す。
 * 二重送信はクライアント側で防ぐ想定のため、ここでは弾かない（集計時に最初の1件を採用する）。
 */
function submitAnswer(payload) {
  const state = rawState_();
  const question = state.question;
  if (state.phase !== 'question' || !question || question.id !== payload.questionId) {
    return { ok: false, reason: 'closed' };
  }

  const choice = Number(payload.choice);
  const elapsedMs = Math.round(Number(payload.elapsedMs));
  const serverElapsedMs = Date.now() - state.startAt;
  const isCorrect = choice === Number(state.correct);
  const isLate = elapsedMs > question.limitSec * 1000;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet_(SHEET.ANSWERS).appendRow([
      new Date(),
      question.id,
      String(payload.name),
      choice,
      elapsedMs,
      isCorrect ? 1 : 0,
      isLate ? 1 : 0,
      Math.round(serverElapsedMs),
    ]);
  } finally {
    lock.releaseLock();
  }

  // 正誤は返さない（発表前に正解が漏れるため）
  return { ok: true };
}

/* ========== 問題の読み込み ========== */

function loadQuestions_() {
  const sh = sheet_(SHEET.QUESTIONS);
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, HEADER.QUESTIONS.length).getValues()
    .map(function (row, i) {
      return {
        rowIndex: i + 2,
        id: String(row[0]).trim(),
        text: String(row[1]),
        type: String(row[2]).trim().toLowerCase() === 'image' ? 'image' : 'text',
        choices: [row[3], row[4], row[5], row[6]].map(String),
        correct: Number(row[7]),
        limitSec: Number(row[8]) || 20,
      };
    })
    .filter(function (q) { return q.id; });
}

function findQuestion_(questionId) {
  const id = String(questionId).trim();
  const matched = loadQuestions_().filter(function (q) { return q.id === id; });
  return matched.length ? matched[0] : null;
}

/** シート由来の行オブジェクトから、画面に渡す形へ変換する（rowIndex と correct を落とす） */
function toClientQuestion_(q) {
  return { id: q.id, text: q.text, type: q.type, choices: q.choices, limitSec: q.limitSec };
}

/* ========== 初期セットアップ ========== */

function setupSheets() {
  const ss = book_();
  ensureSheet_(ss, SHEET.QUESTIONS, HEADER.QUESTIONS);
  ensureSheet_(ss, SHEET.PARTICIPANTS, HEADER.PARTICIPANTS);
  ensureSheet_(ss, SHEET.ANSWERS, HEADER.ANSWERS);
  ensureSheet_(ss, SHEET.STATE, HEADER.STATE);
  seedSampleQuestions_();

  // 新規スプレッドシート作成時の空シートを片付ける
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  saveState_(defaultState_());
  showAdminUrl();
}

function ensureSheet_(ss, name, header) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function seedSampleQuestions_() {
  const sh = sheet_(SHEET.QUESTIONS);
  if (sh.getLastRow() > 1) return; // 既に問題があれば触らない

  const rows = [
    ['q01', 'ミュンヘンのオクトーバーフェストが初めて開催された年は？', 'text',
      '1810年', '1850年', '1900年', '1946年', 1, 20, ''],
    ['q02', 'オクトーバーフェストの会場となっている広場の名前は？', 'text',
      'マリエンプラッツ', 'テレージエンヴィーゼ', 'オデオンスプラッツ', 'カールスプラッツ', 2, 20, ''],
    ['q03', 'オクトーバーフェストで提供されるビールジョッキ「マース」の容量は？', 'text',
      '0.5リットル', '0.75リットル', '1リットル', '1.5リットル', 3, 20, ''],
  ];
  sh.getRange(2, 1, rows.length, HEADER.QUESTIONS.length).setValues(rows);
}

/* ========== 小物 ========== */

function book_() {
  return SpreadsheetApp.getActive();
}

function sheet_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name + '（setupSheets を実行してください）');
  return sh;
}

function formatTime_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
