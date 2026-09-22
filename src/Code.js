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
  RESULT_BY_QUESTION: '問題別結果',
  RESULT_FINAL: '最終結果',
};

const PROP_STATE = 'state';
const PROP_ANSWER_COUNTS = 'answerCounts';

// 出題操作から一斉表示までの猶予。全端末がポーリングで startAt を受け取る時間を確保する
const REVEAL_DELAY_MS = 8000;

const HEADER = {
  QUESTIONS: ['ID', '問題文', '備考', '種別', '選択肢1', '選択肢2', '選択肢3', '選択肢4', '正解', '制限時間(秒)', '最終出題時刻'],
  PARTICIPANTS: ['登録時刻', '名前'],
  ANSWERS: ['記録時刻', '問題ID', '名前', '回答', '回答タイムms', '正誤', '時間外', 'サーバー計測ms(参考)'],
  STATE: ['項目', '値'],
};

// 問題シートの列番号（1始まり）
const COL_QUESTION_TYPE = 4;
const COL_LAST_STARTED_AT = 11;

/* ========== エントリポイント ========== */

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  const file = page === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle('早押しクイズ大会')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 画面ファイルが共通部分を取り込むために呼ぶ */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ========== 状態管理 ========== */

function defaultState_() {
  return { rev: 0, phase: 'waiting', startAt: 0, question: null, correct: null };
}

function parseState_(raw) {
  return raw ? JSON.parse(raw) : defaultState_();
}

function rawState_() {
  return parseState_(PropertiesService.getScriptProperties().getProperty(PROP_STATE));
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
  const props = PropertiesService.getScriptProperties().getProperties();
  const state = parseState_(props[PROP_STATE]);
  if (state.phase !== 'answer') {
    state.correct = null;
  }
  // 選択肢ごとの回答数は締切後だけ渡す。回答受付中は値そのものを配らない
  state.counts = shouldRevealCounts_(state) ? readAnswerCounts_(props, state.question.id) : null;
  state.serverNow = Date.now();
  return state;
}

function shouldRevealCounts_(state) {
  if (!state.question) return false;
  if (state.phase === 'answer') return true;
  if (state.phase !== 'question') return false;
  return Date.now() > state.startAt + state.question.limitSec * 1000;
}

function readAnswerCounts_(props, questionId) {
  const raw = props[PROP_ANSWER_COUNTS];
  if (!raw) return null;
  const stored = JSON.parse(raw);
  return stored.questionId === questionId ? stored.counts : null;
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
      ensureAnswerCounts_(q.id);
    }
  }
  saveState_(state);
  return state;
}

/* ========== 管理画面から呼ばれる操作 ========== */

function adminStartQuestion(questionId) {
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
  resetAnswerCounts_(q.id);

  // 記録用に出題時刻を残す
  sheet_(SHEET.QUESTIONS).getRange(q.rowIndex, COL_LAST_STARTED_AT).setValue(new Date());
  return adminGetDashboard();
}

function adminRevealAnswer() {
  const state = rawState_();
  if (!state.question) throw new Error('出題中の問題がありません');
  state.rev += 1;
  state.phase = 'answer';
  saveState_(state);
  return adminGetDashboard();
}

function adminSetWaiting() {
  const prev = rawState_();
  const state = defaultState_();
  state.rev = prev.rev + 1;
  saveState_(state);
  return adminGetDashboard();
}

/** 管理画面のポーリング用。参加者画面と違いスプレッドシートを読むが、接続は管理者1台だけ */
function adminGetDashboard() {
  // 状態と回答数を1回の読み出しでまとめて取るため、rawState_ は経由しない
  const props = PropertiesService.getScriptProperties().getProperties();
  const state = parseState_(props[PROP_STATE]);
  return {
    state: state, // 管理画面には正解も渡す
    serverNow: Date.now(),
    questions: loadQuestions_().map(toAdminQuestion_),
    participantCount: Math.max(0, sheet_(SHEET.PARTICIPANTS).getLastRow() - 1),
    answerCount: state.question ? totalAnswerCount_(props, state.question.id) : 0,
  };
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
    incrementAnswerCount_(question.id, choice);
  } finally {
    lock.releaseLock();
  }

  // 正誤は返さない（発表前に正解が漏れるため）
  return { ok: true };
}

/* ========== 選択肢ごとの回答数 ========== */

/**
 * 回答数はシートを数え直さず、カウンタとして積む。
 * ポーリングで呼ばれる getState からスプレッドシートを読まないようにするため。
 * 呼び出し元の submitAnswer が LockService で直列化しているので、加算は競合しない。
 */
function incrementAnswerCount_(questionId, choice) {
  if (choice < 1 || choice > 4) return;

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_ANSWER_COUNTS);
  if (!raw) return;

  const stored = JSON.parse(raw);
  if (stored.questionId !== questionId) return;

  stored.counts[choice - 1] += 1;
  props.setProperty(PROP_ANSWER_COUNTS, JSON.stringify(stored));
}

/** 出題中の問題に集まった回答数。選択肢ごとのカウンタの合計で足りる */
function totalAnswerCount_(props, questionId) {
  const counts = readAnswerCounts_(props, questionId);
  if (!counts) return 0;
  return counts.reduce(function (sum, count) { return sum + count; }, 0);
}

/**
 * 回答を積める状態にする。
 * カウンタが無いまま回答が来ると incrementAnswerCount_ が何もせず、
 * 管理画面の回答数も参加者画面の選択肢別回答数も出なくなる。
 * 同じ問題のカウンタが既にあるときは、積んだ数を失わないよう温存する。
 */
function ensureAnswerCounts_(questionId) {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_ANSWER_COUNTS);
  const stored = raw ? JSON.parse(raw) : null;
  if (stored && stored.questionId === questionId) return;
  resetAnswerCounts_(questionId);
}

function resetAnswerCounts_(questionId) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_ANSWER_COUNTS, JSON.stringify({ questionId: questionId, counts: [0, 0, 0, 0] }));
}

/* ========== 問題の読み込み ========== */

function loadQuestions_() {
  const sh = sheet_(SHEET.QUESTIONS);
  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, HEADER.QUESTIONS.length).getValues()
    .map(function (row, i) {
      const type = String(row[3]).trim().toLowerCase() === 'image' ? 'image' : 'text';
      const choices = [row[4], row[5], row[6], row[7]].map(String);
      return {
        rowIndex: i + 2,
        id: String(row[0]).trim(),
        text: String(row[1]),
        // 参加者には渡さず、管理画面にだけ表示する（正解の解説を書く想定のため）
        note: String(row[2]),
        type: type,
        // 画像の選択肢だけ配列になる（1つの選択肢に複数枚を指定できるため）
        choices: type === 'image' ? choices.map(splitImageUrls_) : choices,
        correct: Number(row[8]),
        limitSec: Number(row[9]) || 20,
      };
    })
    .filter(function (q) { return q.id; });
}

/**
 * 画像の選択肢は | 区切りで複数枚を指定できる。
 * URLに現れない文字のため、カンマと違って値の一部と取り違える余地がない。
 */
function splitImageUrls_(value) {
  return String(value)
    .split('|')
    .map(function (url) { return url.trim(); })
    .filter(Boolean);
}

function findQuestion_(questionId) {
  const id = String(questionId).trim();
  const matched = loadQuestions_().filter(function (q) { return q.id === id; });
  return matched.length ? matched[0] : null;
}

/** シート由来の行オブジェクトから、参加者画面に渡す形へ変換する（rowIndex・correct・note を落とす） */
function toClientQuestion_(q) {
  return { id: q.id, text: q.text, type: q.type, choices: q.choices, limitSec: q.limitSec };
}

/** 管理画面に渡す形。備考を含める */
function toAdminQuestion_(q) {
  const question = toClientQuestion_(q);
  question.note = q.note;
  return question;
}

/* ========== 初期セットアップ ========== */

function setupSheets() {
  const ss = book_();
  applyQuestionTypeValidation_(ensureSheet_(ss, SHEET.QUESTIONS, HEADER.QUESTIONS));
  ensureSheet_(ss, SHEET.PARTICIPANTS, HEADER.PARTICIPANTS);
  ensureSheet_(ss, SHEET.ANSWERS, HEADER.ANSWERS);
  ensureSheet_(ss, SHEET.STATE, HEADER.STATE);
  setupResultByQuestionSheet_(ss);
  setupFinalResultSheet_(ss);
  seedSampleQuestions_();

  // 新規スプレッドシート作成時の空シートを片付ける
  const blank = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  saveState_(defaultState_());
}

function ensureSheet_(ss, name, header) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/**
 * 問題別結果シートを組み立てる。
 * 集計はすべてシートの関数で行い、回答シートは生データとして触らない。
 * 中身は生成物なので、実行のたびに全体を書き直して問題ない。
 */
function setupResultByQuestionSheet_(ss) {
  const sh = ss.getSheetByName(SHEET.RESULT_BY_QUESTION) || ss.insertSheet(SHEET.RESULT_BY_QUESTION);

  sh.getRange('A1:A2').setValues([['問題ID'], ['問題文']]).setFontWeight('bold');
  sh.getRange('B2').setFormula("=IFERROR(VLOOKUP($B$1,'問題'!$A:$B,2,FALSE),\"\")");

  sh.getRange('A4:D4').setValues([['順位', '名前', '回答タイムms', 'タイム(秒)']]).setFontWeight('bold');
  sh.getRange('A5').setFormula('=ARRAYFORMULA(IF(LEN($B$5:$B),ROW($B$5:$B)-4,""))');
  // 正解かつ時間内の回答だけを、回答タイムの昇順で並べる
  sh.getRange('B5').setFormula(
    "=IFERROR(SORT(FILTER({'回答'!$C$2:$C,'回答'!$E$2:$E}," +
    "'回答'!$B$2:$B=$B$1,'回答'!$F$2:$F=1,'回答'!$G$2:$G=0),2,TRUE),\"\")"
  );
  sh.getRange('D5').setFormula('=ARRAYFORMULA(IF(LEN($C$5:$C),$C$5:$C/1000,""))');

  // 問題IDは問題シートから選ぶ
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET.QUESTIONS).getRange('A2:A1000'), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B1').setDataValidation(rule);

  sh.setFrozenRows(4);
  return sh;
}

/**
 * 最終結果シートを組み立てる。
 * 正解数の多い順に並べ、同数なら正解した問題の合計タイムが短い方を上位とする。
 * 時間外の回答は正解数・タイムのどちらからも除外し、問題別結果シートと条件を揃える。
 *
 * 集計対象の名前は、参加者シートと回答シートの両方から集める。
 * 参加者シートだけを見ると、名前が一致しない回答が黙って0正解に沈んで気づけない。
 * 前後の空白は落としたうえで突き合わせる。
 *
 * 集計はG列より右の作業用セルに一度出してから並べ替える。
 * 検索条件に計算結果の配列を渡すと展開されないため、条件には実際のセル範囲を渡す。
 * 合計タイムは SUMIFS が展開されず0になったため、回答1行ごとの有効タイムを
 * K列に出したうえで SUMIF で合計している。
 */
function setupFinalResultSheet_(ss) {
  const sh = ss.getSheetByName(SHEET.RESULT_FINAL) || ss.insertSheet(SHEET.RESULT_FINAL);

  sh.getRange('A1:E1')
    .setValues([['順位', '名前', '正解数', '合計タイムms', '合計タイム(秒)']])
    .setFontWeight('bold');
  sh.getRange('G1:L1')
    .setValues([[
      '作業用: 名前', '作業用: 正解数', '作業用: 合計タイムms',
      '作業用: 回答の名前', '作業用: 有効タイムms', '作業用: 参加者の名前',
    ]])
    .setFontWeight('bold');

  // 参加者シートと回答シートの両方に出てくる名前を集める
  sh.getRange('G2').setFormula(
    "=IFERROR(UNIQUE(FILTER({$L$2:$L$500;$J$2:$J$2000}," +
    "{$L$2:$L$500;$J$2:$J$2000}<>\"\")),\"\")"
  );
  sh.getRange('H2').setFormula(
    "=ARRAYFORMULA(IF(LEN($G$2:$G$200)," +
    "COUNTIFS($J$2:$J,$G$2:$G$200,$K$2:$K,\">0\"),\"\"))"
  );
  sh.getRange('I2').setFormula(
    "=ARRAYFORMULA(IF(LEN($G$2:$G$200),SUMIF($J$2:$J,$G$2:$G$200,$K$2:$K),\"\"))"
  );

  // 回答1行ごとの値。正解かつ時間内のときだけタイムが入る
  sh.getRange('J2').setFormula(
    "=ARRAYFORMULA(IF(LEN('回答'!$C$2:$C),TRIM('回答'!$C$2:$C),\"\"))"
  );
  sh.getRange('K2').setFormula(
    "=ARRAYFORMULA(IF(LEN('回答'!$C$2:$C)," +
    "'回答'!$E$2:$E*('回答'!$F$2:$F=1)*('回答'!$G$2:$G=0),\"\"))"
  );
  sh.getRange('L2').setFormula(
    "=ARRAYFORMULA(IF(LEN('参加者'!$B$2:$B),TRIM('参加者'!$B$2:$B),\"\"))"
  );

  sh.getRange('A2').setFormula('=ARRAYFORMULA(IF(LEN($B$2:$B),ROW($B$2:$B)-1,""))');
  sh.getRange('B2').setFormula(
    "=IFERROR(SORT(FILTER({$G$2:$G$200,$H$2:$H$200,$I$2:$I$200}," +
    "LEN($G$2:$G$200)),2,FALSE,3,TRUE),\"\")"
  );
  sh.getRange('E2').setFormula('=ARRAYFORMULA(IF(LEN($D$2:$D),$D$2:$D/1000,""))');

  sh.setFrozenRows(1);
  return sh;
}

/** 種別をプルダウンにする。typoしても text にフォールバックしてしまい、気づきにくいため */
function applyQuestionTypeValidation_(sh) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['text', 'image'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, COL_QUESTION_TYPE, 999, 1).setDataValidation(rule);
}

function seedSampleQuestions_() {
  const sh = sheet_(SHEET.QUESTIONS);
  if (sh.getLastRow() > 1) return; // 既に問題があれば触らない

  const rows = [
    ['q01', 'ミュンヘンのオクトーバーフェストが初めて開催された年は？',
      'バイエルン王太子ルートヴィヒとテレーゼの結婚を祝う祭りが起源', 'text',
      '1810年', '1850年', '1900年', '1946年', 1, 20, ''],
    ['q02', 'オクトーバーフェストの会場となっている広場の名前は？',
      '王太子妃テレーゼの名前に由来する', 'text',
      'マリエンプラッツ', 'テレージエンヴィーゼ', 'オデオンスプラッツ', 'カールスプラッツ', 2, 20, ''],
    ['q03', 'オクトーバーフェストで提供されるビールジョッキ「マース」の容量は？',
      '', 'text',
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
