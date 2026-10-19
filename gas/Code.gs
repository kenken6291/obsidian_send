/**
 * 健二保管庫（obsidian_send） — GAS バックエンド 完全版
 *
 * ■ スクリプトプロパティ（プロジェクトの設定 → スクリプト プロパティ）
 *   GEMINI_API_KEY  : Gemini APIキー（必須）
 *   GEMINI_MODEL    : Geminiモデル名（必須 例: gemini-2.5-flash）
 *   MEMO_FOLDER_ID  : メモ保存先フォルダID（省略時は 10_free_memo の既定ID）
 *   USERS_SHEET_ID  : 会員台帳スプレッドシートID（省略時は初回に自動作成して保存）
 *
 * ■ アクション（POST / body は JSON 文字列）
 *   register, login, forgotPassword, resetPassword, updateNickname,
 *   polishText, createMemo, listMemos, getMemo, updateMemo, deleteMemo
 */

const GAS_VERSION = '1.3.1';
const DEFAULT_MEMO_FOLDER_ID = '1k021VxuchCB_vITpVOtSM_RARCCJ9GvO';
const TOKEN_TTL_DAYS = 30;
const MAX_LOGIN_FAILS = 5;
const LOCK_MINUTES = 15;
const TZ = 'Asia/Tokyo';

const USER_HEADERS = [
  'email', 'nickname', 'salt', 'passwordHash', 'requiresReset',
  'token', 'tokenExpires', 'failCount', 'lockedUntil', 'createdAt', 'updatedAt'
];

// =====================================================================
// エントリポイント
// =====================================================================
function doGet() {
  return json_({ ok: true, app: '健二保管庫', version: GAS_VERSION });
}

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません。' });
  }

  try {
    switch (data.action) {
      // 会員
      case 'register':       return json_(register_(data));
      case 'login':          return json_(login_(data));
      case 'forgotPassword': return json_(forgotPassword_(data));
      case 'resetPassword':  return json_(withUser_(data, resetPassword_));
      case 'updateNickname': return json_(withUser_(data, updateNickname_));
      // メモ
      case 'polishText':     return json_(withUser_(data, polishText_));
      case 'createMemo':     return json_(withUser_(data, createMemo_));
      case 'listMemos':      return json_(withUser_(data, listMemos_));
      case 'getMemo':        return json_(withUser_(data, getMemo_));
      case 'updateMemo':     return json_(withUser_(data, updateMemo_));
      case 'deleteMemo':     return json_(withUser_(data, deleteMemo_));
      case 'version':        return json_({ ok: true, version: GAS_VERSION });
      default:
        return json_({ ok: false, error: '不明なアクションです: ' + data.action });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'サーバーエラー: ' + (err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// 会員台帳（スプレッドシート）
// =====================================================================
function getUserSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('USERS_SHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('健二保管庫_会員台帳');
    props.setProperty('USERS_SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName('users');
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName('users');
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function readUsers_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  return values.map(function (row, i) {
    const u = { _row: i + 2 };
    USER_HEADERS.forEach(function (h, j) { u[h] = row[j]; });
    return u;
  });
}

function findUserBy_(sh, key, value) {
  if (!value) return null;
  const v = String(value);
  const users = readUsers_(sh);
  for (let i = 0; i < users.length; i++) {
    if (String(users[i][key]) === v) return users[i];
  }
  return null;
}

function saveUser_(sh, u) {
  u.updatedAt = now_();
  const row = USER_HEADERS.map(function (h) { return u[h] === undefined ? '' : u[h]; });
  sh.getRange(u._row, 1, 1, USER_HEADERS.length).setValues([row]);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// トークン確認してから処理を実行
function withUser_(data, handler) {
  const sh = getUserSheet_();
  const user = findUserBy_(sh, 'token', data.token);
  if (!user || !user.tokenExpires || new Date(user.tokenExpires).getTime() < Date.now()) {
    return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  }
  return handler(data, user, sh);
}

// =====================================================================
// 会員機能
// =====================================================================
function register_(data) {
  const email = normalizeEmail_(data.email);
  const nickname = String(data.nickname || '').trim().slice(0, 30);
  if (!isValidEmail_(email)) return { ok: false, error: 'メールアドレスの形式が正しくありません。' };

  return withLock_(function () {
    const sh = getUserSheet_();
    if (findUserBy_(sh, 'email', email)) {
      return { ok: false, error: 'このメールアドレスは既に登録されています。パスワードがわからない場合は「パスワードをお忘れの方」から仮パスワードを受け取ってください。' };
    }
    const temp = makeTempPassword_();
    const salt = Utilities.getUuid();
    const row = {
      email: email, nickname: nickname || email.split('@')[0], salt: salt,
      passwordHash: hash_(temp, salt), requiresReset: true,
      token: '', tokenExpires: '', failCount: 0, lockedUntil: '',
      createdAt: now_(), updatedAt: now_()
    };

    // 先にメール送信（失敗したら登録しない）
    const mailErr = sendTempPasswordMail_(email, row.nickname, temp, '会員登録ありがとうございます。');
    if (mailErr) return { ok: false, error: mailErr };

    sh.appendRow(USER_HEADERS.map(function (h) { return row[h]; }));
    return { ok: true, message: email + ' 宛に仮パスワードを送信しました。メールの仮パスワードでログインしてください（届かない場合は迷惑メールフォルダもご確認ください）。' };
  });
}

function login_(data) {
  const email = normalizeEmail_(data.email);
  const password = String(data.password || '');
  if (!email || !password) return { ok: false, error: 'メールアドレスとパスワードを入力してください。' };

  return withLock_(function () {
    const sh = getUserSheet_();
    const u = findUserBy_(sh, 'email', email);
    if (!u) return { ok: false, error: 'メールアドレスまたはパスワードが違います。' };

    if (u.lockedUntil && new Date(u.lockedUntil).getTime() > Date.now()) {
      return { ok: false, error: 'ログインの失敗が続いたため、しばらくロックしています。' + LOCK_MINUTES + '分ほどおいて再度お試しください。' };
    }

    if (hash_(password, u.salt) !== String(u.passwordHash)) {
      u.failCount = Number(u.failCount || 0) + 1;
      if (u.failCount >= MAX_LOGIN_FAILS) {
        u.lockedUntil = new Date(Date.now() + LOCK_MINUTES * 60000).toISOString();
        u.failCount = 0;
      }
      saveUser_(sh, u);
      return { ok: false, error: 'メールアドレスまたはパスワードが違います。' };
    }

    u.failCount = 0;
    u.lockedUntil = '';
    u.token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    u.tokenExpires = new Date(Date.now() + TOKEN_TTL_DAYS * 86400000).toISOString();
    saveUser_(sh, u);
    return {
      ok: true, token: u.token, email: u.email, nickname: u.nickname,
      requiresPasswordReset: u.requiresReset === true || String(u.requiresReset).toUpperCase() === 'TRUE'
    };
  });
}

function forgotPassword_(data) {
  const email = normalizeEmail_(data.email);
  if (!isValidEmail_(email)) return { ok: false, error: 'メールアドレスの形式が正しくありません。' };

  return withLock_(function () {
    const sh = getUserSheet_();
    const u = findUserBy_(sh, 'email', email);
    if (!u) return { ok: false, error: 'このメールアドレスは登録されていません。会員登録をしてください。' };

    const temp = makeTempPassword_();
    const mailErr = sendTempPasswordMail_(email, u.nickname, temp, '仮パスワードを再発行しました。');
    if (mailErr) return { ok: false, error: mailErr };

    // メール送信に成功してからパスワードを置き換える
    u.salt = Utilities.getUuid();
    u.passwordHash = hash_(temp, u.salt);
    u.requiresReset = true;
    u.token = '';
    u.tokenExpires = '';
    u.failCount = 0;
    u.lockedUntil = '';
    saveUser_(sh, u);
    return { ok: true, message: email + ' 宛に仮パスワードを送信しました。メールをご確認ください（届かない場合は迷惑メールフォルダもご確認ください）。' };
  });
}

// 仮パスワードのメール送信。成功時は null、失敗時はエラーメッセージを返す
function sendTempPasswordMail_(email, nickname, temp, lead) {
  try {
    if (MailApp.getRemainingDailyQuota() < 1) {
      return '本日のメール送信上限に達しました。明日もう一度お試しください。';
    }
    MailApp.sendEmail({
      to: email,
      subject: '【健二保管庫】仮パスワードのお知らせ',
      name: '健二保管庫',
      body: [
        (nickname || '') + ' 様',
        '',
        lead,
        '',
        '仮パスワード： ' + temp,
        '',
        '下のページからログインし、新しいパスワードを設定してください。',
        'https://kenken6291.github.io/obsidian_send/',
        '',
        '※ このメールに心当たりがない場合は破棄してください。'
      ].join('\n')
    });
    return null;
  } catch (err) {
    console.error(err);
    return 'メールを送信できませんでした: ' + (err && err.message ? err.message : err);
  }
}

function resetPassword_(data, user, sh) {
  const newPassword = String(data.newPassword || '');
  if (newPassword.length < 6) return { ok: false, error: '新しいパスワードは6文字以上にしてください。' };
  return withLock_(function () {
    const u = findUserBy_(sh, 'email', user.email);
    u.salt = Utilities.getUuid();
    u.passwordHash = hash_(newPassword, u.salt);
    u.requiresReset = false;
    saveUser_(sh, u);
    return { ok: true };
  });
}

function updateNickname_(data, user, sh) {
  const nickname = String(data.nickname || '').trim().slice(0, 30);
  if (!nickname) return { ok: false, error: 'ニックネームを入力してください。' };
  return withLock_(function () {
    const u = findUserBy_(sh, 'email', user.email);
    u.nickname = nickname;
    saveUser_(sh, u);
    return { ok: true, nickname: nickname };
  });
}

// =====================================================================
// Gemini
// =====================================================================
function getGeminiConfig_() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = String(props.getProperty('GEMINI_MODEL') || '').trim();
  if (!apiKey) throw new Error('Gemini APIキーが未設定です（スクリプトプロパティ GEMINI_API_KEY）。');
  if (!model) throw new Error('Geminiモデルが未設定です（スクリプトプロパティ GEMINI_MODEL）。');
  return { apiKey: apiKey, model: model };
}

function callGemini_(parts, temperature) {
  const cfg = getGeminiConfig_();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg.model) + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': cfg.apiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: parts }],
      generationConfig: { temperature: temperature === undefined ? 0.3 : temperature }
    }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    let detail = '';
    try { detail = JSON.parse(res.getContentText()).error.message || ''; } catch (e) {}
    throw new Error('Gemini APIエラー (' + code + ', ' + cfg.model + ')' + (detail ? ': ' + detail : ''));
  }
  const body = JSON.parse(res.getContentText());
  const out = ((((body.candidates || [])[0] || {}).content || {}).parts || [])
    .map(function (p) { return p.text || ''; }).join('').trim();
  if (!out) throw new Error('Geminiから結果が返りませんでした。');
  return out;
}

function stripCodeFence_(s) {
  return String(s).replace(/^```[a-zA-Z]*\s*\n?/, '').replace(/\n?```\s*$/, '').trim();
}

// 音声認識テキストを整える（保存はしない）
function polishText_(data) {
  const text = String(data.text || '').trim();
  if (!text) return { ok: false, error: '整える文章がありません。' };
  if (text.length > 20000) return { ok: false, error: '文章が長すぎます。' };

  const prompt = [
    'あなたは音声入力メモの清書係です。以下は日本語の音声認識結果です。',
    '次のルールで読みやすく整えてください。',
    '・句読点を補い、誤変換と思われる語を文脈から正しく直す',
    '・「えー」「あの」「えっと」などの言いよどみ、言い直し、重複を取り除く',
    '・話題が変わる所や長い間（改行）の所で段落を分ける',
    '・内容の追加・要約・意見は禁止。話した内容と意味はそのまま残す',
    '・見出し、箇条書き、前置き、説明文は付けず、整えた本文だけを出力する',
    '',
    '--- 音声認識結果 ---',
    text
  ].join('\n');

  const out = stripCodeFence_(callGemini_([{ text: prompt }], 0.2));
  return { ok: true, text: out };
}

// =====================================================================
// メモ
// =====================================================================
function getMemoFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('MEMO_FOLDER_ID') || DEFAULT_MEMO_FOLDER_ID;
  return DriveApp.getFolderById(id);
}

function createMemo_(data, user) {
  const text = String(data.text || '').trim();
  const images = Array.isArray(data.images) ? data.images.slice(0, 10) : [];
  if (!text && images.length === 0) return { ok: false, error: '内容を入力するか、手書き・写真を追加してください。' };

  const now = new Date();
  const dateStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, TZ, 'HH:mm');

  const prompt = [
    'あなたはObsidian用のメモ整理係です。以下の入力（文章・手書き画像・写真）から、Obsidian用のMarkdownメモを1つ作成してください。',
    '',
    '出力形式（この形式のみを出力。コードブロックや説明は付けない）：',
    '---',
    'created: ' + dateStr + ' ' + timeStr,
    'tags: [内容に合うタグを3〜6個。日本語可・スペースなし]',
    'source: 健二保管庫',
    '---',
    '# 内容を表す短いタイトル',
    '',
    '本文',
    '',
    'ルール：',
    '・本文は入力の内容と意味をそのまま残し、読みやすく整える（要約・創作・意見の追加はしない）',
    '・文章の入力は利用者が確認済みの文章なので、言い回しは大きく変えない',
    '・手書き画像は文字を読み取って本文に含める。写真は写っている内容を簡潔に書き添える',
    '・関連しそうな語句は [[ ]] で囲んで内部リンクにしてよい（多用しない）',
    '',
    '--- 文章 ---',
    text || '（なし）'
  ].join('\n');

  const parts = [{ text: prompt }];
  images.forEach(function (img) {
    if (img && img.data) parts.push({ inline_data: { mime_type: img.mimeType || 'image/jpeg', data: img.data } });
  });

  let content;
  try {
    content = stripCodeFence_(callGemini_(parts, 0.3));
  } catch (err) {
    return { ok: false, error: err.message };
  }

  const folder = getMemoFolder_();
  const base = dateStr + '_' + Utilities.formatDate(now, TZ, 'HHmm') + '_メモ帳';
  let fileName = base + '.md';
  let n = 2;
  while (folder.getFilesByName(fileName).hasNext()) { fileName = base + '_' + n + '.md'; n++; }

  folder.createFile(fileName, content, 'text/markdown');
  return { ok: true, fileName: fileName, content: content };
}

function listMemos_(data) {
  const date = String(data.date || '').trim(); // yyyy-MM-dd
  const folder = getMemoFolder_();
  const it = folder.getFiles();
  const files = [];
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (!/\.md$/i.test(name)) continue;
    if (date && name.indexOf(date) !== 0) continue;
    files.push(f);
  }
  files.sort(function (a, b) {
    const an = a.getName(), bn = b.getName();
    if (an !== bn) return an < bn ? 1 : -1;
    return b.getLastUpdated() - a.getLastUpdated();
  });

  const memos = files.slice(0, 100).map(function (f, i) {
    let snippet = '';
    if (i < 50) {
      try { snippet = makeSnippet_(f.getBlob().getDataAsString('UTF-8')); } catch (e) {}
    }
    return { fileId: f.getId(), fileName: f.getName(), snippet: escapeHtml_(snippet) };
  });
  return { ok: true, memos: memos };
}

function getMemo_(data) {
  const f = getMemoFile_(data.fileId);
  if (!f) return { ok: false, error: 'メモが見つかりません。' };
  return { ok: true, fileName: f.getName(), content: f.getBlob().getDataAsString('UTF-8') };
}

function updateMemo_(data) {
  const f = getMemoFile_(data.fileId);
  if (!f) return { ok: false, error: 'メモが見つかりません。' };
  f.setContent(String(data.content || ''));
  return { ok: true };
}

function deleteMemo_(data) {
  const f = getMemoFile_(data.fileId);
  if (!f) return { ok: false, error: 'メモが見つかりません。' };
  f.setTrashed(true); // ゴミ箱へ移動（30日以内なら復元可）
  return { ok: true };
}

// 保存フォルダ内のファイルだけ操作できるようにする
function getMemoFile_(fileId) {
  if (!fileId) return null;
  let f;
  try { f = DriveApp.getFileById(String(fileId)); } catch (e) { return null; }
  if (f.isTrashed()) return null;
  const folderId = getMemoFolder_().getId();
  const parents = f.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return f;
  }
  return null;
}

function makeSnippet_(md) {
  const body = String(md).replace(/^---[\s\S]*?\n---\s*\n?/, '');
  const line = body.split('\n')
    .map(function (s) { return s.replace(/^#+\s*/, '').trim(); })
    .filter(function (s) { return s; })
    .join(' ');
  return line.length > 80 ? line.slice(0, 80) + '…' : line;
}

// =====================================================================
// ユーティリティ
// =====================================================================
function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function makeTempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 10; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function normalizeEmail_(email) { return String(email || '').trim().toLowerCase(); }
function isValidEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email); }
function now_() { return new Date().toISOString(); }
function escapeHtml_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

// =====================================================================
// 初回確認用（エディタから手動実行）
// =====================================================================
function setupCheck() {
  const sh = getUserSheet_();
  const folder = getMemoFolder_();
  const cfg = getGeminiConfig_();
  console.log('会員台帳: ' + sh.getParent().getUrl());
  console.log('保存フォルダ: ' + folder.getName());
  console.log('Geminiモデル: ' + cfg.model);
  console.log('バージョン: ' + GAS_VERSION);
  console.log('メール残り送信可能数: ' + MailApp.getRemainingDailyQuota());
}

// メール送信テスト（エディタから手動実行。自分宛に届けばOK）
function testMail() {
  const me = Session.getEffectiveUser().getEmail();
  const err = sendTempPasswordMail_(me, 'テスト', 'TEST1234', 'これはメール送信のテストです。');
  console.log(err ? 'NG: ' + err : 'OK: ' + me + ' に送信しました');
}
