/**
 * 健二保管庫 - Obsidian用フリーメモWebアプリ バックエンド (GAS)
 *
 * 機能:
 *  - 会員認証（登録・ログイン・パスワード忘れ→仮パスワード発行・強制再登録・ニックネーム編集）
 *  - テキスト/音声/手書き/写真からの入力を受け取り、Gemini APIで整形してObsidian用Markdownメモを生成
 *  - 指定Googleドライブフォルダに「年-月-日_時分_メモ帳.md」形式で新規ファイルを作成
 *  - 作成済みメモの一覧表示（日付指定フィルタ可）・再表示・編集・削除
 *
 * スクリプトプロパティ:
 *   GEMINI_API_KEY : Gemini APIキー
 *   MEMO_FOLDER_ID : 保存先フォルダID（"10_free_memo"）
 *   USER_SHEET_ID  : 会員管理スプレッドシートID（未設定なら初回登録時に自動作成）
 *   PEPPER         : パスワードハッシュ用の秘密文字列（未設定なら自動生成して保存）
 */

var GEMINI_MODEL = 'gemini-3.6-flash';
var SESSION_TTL_SEC = 21600; // 6時間
var SHEET_NAME = 'Users';

// ============ エントリポイント ============

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    switch (action) {
      case 'register':        result = handleRegister_(body); break;
      case 'login':            result = handleLogin_(body); break;
      case 'forgotPassword':   result = handleForgotPassword_(body); break;
      case 'resetPassword':    result = handleResetPassword_(body); break;
      case 'updateNickname':   result = handleUpdateNickname_(body); break;
      case 'createMemo':       result = handleCreateMemo_(body); break;
      case 'listMemos':        result = handleListMemos_(body); break;
      case 'getMemo':          result = handleGetMemo_(body); break;
      case 'updateMemo':       result = handleUpdateMemo_(body); break;
      case 'deleteMemo':       result = handleDeleteMemo_(body); break;
      default:
        result = { ok: false, error: '不明なアクションです: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: 'サーバーエラー: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('健二保管庫 API is running.');
}

// ============ 会員管理まわり ============

function getUserSheet_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('USER_SHEET_ID');
  var ss;
  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.create('ObsidianSend_Users');
    props.setProperty('USER_SHEET_ID', ss.getId());
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Email', 'PasswordHash', 'Salt', 'Nickname', 'IsTempPassword']);
    var defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  }
  return sheet;
}

function getPepper_() {
  var props = PropertiesService.getScriptProperties();
  var pepper = props.getProperty('PEPPER');
  if (!pepper) {
    pepper = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PEPPER', pepper);
  }
  return pepper;
}

function hashPassword_(password, salt) {
  var pepper = getPepper_();
  var raw = Utilities.computeHmacSha256Signature(password + ':' + salt, pepper);
  return Utilities.base64Encode(raw);
}

function findUserRowIndex_(sheet, email) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) return i + 1; // 1-based row
  }
  return -1;
}

function handleRegister_(body) {
  var email = (body.email || '').trim();
  var password = body.password || '';
  var nickname = (body.nickname || '').trim();
  if (!email || password.length < 6) {
    return { ok: false, error: 'メールアドレスと6文字以上のパスワードを入力してください。' };
  }
  var sheet = getUserSheet_();
  if (findUserRowIndex_(sheet, email) !== -1) {
    return { ok: false, error: 'このメールアドレスは既に登録されています。' };
  }
  if (!nickname) nickname = email.split('@')[0];
  var salt = Utilities.getUuid();
  var hash = hashPassword_(password, salt);
  sheet.appendRow([email, hash, salt, nickname, 'FALSE']);
  return { ok: true, message: '登録が完了しました。ログインしてください。' };
}

function handleLogin_(body) {
  var email = (body.email || '').trim();
  var password = body.password || '';
  var sheet = getUserSheet_();
  var row = findUserRowIndex_(sheet, email);
  if (row === -1) return { ok: false, error: 'メールアドレスまたはパスワードが違います。' };

  var values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  var storedHash = values[1], salt = values[2], nickname = values[3], isTemp = String(values[4]) === 'TRUE';
  var hash = hashPassword_(password, salt);
  if (hash !== storedHash) return { ok: false, error: 'メールアドレスまたはパスワードが違います。' };

  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, email, SESSION_TTL_SEC);
  return { ok: true, token: token, email: email, nickname: nickname, requiresPasswordReset: isTemp };
}

function getSessionEmail_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var email = cache.get('session_' + token);
  if (email) cache.put('session_' + token, email, SESSION_TTL_SEC); // 延長
  return email;
}

function handleForgotPassword_(body) {
  var email = (body.email || '').trim();
  var sheet = getUserSheet_();
  var row = findUserRowIndex_(sheet, email);
  // メールアドレスの存在有無を外部から推測させないため、常に同じ成功メッセージを返す
  if (row !== -1) {
    var tempPassword = Utilities.getUuid().replace(/-/g, '').substring(0, 8);
    var salt = Utilities.getUuid();
    var hash = hashPassword_(tempPassword, salt);
    sheet.getRange(row, 2, 1, 3).setValues([[hash, salt, sheet.getRange(row, 4).getValue()]]);
    sheet.getRange(row, 5).setValue('TRUE');
    try {
      MailApp.sendEmail(email, '【健二保管庫】仮パスワードのお知らせ',
        '仮パスワードを発行しました。\n\n仮パスワード: ' + tempPassword +
        '\n\nログイン後、必ず新しいパスワードを再登録してください。');
    } catch (e) {
      // メール送信失敗時もユーザーには成功と同じ応答を返す（仮パスワードは発行済みのため）
    }
  }
  return { ok: true, message: '登録済みのメールアドレスであれば、仮パスワードを送信しました。' };
}

function handleResetPassword_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  var newPassword = body.newPassword || '';
  if (newPassword.length < 6) return { ok: false, error: '新しいパスワードは6文字以上にしてください。' };

  var sheet = getUserSheet_();
  var row = findUserRowIndex_(sheet, email);
  if (row === -1) return { ok: false, error: 'ユーザーが見つかりません。' };
  var salt = Utilities.getUuid();
  var hash = hashPassword_(newPassword, salt);
  sheet.getRange(row, 2, 1, 2).setValues([[hash, salt]]);
  sheet.getRange(row, 5).setValue('FALSE');
  return { ok: true };
}

function handleUpdateNickname_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  var nickname = (body.nickname || '').trim();
  if (!nickname) return { ok: false, error: 'ニックネームを入力してください。' };
  var sheet = getUserSheet_();
  var row = findUserRowIndex_(sheet, email);
  if (row === -1) return { ok: false, error: 'ユーザーが見つかりません。' };
  sheet.getRange(row, 4).setValue(nickname);
  return { ok: true, nickname: nickname };
}

// ============ メモ本体（作成・一覧・取得・更新・削除） ============

function getMemoFolder_() {
  var folderId = PropertiesService.getScriptProperties().getProperty('MEMO_FOLDER_ID');
  if (!folderId) throw new Error('MEMO_FOLDER_ID が設定されていません。');
  return DriveApp.getFolderById(folderId);
}

// ファイル名: 年-月-日_時分_メモ帳 （重複時は _2, _3 ... を付与）
function buildMemoFileName_(folder, date) {
  var tz = Session.getScriptTimeZone();
  var base = Utilities.formatDate(date, tz, 'yyyy-MM-dd_HHmm') + '_メモ帳';
  var name = base;
  var suffix = 2;
  while (folder.getFilesByName(name + '.md').hasNext()) {
    name = base + '_' + suffix;
    suffix++;
  }
  return name + '.md';
}

function callGemini_(promptText, images) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  var parts = [{ text: promptText }];
  (images || []).forEach(function (img) {
    parts.push({ inlineData: { mimeType: img.mimeType, data: img.data } });
  });

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  var payload = { contents: [{ parts: parts }] };
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var json = JSON.parse(res.getContentText());
  if (json.error) throw new Error('Gemini APIエラー: ' + json.error.message);
  var text = json.candidates && json.candidates[0] && json.candidates[0].content &&
      json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
      json.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Geminiから有効な応答がありませんでした。');
  return text.trim();
}

function handleCreateMemo_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };

  var text = (body.text || '').trim();
  var images = body.images || []; // [{mimeType, data(base64)}]
  if (!text && images.length === 0) {
    return { ok: false, error: '内容を入力するか、手書き・写真を追加してください。' };
  }

  var prompt =
    '以下はObsidianに保存するフリーメモの元になる入力です。テキストと（あれば）画像の内容をもとに、' +
    '読みやすいObsidian用Markdownメモを1つ作成してください。\n' +
    '- 画像に手書き文字や印字された文字が写っている場合は、それを読み取ってメモ本文に含めてください。\n' +
    '- 写真（風景・物・人など）の場合は、内容を簡潔に描写して本文に含めてください。\n' +
    '- 内容から適切なタイトルを先頭に「# 」見出しで付けてください。\n' +
    '- 内容に応じたタグを3つ程度、本文末尾に「#タグ名」の形式で付けてください。\n' +
    '- 前置きや説明文は不要です。メモ本文のみをMarkdownで出力してください。\n\n' +
    '--- 入力テキスト ---\n' + (text || '（テキスト入力なし。画像を参照）');

  var formatted;
  try {
    formatted = callGemini_(prompt, images);
  } catch (err) {
    // Gemini整形に失敗した場合も、入力内容は失わずそのまま保存する
    formatted = '# メモ\n\n' + (text || '（画像のみの入力・整形失敗のため画像内容は含まれていません）') +
      '\n\n> 整形失敗: ' + err.message;
  }

  var now = new Date();
  var folder = getMemoFolder_();
  var fileName = buildMemoFileName_(folder, now);
  var tz = Session.getScriptTimeZone();
  var frontMatter = '---\ncreated: ' + Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm') + '\n---\n\n';
  var content = frontMatter + formatted;

  var file = folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  return { ok: true, fileId: file.getId(), fileName: fileName, content: content };
}

function getMemoFileOrThrow_(fileId, folder) {
  var file = DriveApp.getFileById(fileId);
  var parents = file.getParents();
  var belongsToFolder = false;
  while (parents.hasNext()) {
    if (parents.next().getId() === folder.getId()) { belongsToFolder = true; break; }
  }
  if (!belongsToFolder) throw new Error('指定されたメモが見つかりません。');
  return file;
}

function handleListMemos_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };

  var dateFilter = (body.date || '').trim(); // 'yyyy-MM-dd' 形式、空なら全件
  var folder = getMemoFolder_();
  var it = folder.getFilesByType(MimeType.PLAIN_TEXT);
  var memos = [];
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (!/^\d{4}-\d{2}-\d{2}_\d{4}_メモ帳/.test(name)) continue;
    if (dateFilter && name.indexOf(dateFilter + '_') !== 0) continue;
    var snippet = '';
    try {
      var raw = f.getBlob().getDataAsString('UTF-8');
      var body2 = raw.replace(/^---[\s\S]*?---\n*/, '');
      snippet = body2.replace(/\n/g, ' ').substring(0, 60);
    } catch (e) {}
    memos.push({
      fileId: f.getId(),
      fileName: name,
      updatedAt: f.getLastUpdated().toISOString(),
      snippet: snippet
    });
  }
  memos.sort(function (a, b) { return a.fileName < b.fileName ? 1 : -1; });
  return { ok: true, memos: memos };
}

function handleGetMemo_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  try {
    var folder = getMemoFolder_();
    var file = getMemoFileOrThrow_(body.fileId, folder);
    var content = file.getBlob().getDataAsString('UTF-8');
    return { ok: true, fileName: file.getName(), content: content };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function handleUpdateMemo_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  try {
    var folder = getMemoFolder_();
    var file = getMemoFileOrThrow_(body.fileId, folder);
    file.setContent(body.content || '');
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function handleDeleteMemo_(body) {
  var email = getSessionEmail_(body.token);
  if (!email) return { ok: false, error: 'セッションが切れています。再ログインしてください。' };
  try {
    var folder = getMemoFolder_();
    var file = getMemoFileOrThrow_(body.fileId, folder);
    file.setTrashed(true);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ============ 動作確認用テスト関数（GASエディタから手動実行） ============
function testGemini_() {
  var result = callGemini_('こんにちは、今日は良い天気です。買い物リスト：牛乳、卵、パン。', []);
  Logger.log(result);
}
