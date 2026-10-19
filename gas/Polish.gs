/**
 * Polish.gs — 音声認識テキストを Gemini で整える（action: 'polishText'）
 * GASプロジェクトに「Polish.gs」として新規追加してください（Code.gs と同じプロジェクト内）。
 *
 * Code.gs の doPost の振り分けに、次の1行を追加：
 *   case 'polishText': return polishText_(data);
 * （if文方式なら： if (data.action === 'polishText') return polishText_(data); ）
 */

// Code.gs と同じものに合わせてください
const POLISH_GEMINI_MODEL = 'gemini-2.5-flash';
const POLISH_KEY_PROP = 'GEMINI_API_KEY';

function polishText_(data) {
  try {
    // ▼ createMemo と同じトークン確認関数に置き換えてください（未ログインの利用を防ぐため）
    // const user = getUserByToken_(data.token);
    // if (!user) return polishJson_({ ok: false, error: 'セッションが切れています。再ログインしてください。' });
    if (!data.token) return polishJson_({ ok: false, error: 'セッションが切れています。再ログインしてください。' });

    const text = String(data.text || '').trim();
    if (!text) return polishJson_({ ok: false, error: '整える文章がありません。' });
    if (text.length > 20000) return polishJson_({ ok: false, error: '文章が長すぎます。' });

    const apiKey = PropertiesService.getScriptProperties().getProperty(POLISH_KEY_PROP);
    if (!apiKey) return polishJson_({ ok: false, error: 'Gemini APIキーが未設定です。' });

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

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + POLISH_GEMINI_MODEL + ':generateContent';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.2 }
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return polishJson_({ ok: false, error: 'Gemini APIエラー (' + res.getResponseCode() + ')' });
    }
    const json = JSON.parse(res.getContentText());
    const parts = (((json.candidates || [])[0] || {}).content || {}).parts || [];
    const out = parts.map(function (p) { return p.text || ''; }).join('').trim();
    if (!out) return polishJson_({ ok: false, error: 'Geminiから結果が返りませんでした。' });

    return polishJson_({ ok: true, text: out });
  } catch (err) {
    return polishJson_({ ok: false, error: String(err && err.message || err) });
  }
}

function polishJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
