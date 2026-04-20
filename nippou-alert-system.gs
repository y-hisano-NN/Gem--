/**
 * =============================================================================
 * 日報 AI 解析アラートシステム
 * =============================================================================
 * 概要:
 *   サイボウズOfficeから前日の全日報を取得し、Gemini APIで経営課題の兆候を
 *   3段階（赤・青・黄）に分類して社長へHTML形式のメールを送信する。
 *
 * 【セットアップ手順】
 *   1. Google Apps Script (script.google.com) で新規プロジェクトを作成
 *   2. このファイルの内容をエディタに貼り付ける
 *   3. 下記 CONFIG の各値を環境に合わせて設定する
 *   4. 「スクリプトのプロパティ」に機密情報を保存する（推奨・後述）
 *   5. createDailyTrigger() を一度だけ手動実行してトリガーを登録する
 *
 * 【スクリプトプロパティへの機密情報の保存（推奨）】
 *   GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」に
 *   以下のキーと値を追加してください:
 *     - CYBOZU_USER       : サイボウズOfficeのログインID（メールアドレス）
 *     - CYBOZU_PASSWORD   : サイボウズOfficeのパスワード
 *     - GEMINI_API_KEY    : Gemini APIキー
 * =============================================================================
 */

// ===========================================================================
// ▼ 設定値（自社環境に合わせて変更してください）
// ===========================================================================
const CONFIG = {
  // ── サイボウズOffice（DevToolsで確認済みの値）────────────────────────────
  CYBOZU_SUBDOMAIN: 'jjtsk30j23wk',   // URLバーで確認済み

  CYBOZU_USER:     PropertiesService.getScriptProperties().getProperty('CYBOZU_USER')     || '',
  CYBOZU_PASSWORD: PropertiesService.getScriptProperties().getProperty('CYBOZU_PASSWORD') || '',

  CYBOZU_DID:  '7',                      // アプリID（DID=7）
  CYBOZU_VID:  '532',                    // ビューID（VID=532）
  CYBOZU_FSID: '3-353-3427997-u347',     // フィルターID（FSID）
  CYBOZU_OID:  '501',                    // クエリID（oid=501）

  /**
   * 日報CSVの列定義（実際のCSVで確認済み）
   * ヘッダー順: 日報日付(0), 氏名(1), 所属部署(2), 本日の業務内容(3), 気づき(4), 明日の予定(5)
   */
  CSV_COLUMNS: {
    date:       0,
    name:       1,
    department: 2,
    body:       3,
    notes:      4,
    tomorrow:   5,
  },

  // ── Gemini API ────────────────────────────────────────────────────────────
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '',
  GEMINI_MODEL:   'gemini-2.5-flash',

  // ── Gmail ─────────────────────────────────────────────────────────────────
  RECIPIENT_EMAIL: 'president@your-company.co.jp',
  SENDER_NAME:     '社長室AIアラートシステム',
};

// ===========================================================================
// ▼ エントリーポイント（タイムドリブントリガーから呼び出す関数）
// ===========================================================================
/**
 * メイン処理。毎朝8:00にトリガーから自動実行される。
 * 前日の日報を取得 → Gemini解析 → HTMLメール生成 → Gmail送信
 */
function main() {
  try {
    const yesterday = getYesterdayDate();
    Logger.log(`[START] 対象日: ${yesterday}`);

    // 1. サイボウズOfficeから前日の日報を取得
    const reports = fetchCybozuDailyReports(yesterday);

    if (reports.length === 0) {
      Logger.log('[INFO] 前日の日報が0件のため処理を終了します。');
      sendNoReportEmail(yesterday);
      return;
    }

    Logger.log(`[INFO] 取得件数: ${reports.length} 件`);

    // 2. Gemini APIで解析
    const analysisResult = analyzeWithGemini(reports, yesterday);

    // 3. HTML形式のレポートメールを生成して送信
    const htmlBody = generateHtmlEmail(analysisResult, yesterday, reports.length);
    sendAlertEmail(htmlBody, yesterday);

    Logger.log('[DONE] メール送信完了');

  } catch (e) {
    Logger.log(`[ERROR] ${e.message}\n${e.stack}`);
    sendErrorEmail(e);
  }
}

// ===========================================================================
// ▼ 日付ユーティリティ
// ===========================================================================
/**
 * 前日の日付を "YYYY-MM-DD" 形式の文字列で返す
 * @returns {string}
 */
function getYesterdayDate() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * "YYYY-MM-DD" を日本語表記 "YYYY年M月D日" に変換する
 * @param {string} dateStr
 * @returns {string}
 */
function formatDateJp(dateStr) {
  const [y, m, d] = dateStr.split('-');
  return `${y}年${parseInt(m)}月${parseInt(d)}日`;
}

// ===========================================================================
// ▼ サイボウズOffice データ取得（セッション認証 + CSV方式）
// ===========================================================================
/**
 * 【実装方式について】
 *
 * サイボウズOffice の「カスタムアプリ」には公式REST APIが存在しません。
 * そのため本実装では「セッション認証 + CSVダウンロード」方式を採用します。
 *
 * ┌──────────────────────────────────────────────────────────────┐
 * │  STEP1: GASがサイボウズOfficeにフォームPOSTでログイン          │
 * │         → Set-Cookieヘッダーでセッションクッキーを取得         │
 * │  STEP2: セッションクッキーを使ってCSVダウンロードURLにアクセス  │
 * │         → 前日分の日報CSVを取得                               │
 * │  STEP3: CSVをパースして日報データの配列に変換                  │
 * └──────────────────────────────────────────────────────────────┘
 *
 * 【事前準備（必須）】
 *   A. ChromeでDevToolsを開き（F12）、Networkタブを表示した状態で
 *      サイボウズOfficeにログインしてCSVを書き出す操作を行う。
 *      以下のURLを確認してCONFIGに設定する:
 *        - ログインURL  → CONFIG の CYBOZU_LOGIN_URL
 *        - CSVダウンロードURL → CONFIG の CYBOZU_CSV_BASE_URL
 *        - アプリID    → CONFIG の CYBOZU_APP_ID
 *
 *   B. カスタムアプリで1件だけCSVを手動書き出しし、列番号を確認する
 *        → CONFIG の CSV_COLUMNS を実際の列番号に合わせて修正する
 *
 * 【注意事項】
 *   ・SAML/SSO（Google Workspaceとのシングルサインオン）を使っている場合、
 *     パスワード認証が無効になっている可能性があります。
 *     その場合はサイボウズ管理者に「ローカル認証（パスワード認証）」の
 *     有効化を依頼するか、選択肢A（Googleスプレッドシート）への移行を検討してください。
 *   ・サイボウズOfficeの画面URLが変わった場合（アップデート等）は
 *     このコードのURL部分を修正する必要があります。
 */

/**
 * ─── 追加設定（CONFIG に追記してください） ──────────────────────────────────
 *
 * CONFIG に以下を追加してください（このファイルの上部CONFIGオブジェクト内）:
 *
 *   // ログインURL（Chromeのネットワークタブで確認した実際のURLを設定）
 *   CYBOZU_LOGIN_URL: 'https://your-subdomain.cybozu.com/o/',
 *
 *   // CSVダウンロードの基本URL（ネットワークタブで確認したURLのベース部分）
 *   // 例: 'https://your-subdomain.cybozu.com/o/a.cgi'
 *   CYBOZU_CSV_BASE_URL: 'https://your-subdomain.cybozu.com/o/a.cgi',
 */

/**
 * レスポンスの Set-Cookie ヘッダーを "name=value" の文字列に変換して返す
 *
 * @param {HTTPResponse} response
 * @returns {string}
 */
function extractCookies_(response) {
  const raw = response.getAllHeaders()['Set-Cookie'];
  if (!raw) return '';
  const list = Array.isArray(raw) ? raw : [raw];
  return list.map(c => c.split(';')[0]).join('; ');
}

/**
 * サイボウズOfficeにJSON方式でログインし、セッションCookieを返す
 *
 * 【ログインの仕組み（DevToolsで確認済み）】
 *   STEP1: GET /o/ でページを取得し、_REQUEST_TOKEN_（CSRFトークン）を抽出する
 *   STEP2: POST /o/login.json へ JSON で username/password/_REQUEST_TOKEN_ を送信
 *   STEP3: レスポンスの Set-Cookie から JSESSIONID を含む Cookie を取得する
 *
 * @returns {string} セッションCookie文字列（後続リクエストのCookieヘッダーに使用）
 */
function loginToCybozuOffice() {
  const user     = CONFIG.CYBOZU_USER;
  const password = CONFIG.CYBOZU_PASSWORD;
  const baseUrl  = `https://${CONFIG.CYBOZU_SUBDOMAIN}.cybozu.com`;

  if (!user || !password) {
    throw new Error('CYBOZU_USER または CYBOZU_PASSWORD がスクリプトプロパティに未設定です。');
  }

  // ─── STEP 1: ログインページを取得して _REQUEST_TOKEN_ を抽出 ─────────────
  const pageRes     = UrlFetchApp.fetch(`${baseUrl}/o/`, { muteHttpExceptions: true });
  const pageHtml    = pageRes.getContentText();
  const pageCookies = extractCookies_(pageRes);

  // HTML内から _REQUEST_TOKEN_ を複数パターンで検索
  let requestToken = null;
  const tokenPatterns = [
    /"_REQUEST_TOKEN_"\s*:\s*"([^"]+)"/,
    /'_REQUEST_TOKEN_'\s*:\s*'([^']+)'/,
    /name="_REQUEST_TOKEN_"\s+value="([^"]+)"/,
    /name='_REQUEST_TOKEN_'\s+value='([^']+)'/,
    /requestToken['":\s]+([0-9a-f-]{36})/i,
  ];
  for (const pattern of tokenPatterns) {
    const match = pageHtml.match(pattern);
    if (match) { requestToken = match[1]; break; }
  }

  if (requestToken) {
    Logger.log(`[LOGIN] _REQUEST_TOKEN_ 取得成功: ${requestToken.substring(0, 8)}...`);
  } else {
    Logger.log('[LOGIN] _REQUEST_TOKEN_ がHTMLに見つかりませんでした。トークンなしで続行します。');
  }

  // ─── STEP 2: JSON POSTでログイン ─────────────────────────────────────────
  const loginUrl = `${baseUrl}/o/login.json?_lc=ja`;
  const body     = { username: user, password: password, keepUsername: true, redirect: '' };
  if (requestToken) body['_REQUEST_TOKEN_'] = requestToken;

  const loginRes    = UrlFetchApp.fetch(loginUrl, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(body),
    headers:            { 'Cookie': pageCookies },
    followRedirects:    false,
    muteHttpExceptions: true,
  });
  const loginStatus = loginRes.getResponseCode();

  if (loginStatus !== 200 && loginStatus !== 302) {
    throw new Error(
      `サイボウズOfficeへのログイン失敗: HTTPステータス ${loginStatus}\n` +
      `レスポンス: ${loginRes.getContentText().substring(0, 300)}`
    );
  }

  // ─── STEP 3: JSESSIONID を含む Cookie を取得 ─────────────────────────────
  const loginCookies = extractCookies_(loginRes);

  // ページCookieとログインCookieを結合（両方必要な場合に備える）
  const allCookies = [pageCookies, loginCookies].filter(Boolean).join('; ');

  if (!allCookies.includes('JSESSIONID')) {
    Logger.log('[WARN] JSESSIONID が Cookie に見つかりません。ログイン失敗の可能性があります。');
    Logger.log(`[DEBUG] 取得Cookie: ${allCookies.substring(0, 100)}`);
  } else {
    Logger.log(`[LOGIN] ログイン成功（JSESSIONID取得済み）`);
  }

  return allCookies;
}

/**
 * サイボウズOffice カスタムアプリから前日分のCSVをダウンロードして日報データを返す
 *
 * 【取得の仕組み（DevToolsで確認済み）】
 *   STEP1: ログインしてセッションCookieを取得
 *   STEP2: 書き出しページ（DBFileExport）をGETしてcsrf_ticketを抽出
 *   STEP3: csrf_ticketを含むフォームをPOSTしてShift-JIS形式のCSVを取得
 *   STEP4: CSVをパースして対象日のデータに絞り込む
 *
 * @param {string} date - "YYYY-MM-DD" 形式（前日の日付）
 * @returns {Array<{name:string, department:string, date:string, body:string}>}
 */
function fetchCybozuDailyReports(date) {
  const baseUrl = `https://${CONFIG.CYBOZU_SUBDOMAIN}.cybozu.com/o/ag.cgi`;

  // ── STEP 1: ログイン ──────────────────────────────────────────────────────
  const sessionCookie = loginToCybozuOffice();

  // ── STEP 2: 書き出しページを取得して csrf_ticket を抽出 ───────────────────
  const exportPageRes = UrlFetchApp.fetch(baseUrl, {
    method:             'get',
    headers:            { 'Cookie': sessionCookie },
    payload:            { page: 'DBFileExport', DID: CONFIG.CYBOZU_DID, VID: CONFIG.CYBOZU_VID,
                          FSID: CONFIG.CYBOZU_FSID, oid: CONFIG.CYBOZU_OID },
    muteHttpExceptions: true,
    followRedirects:    true,
  });
  const exportHtml = exportPageRes.getContentText('Shift_JIS');

  let csrfTicket = '';
  const csrfPatterns = [
    /name="csrf_ticket"\s+value="([^"]+)"/,
    /csrf_ticket['":\s]+([0-9a-f]{32})/i,
    /"csrf_ticket":"([^"]+)"/,
  ];
  for (const pattern of csrfPatterns) {
    const m = exportHtml.match(pattern);
    if (m) { csrfTicket = m[1]; break; }
  }
  Logger.log(csrfTicket
    ? `[CSV] csrf_ticket取得成功: ${csrfTicket.substring(0, 8)}...`
    : '[CSV] csrf_ticketが見つかりません。なしで続行します。'
  );

  // ── STEP 3: フォームPOSTでCSVダウンロード ────────────────────────────────
  const postPayload = {
    Page:      'DBFileExport',
    DID:       CONFIG.CYBOZU_DID,
    CT:        '1',
    cnt:       '1',
    encoding:  'Shift_JIS:NCR',
    fieldName: '',
    VID:       CONFIG.CYBOZU_VID,
    FSID:      CONFIG.CYBOZU_FSID,
    oid:       CONFIG.CYBOZU_OID,
    export:    '\u66f8\u304d\u51fa\u3059',   // 「書き出す」
    sid:       'csv',
  };
  if (csrfTicket) postPayload.csrf_ticket = csrfTicket;

  const csvRes    = UrlFetchApp.fetch(baseUrl, {
    method:             'post',
    payload:            postPayload,
    headers:            { 'Cookie': sessionCookie },
    muteHttpExceptions: true,
    followRedirects:    true,
  });
  const csvStatus = csvRes.getResponseCode();

  if (csvStatus !== 200) {
    throw new Error(
      `CSV取得エラー: HTTPステータス ${csvStatus}\n` +
      `レスポンス: ${csvRes.getContentText('Shift_JIS').substring(0, 300)}`
    );
  }

  // Shift-JIS で読み込む（encoding=Shift_JIS:NCR で書き出されるため）
  const csvText = csvRes.getContentText('Shift_JIS');
  Logger.log(`[CSV] 取得文字数: ${csvText.length}`);

  if (!csvText || csvText.trim() === '') {
    Logger.log('[INFO] CSVデータが空です（日報0件の可能性）。');
    return [];
  }

  // ── STEP 4: CSV パース → 対象日で絞り込み ────────────────────────────────
  return parseCsvToReports(csvText, date);
}

/**
 * CSVテキストを日報データの配列に変換する
 *
 * @param {string} csvText - CSVテキスト（UTF-8）
 * @param {string} date    - 対象日（"YYYY-MM-DD"）※CSVに日付がない場合のフォールバック用
 * @returns {Array<{name:string, department:string, date:string, body:string}>}
 */
function parseCsvToReports(csvText, date) {
  const col = CONFIG.CSV_COLUMNS;

  const rows = parseCsv(csvText);

  if (rows.length <= 1) {
    Logger.log('[INFO] CSVのデータ行が0件です。');
    return [];
  }

  const dataRows = rows.slice(1);

    const records = dataRows
      .map(row => {
        const bodyParts = [
          row[col.body]     ? `【本日の業務内容】\n${row[col.body].trim()}`  : '',
          row[col.notes]    ? `【気づき】\n${row[col.notes].trim()}`         : '',
          row[col.tomorrow] ? `【明日の予定】\n${row[col.tomorrow].trim()}`  : '',
        ].filter(Boolean);

        return {
          name:       (row[col.name]       || '不明').trim(),
          department: (row[col.department] || '不明').trim(),
          date:       (row[col.date]       || date).trim().replace(/\//g, '-').substring(0, 10),
          body:       bodyParts.join('\n\n'),
        };
      })
      .filter(r => r.body !== '');

    // 対象日のレコードのみに絞り込む（FSIDフィルターが全件返す場合の安全策）
    const filtered = records.filter(r => r.date === date);
    Logger.log(`[CSV] 全件: ${records.length} 件 → 対象日(${date}): ${filtered.length} 件`);
    return filtered;
}

/**
 * RFC 4180 準拠のCSVパーサー
 * ダブルクォート囲み・セル内改行・カンマを正しく処理する
 *
 * @param {string} csv
 * @returns {string[][]} 2次元配列
 */
function parseCsv(csv) {
  const rows   = [];
  let   row    = [];
  let   field  = '';
  let   inQuote = false;
  const text   = csv.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  for (let i = 0; i < text.length; i++) {
    const ch   = text[i];
    const next = text[i + 1];

    if (inQuote) {
      if (ch === '"' && next === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        inQuote = false;
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') {
        inQuote = true;
      } else if (ch === ',') {
        row.push(field);
        field = '';
      } else if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row   = [];
        field = '';
      } else {
        field += ch;
      }
    }
  }

  // 末尾の処理
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter(r => r.some(cell => cell.trim() !== ''));
}

// ===========================================================================
// ▼ Gemini API 解析
// ===========================================================================
/**
 * 日報データをGemini APIに送り、3段階アラート解析結果をJSONで受け取る
 *
 * @param {Array<{name,department,date,body}>} reports
 * @param {string} targetDate - "YYYY-MM-DD"
 * @returns {{alerts: Array}} - Geminiが返すJSONオブジェクト
 */
function analyzeWithGemini(reports, targetDate) {
  // APIキーはURLではなくヘッダーで渡す（ログへの漏洩を防ぐため）
  const endpoint =
    `https://generativelanguage.googleapis.com/v1beta/models/` +
    `${CONFIG.GEMINI_MODEL}:generateContent`;

  // 日報テキストをひとまとめにしてプロンプトへ埋め込む
  const reportsText = reports.map((r, i) =>
    `--- 日報 ${i + 1} ---\n` +
    `投稿者: ${r.name}\n` +
    `部署: ${r.department}\n` +
    `日付: ${r.date}\n` +
    `本文:\n${r.body}`
  ).join('\n\n');

  // ─── システムプロンプト（CoSとしての解析ロジック）────────────────────
  const systemPrompt = `あなたは株式会社ナガネツの有能な社長室長（CoS）です。
以下の日報テキストを読み込み、ナガネツのOS（COMPASS）から逸脱している記述を抽出し、3つのカテゴリーに分類してください。該当しない日報は無視してください。

【フィルター条件】
🔴 赤（最優先アラート：バッドニュース）
・条件：「クレーム」「怒られた」「遅れている」「失敗した」等、トラブルの火種や顧客からの厳しい指摘を示す事実。
・CoSコメント：大至急の事実確認と対応を促すコメント。

🔵 青（100%当事者の欠如：他責・行動停止）
・条件：主語が「自分」ではない他責の文脈。「気をつけます」「意識します」といった精神論で締めくくられ、「誰が(Who)・いつまでに(When)・何をする(What)」の次の一手（ボール）が存在しない記述。
・CoSコメント：具体的な行動（What）や、100%当事者としてのあり方を問うコメント。

🟡 黄（事実と解釈の混同：思考の解像度低下）
・条件：客観的な「事実」と、本人の「解釈（感情・推測）」が混ざっている文章。「何が起きたか」の羅列のみで「だからどうするか」の考察がない、または根拠のない推測。
・CoSコメント：事実と解釈を分離させ、思考の解像度を上げるための問いかけ。

【出力フォーマット（厳格なJSON形式・他のテキストは一切出力しないこと）】
{
  "alerts": [
    {
      "level": "red",
      "name": "投稿者名",
      "department": "部署名",
      "date": "投稿日",
      "extracted_text": "日報から抽出した原文の該当箇所（要約しないこと）",
      "cos_comment": "CoSとしての短いフィードバック案（問いかけ）"
    }
  ]
}

levelは "red" / "blue" / "yellow" のいずれかを使用すること。
alertsが空の場合は {"alerts": []} を返すこと。`;

  const requestBody = {
    system_instruction: {
      parts: [{ text: systemPrompt }],
    },
    contents: [
      {
        role: 'user',
        parts: [
          {
            text:
              `以下は${formatDateJp(targetDate)}（${targetDate}）の全社員日報です。\n\n` +
              reportsText,
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.2,         // 解析の一貫性を高めるため低めに設定
      responseMimeType: 'application/json', // JSON出力を強制
    },
  };

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers:     { 'x-goog-api-key': CONFIG.GEMINI_API_KEY },
    payload:     JSON.stringify(requestBody),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    throw new Error(
      `Gemini APIエラー: HTTPステータス ${statusCode}\n${response.getContentText()}`
    );
  }

  const json = JSON.parse(response.getContentText());

  // Geminiのレスポンスからテキスト部分を取り出す
  const rawText = json.candidates?.[0]?.content?.parts?.[0]?.text;

  if (!rawText) {
    throw new Error('Gemini APIからの応答が空です。');
  }

  // JSONブロックのみを抽出（念のためコードフェンスを除去）
  const cleaned = rawText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error(`GeminiのレスポンスJSONパース失敗:\n${cleaned}`);
  }
}

// ===========================================================================
// ▼ HTML メール生成
// ===========================================================================
/**
 * 解析結果をもとにHTML形式のレポートメールを生成する
 *
 * @param {{alerts: Array}} analysisResult
 * @param {string} targetDate - "YYYY-MM-DD"
 * @param {number} totalReports - 取得した日報の総件数
 * @returns {string} HTML文字列
 */
function generateHtmlEmail(analysisResult, targetDate, totalReports) {
  const alerts   = analysisResult?.alerts || [];
  const dateJp   = formatDateJp(targetDate);

  // アラートを優先度順（赤→青→黄）に並べ替え
  const ORDER = { red: 0, blue: 1, yellow: 2 };
  const sorted = [...alerts].sort((a, b) => (ORDER[a.level] ?? 9) - (ORDER[b.level] ?? 9));

  // ─── カラー定義 ──────────────────────────────────────────────────────────
  const COLORS = {
    red:    { bg: '#FFF0F0', border: '#E53935', badge: '#E53935', label: '🔴 最優先アラート（バッドニュース）' },
    blue:   { bg: '#F0F4FF', border: '#1E88E5', badge: '#1E88E5', label: '🔵 当事者意識の欠如（他責・行動停止）' },
    yellow: { bg: '#FFFDE7', border: '#F9A825', badge: '#F9A825', label: '🟡 事実と解釈の混同（思考の解像度低下）' },
  };

  // ─── アラートカードのHTML生成 ────────────────────────────────────────────
  const alertCardsHtml = sorted.length === 0
    ? `<div style="text-align:center;padding:40px 20px;color:#666;font-size:16px;">
         ✅ 前日の日報にCOMPASSからの逸脱は検出されませんでした。
       </div>`
    : sorted.map(alert => {
        const c = COLORS[alert.level] || COLORS.yellow;
        return `
          <div style="
            background:${c.bg};
            border-left:5px solid ${c.border};
            border-radius:6px;
            margin-bottom:20px;
            padding:18px 20px;
            font-family:sans-serif;
          ">
            <div style="margin-bottom:10px;">
              <span style="
                background:${c.badge};
                color:#fff;
                font-size:12px;
                font-weight:bold;
                padding:3px 10px;
                border-radius:20px;
              ">${c.label}</span>
            </div>
            <table style="border-collapse:collapse;width:100%;font-size:13px;color:#444;margin-bottom:12px;">
              <tr>
                <td style="padding:3px 8px 3px 0;font-weight:bold;white-space:nowrap;">投稿者</td>
                <td style="padding:3px 0;">${escapeHtml(alert.name)}</td>
                <td style="padding:3px 8px 3px 16px;font-weight:bold;white-space:nowrap;">部署</td>
                <td style="padding:3px 0;">${escapeHtml(alert.department)}</td>
                <td style="padding:3px 8px 3px 16px;font-weight:bold;white-space:nowrap;">日付</td>
                <td style="padding:3px 0;">${escapeHtml(alert.date)}</td>
              </tr>
            </table>
            <div style="margin-bottom:8px;">
              <div style="font-size:12px;font-weight:bold;color:#888;margin-bottom:4px;">📄 抽出原文</div>
              <div style="
                background:#fff;
                border:1px solid #ddd;
                border-radius:4px;
                padding:10px 14px;
                font-size:14px;
                line-height:1.7;
                color:#333;
                white-space:pre-wrap;
              ">${escapeHtml(alert.extracted_text)}</div>
            </div>
            <div>
              <div style="font-size:12px;font-weight:bold;color:#888;margin-bottom:4px;">💬 CoSフィードバック</div>
              <div style="
                background:rgba(0,0,0,0.04);
                border-radius:4px;
                padding:10px 14px;
                font-size:14px;
                line-height:1.7;
                color:#333;
                font-style:italic;
              ">${escapeHtml(alert.cos_comment)}</div>
            </div>
          </div>`;
      }).join('\n');

  // ─── サマリーバッジ ──────────────────────────────────────────────────────
  const countBy = level => alerts.filter(a => a.level === level).length;
  const summaryHtml = `
    <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:24px;">
      ${[
        { level: 'red',    label: '赤（最優先）' },
        { level: 'blue',   label: '青（当事者性）' },
        { level: 'yellow', label: '黄（思考解像度）' },
      ].map(({ level, label }) => {
        const c = COLORS[level];
        return `<div style="
          border:2px solid ${c.border};
          border-radius:8px;
          padding:10px 20px;
          text-align:center;
          min-width:120px;
        ">
          <div style="font-size:28px;font-weight:bold;color:${c.border};">${countBy(level)}</div>
          <div style="font-size:12px;color:#666;">${label}</div>
        </div>`;
      }).join('')}
      <div style="
        border:2px solid #ccc;
        border-radius:8px;
        padding:10px 20px;
        text-align:center;
        min-width:120px;
      ">
        <div style="font-size:28px;font-weight:bold;color:#555;">${totalReports}</div>
        <div style="font-size:12px;color:#666;">総日報数</div>
      </div>
    </div>`;

  // ─── メール全体のHTML ────────────────────────────────────────────────────
  return `<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:'Helvetica Neue',Arial,sans-serif;">
  <div style="max-width:700px;margin:24px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.12);">

    <!-- ヘッダー -->
    <div style="background:linear-gradient(135deg,#1a237e 0%,#283593 100%);padding:24px 28px;">
      <div style="color:#fff;font-size:20px;font-weight:bold;">🏢 日報 AI 解析アラート</div>
      <div style="color:#c5cae9;font-size:13px;margin-top:4px;">${dateJp}（前日分）の日報解析レポート</div>
    </div>

    <!-- サマリー -->
    <div style="padding:24px 28px 8px;">
      <h2 style="font-size:15px;color:#444;margin:0 0 16px;border-bottom:2px solid #e0e0e0;padding-bottom:8px;">📊 サマリー</h2>
      ${summaryHtml}
    </div>

    <!-- アラート一覧 -->
    <div style="padding:0 28px 28px;">
      <h2 style="font-size:15px;color:#444;margin:0 0 16px;border-bottom:2px solid #e0e0e0;padding-bottom:8px;">⚠️ アラート詳細（${alerts.length}件）</h2>
      ${alertCardsHtml}
    </div>

    <!-- フッター -->
    <div style="background:#f9f9f9;border-top:1px solid #eee;padding:14px 28px;font-size:11px;color:#aaa;text-align:center;">
      このメールは「日報AIアラートシステム」により自動生成されました。<br>
      解析モデル: ${CONFIG.GEMINI_MODEL} ／ 送信時刻: ${new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' })}
    </div>

  </div>
</body>
</html>`;
}

// ===========================================================================
// ▼ Gmail 送信
// ===========================================================================
/**
 * 社長へアラートメールを送信する
 *
 * @param {string} htmlBody
 * @param {string} targetDate - "YYYY-MM-DD"
 */
function sendAlertEmail(htmlBody, targetDate) {
  const subject = `【日報AIアラート】${formatDateJp(targetDate)}分の解析レポート`;

  GmailApp.sendEmail(
    CONFIG.RECIPIENT_EMAIL,
    subject,
    '※このメールはHTMLメール対応クライアントでご覧ください。',
    {
      htmlBody:  htmlBody,
      name:      CONFIG.SENDER_NAME,
      noReply:   true,
    }
  );
}

/**
 * 前日の日報が0件だった場合の通知メールを送信する
 *
 * @param {string} targetDate - "YYYY-MM-DD"
 */
function sendNoReportEmail(targetDate) {
  const subject = `【日報AIアラート】${formatDateJp(targetDate)}分：日報の投稿なし`;
  const body    = `${formatDateJp(targetDate)}の日報が0件でした。\n取得設定や投稿状況をご確認ください。`;

  GmailApp.sendEmail(CONFIG.RECIPIENT_EMAIL, subject, body, { name: CONFIG.SENDER_NAME });
}

/**
 * 処理中にエラーが発生した場合の通知メールを送信する
 *
 * @param {Error} error
 */
function sendErrorEmail(error) {
  const subject = '【日報AIアラート】エラーが発生しました';
  const body    =
    `日報AIアラートシステムの実行中にエラーが発生しました。\n\n` +
    `エラー内容:\n${error.message}\n\nスタックトレース:\n${error.stack}`;

  GmailApp.sendEmail(CONFIG.RECIPIENT_EMAIL, subject, body, { name: CONFIG.SENDER_NAME });
}

// ===========================================================================
// ▼ ユーティリティ
// ===========================================================================
/**
 * HTMLエスケープ（XSS対策）
 *
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

// ===========================================================================
// ▼ トリガー設定（初回セットアップ時に一度だけ手動実行してください）
// ===========================================================================
/**
 * 毎日 08:00 に main() を実行するタイムドリブントリガーを登録する。
 *
 * 【実行方法】
 *   1. GASエディタ上部の「関数を選択」ドロップダウンで "createDailyTrigger" を選択
 *   2. ▶ 実行ボタンをクリック
 *   3. 「トリガー」メニューで登録を確認する
 *
 * ※ 重複登録を防ぐため、既存の同名トリガーは先に削除してから登録します。
 */
function createDailyTrigger() {
  // 既存の main トリガーをすべて削除（重複防止）
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'main')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日 07:00〜08:00 の間に実行するトリガーを作成
  // ※ GASのatHour()は「指定時間帯の開始」を意味し、実際の起動は前後にずれる場合があります。
  //   確実に8時台に受信したい場合は 7 を指定してください。
  ScriptApp.newTrigger('main')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('✅ 毎日07:00〜08:00 (JST) に main() を実行するトリガーを登録しました。');
}

/**
 * 登録済みのトリガーをすべて表示する（デバッグ用）
 */
function listTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    Logger.log(`関数: ${t.getHandlerFunction()} / タイプ: ${t.getEventType()}`);
  });
}

/**
 * =============================================================================
 * ▼ サイボウズOffice API 実機確認ツール（最初に必ず実行してください）
 * =============================================================================
 *
 * 【背景】
 *   サイボウズOfficeには公式の日報APIが存在しません。
 *   本コードはガルーン（Garoon）の `/g/api/v1/` エンドポイントを流用していますが、
 *   cybozu.com上で動作するかどうかは実機テストでのみ確認できます。
 *
 * 【手順】
 *   1. CONFIG の CYBOZU_SUBDOMAIN を設定する
 *   2. スクリプトプロパティに CYBOZU_USER / CYBOZU_PASSWORD を設定する
 *   3. この関数を手動実行してログを確認する
 *   4. ログの内容に応じて下記「パターン別の対処」を参照する
 *
 * 【パターン別の対処】
 *   ✅ パターンA: レスポンスにrecordsなどのデータが返ってきた
 *      → fetchCybozuDailyReports() のフィールド名を実際のキー名に合わせて修正
 *
 *   ⚠️  パターンB: 認証エラー（401/403）が返ってきた
 *      → X-Cybozu-Authorization ヘッダーの形式をサイボウズ管理者に確認
 *      → APIアクセス自体がサイボウズの管理設定で制限されている可能性がある
 *
 *   ❌ パターンC: 404エラー / "method not found" 等が返ってきた
 *      → このエンドポイントはサイボウズOfficeでは使用不可
 *      → 下記「代替手段」を検討する必要がある
 *
 * 【サイボウズOffice 日報の代替取得手段】
 *   代替1: サイボウズOffice の「日報」をkintoneアプリに移行し、kintone REST APIで取得
 *           → 最も安定・公式。kintoneライセンスが別途必要。
 *   代替2: サイボウズOfficeの管理画面からCSVエクスポートを自動化し、GASで処理
 *           → GASのWebスクレイピング（UrlFetchApp＋セッション認証）が必要。
 *   代替3: Zapier / Make（旧Integromat）を仲介役に使いデータを取得してGASへ渡す
 *           → ノーコードで実現可能だが月額コストが発生する。
 * =============================================================================
 */

/**
 * サイボウズOffice APIの実機テスト
 * 最初に一度だけ手動実行し、ログでレスポンスの内容を確認してください。
 */
function debugCybozuApi() {
  const subdomain = CONFIG.CYBOZU_SUBDOMAIN;
  const user      = PropertiesService.getScriptProperties().getProperty('CYBOZU_USER');
  const password  = PropertiesService.getScriptProperties().getProperty('CYBOZU_PASSWORD');

  if (!user || !password) {
    Logger.log('❌ CYBOZU_USER または CYBOZU_PASSWORD がスクリプトプロパティに未設定です。');
    return;
  }

  const credentials = Utilities.base64Encode(`${user}:${password}`, Utilities.Charset.UTF_8);
  const yesterday   = getYesterdayDate();

  // ─── テスト1: /g/api/v1/ エンドポイント（Garoon形式） ───────────────────
  Logger.log('=== テスト1: Garoon形式エンドポイント (/g/api/v1/) ===');
  try {
    const url1  = `https://${subdomain}.cybozu.com/g/api/v1/`;
    const body1 = JSON.stringify({
      id: 1, jsonrpc: '2.0', method: 'NippoRecord.get',
      params: { date: yesterday, limit: 3, offset: 0 },
    });
    const res1 = UrlFetchApp.fetch(url1, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Cybozu-Authorization': credentials, 'Authorization': `Basic ${credentials}` },
      payload: body1, muteHttpExceptions: true,
    });
    Logger.log(`HTTP Status: ${res1.getResponseCode()}`);
    Logger.log(`Response: ${res1.getContentText().substring(0, 1000)}`);
  } catch (e) {
    Logger.log(`例外: ${e.message}`);
  }

  // ─── テスト2: /g/api/api.cgi エンドポイント（旧形式） ────────────────────
  Logger.log('\n=== テスト2: 旧形式エンドポイント (/g/api/api.cgi) ===');
  try {
    const url2  = `https://${subdomain}.cybozu.com/g/api/api.cgi`;
    const body2 = JSON.stringify({
      id: 1, jsonrpc: '2.0', method: 'NippoRecord.get',
      params: { date: yesterday, limit: 3, offset: 0 },
    });
    const res2 = UrlFetchApp.fetch(url2, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Cybozu-Authorization': credentials, 'Authorization': `Basic ${credentials}` },
      payload: body2, muteHttpExceptions: true,
    });
    Logger.log(`HTTP Status: ${res2.getResponseCode()}`);
    Logger.log(`Response: ${res2.getContentText().substring(0, 1000)}`);
  } catch (e) {
    Logger.log(`例外: ${e.message}`);
  }

  // ─── テスト3: 接続確認（認証のみ） ────────────────────────────────────────
  Logger.log('\n=== テスト3: 接続確認（User.get） ===');
  try {
    const url3  = `https://${subdomain}.cybozu.com/g/api/v1/`;
    const body3 = JSON.stringify({
      id: 1, jsonrpc: '2.0', method: 'User.get', params: { limit: 1, offset: 0 },
    });
    const res3 = UrlFetchApp.fetch(url3, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Cybozu-Authorization': credentials, 'Authorization': `Basic ${credentials}` },
      payload: body3, muteHttpExceptions: true,
    });
    Logger.log(`HTTP Status: ${res3.getResponseCode()}`);
    Logger.log(`Response: ${res3.getContentText().substring(0, 500)}`);
  } catch (e) {
    Logger.log(`例外: ${e.message}`);
  }

  Logger.log('\n=== デバッグ完了 ===');
  Logger.log('上記のレスポンスを確認し、動作パターン（A/B/C）を特定してください。');
}

/**
 * 動作確認用のテスト実行（今日の日付で動作確認する場合は targetDate を変更してください）
 */
function testRun() {
  const testDate = getYesterdayDate(); // 確認したい日付に変更可能（例: '2025-04-10'）

  Logger.log(`=== テスト実行: 対象日 ${testDate} ===`);

  // サイボウズOfficeへの接続テスト
  Logger.log('--- サイボウズOffice 接続テスト ---');
  const reports = fetchCybozuDailyReports(testDate);
  Logger.log(`取得件数: ${reports.length} 件`);
  if (reports.length > 0) {
    Logger.log(`最初の日報 (投稿者: ${reports[0].name}): ${reports[0].body.substring(0, 100)}...`);
  }

  // Gemini APIテスト（日報が0件の場合はダミーデータで確認）
  const testReports = reports.length > 0 ? reports : [
    { name: 'テスト太郎', department: '営業部', date: testDate, body: 'お客様からクレームがありました。担当者が対応します。気をつけます。' },
  ];
  Logger.log('--- Gemini API 解析テスト ---');
  const result = analyzeWithGemini(testReports, testDate);
  Logger.log(`解析結果 アラート数: ${result.alerts?.length ?? 0} 件`);
  Logger.log(JSON.stringify(result, null, 2));

  // HTML生成テスト
  Logger.log('--- HTML メール生成テスト ---');
  const html = generateHtmlEmail(result, testDate, testReports.length);
  Logger.log(`生成されたHTML文字数: ${html.length}`);

  Logger.log('=== テスト完了 ===');
}
