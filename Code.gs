/**
 * Writing Submission Form — GAS バックエンド（送信の受付・記録・控えメール）
 * =====================================================================
 * フォーム本体（Index.html / index.html）は Google Sites または GitHub Pages に置き、
 * このスクリプトは「送信を受け取ってスプレッドシートに記録し、控えメールを送る」だけを担当します。
 *
 * 最初に一度だけ：エディタ上部の関数選択で「setup」を選んで ▶実行 → 権限を承認。
 * その後「デプロイ → 新しいデプロイ → ウェブアプリ（実行：自分／アクセス：全員）」。
 * ===================================================================== */


/* ===== 設定 ========================================================== */
// 空のままなら、このスクリプトが紐づいているスプレッドシートに記録します。
const SPREADSHEET_ID = "";
const SHEET_NAME = "Submissions";

const SEND_RECEIPT_EMAIL = true;                    // 学生に控えメールを送るか
const MAIL_SENDER_NAME   = "Writing Submission Form"; // 控えメールの差出人名
const MAX_ANSWER_CHARS   = 20000;                   // 解答の最大文字数（異常に長い送信を防ぐ）

// 控えメールを送ってよいメールアドレスのドメイン（@より後ろ）。空なら制限なし。例：["kanda.kifl.ac.jp"]
// 設定すると、それ以外の宛先には控えメールを送りません（提出の記録はされます）。
// ※フォーム側の CONFIG.allowedEmailDomains にも同じものを設定してください。
const ALLOWED_EMAIL_DOMAINS = [];

// スプレッドシートの列（この順で記録されます）
const HEADERS = [
  "課題名", "受信日時", "メールアドレス", "学籍番号", "氏名",
  "解答", "語数", "文字数",
  "入力時間(秒)", "打鍵数", "貼り付けブロック回数",
  "画面離脱回数", "離脱ログ", "端末情報", "受付ID"
];
/* ==================================================================== */

const ID_COL = HEADERS.indexOf("受付ID") + 1;

// 「書式なしテキスト」にする列。
// 自動の書式のままだと、課題名「8-10」が日付になったり、0で始まる学籍番号の0が消えたりするため。
const TEXT_HEADERS = ["課題名", "メールアドレス", "学籍番号", "氏名", "解答", "離脱ログ", "端末情報", "受付ID"];


/** /exec をブラウザで開いたときの表示（デプロイできているかの確認用） */
function doGet() {
  return ContentService.createTextOutput(
    "Writing Submission Form：受付中です。\nこのURLは送信先です（フォームの CONFIG.gasUrl に貼り付けてください）。"
  );
}


/** フォームからの送信口（text/plain で JSON を受け取り、JSON で結果を返す） */
function doPost(e) {
  try {
    let raw = {};
    if (e && e.postData && e.postData.contents) raw = JSON.parse(e.postData.contents);
    else if (e && e.parameter && e.parameter.payload) raw = JSON.parse(e.parameter.payload);
    return json_(process_(raw));
  } catch (err) {
    return json_({ status: "error", error: String((err && err.message) || err) });
  }
}


/**
 * 提出処理の本体。
 *  1) ロックの内側では「重複チェック＋1行追記」だけを最短で行う。
 *  2) ロックを解放してから控えメールを送る（メールが遅くても他の学生の提出を止めない）。
 */
function process_(raw) {
  const d = sanitize_(raw);
  if (!d.answer) return { status: "error", error: "empty" };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { status: "error", error: "busy" };   // 混雑：記録していないことを正直に返す

  let duplicate = false;
  try {
    const sheet = getSheet_();
    if (alreadyRecorded_(sheet, d.submissionId)) {
      duplicate = true;                                   // 再送：すでに記録済み
    } else {
      sheet.appendRow([
        cell_(d.assignment), new Date(), cell_(d.email), cell_(d.studentId), cell_(d.name),
        cell_(d.answer), d.wordCount, d.charCount,
        d.durationSec, d.keystrokes, d.pasteBlocked,
        d.blurCount, cell_(d.blurLog), cell_(d.userAgent), cell_(d.submissionId)
      ]);
      SpreadsheetApp.flush();                             // 次の送信の重複チェックに確実に反映させる
      if (d.submissionId) CacheService.getScriptCache().put("sub_" + d.submissionId, "1", 21600);
    }
  } finally {
    lock.releaseLock();
  }

  let mail = "off";
  if (duplicate) mail = "dup";
  else if (SEND_RECEIPT_EMAIL) mail = sendReceipt_(d);

  return { status: "ok", duplicate: duplicate, mail: mail, id: shortId_(d.submissionId) };
}


/** 送られてきた値を整える（文字数の上限・数値化。語数と文字数はサーバー側で数え直す） */
function sanitize_(raw) {
  raw = raw || {};
  const str = (v, max) => { let s = (v == null ? "" : String(v)).trim(); return max && s.length > max ? s.slice(0, max) : s; };
  const num = v => { const n = Number(v); return isFinite(n) && n >= 0 ? Math.floor(n) : 0; };
  const answer = str(raw.answer, MAX_ANSWER_CHARS);
  return {
    submissionId: str(raw.submissionId, 64),
    assignment:   str(raw.assignment || raw.problemNumber, 100),   // 旧フォーム（問題番号）からの送信にも対応
    name:         str(raw.name, 100),
    studentId:    str(raw.studentId, 50),
    email:        str(raw.email, 200),
    answer:       answer,
    wordCount:    answer ? answer.split(/\s+/).length : 0,
    charCount:    answer.length,
    durationSec:  num(raw.durationSec),
    keystrokes:   num(raw.keystrokes),
    pasteBlocked: num(raw.pasteBlocked),
    blurCount:    num(raw.blurCount),
    blurLog:      str(raw.blurLog, 5000),
    userAgent:    str(raw.userAgent, 500)
  };
}


/**
 * スプレッドシートの数式として解釈されないようにする。
 * 解答が「-」「+」「=」「@」で始まると、そのままでは数式扱いになり #ERROR! 等で内容が壊れるため。
 */
function cell_(v) {
  v = String(v == null ? "" : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}


/** 同じ受付IDがすでに記録されているか（キャッシュ → シートの順に確認） */
function alreadyRecorded_(sheet, id) {
  if (!id) return false;
  if (CacheService.getScriptCache().get("sub_" + id)) return true;
  const last = sheet.getLastRow();
  if (last < 2) return false;
  return !!sheet.getRange(2, ID_COL, last - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
}


/**
 * 記録用シートを取得する。
 * 以前の版で作られた列構成の違うシートがある場合は、列がずれて記録されないよう
 * 「Submissions_旧形式_日時」という名前で残し、新しい列構成のシートを作る。
 */
function getSheet_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("スプレッドシートが見つかりません。SPREADSHEET_ID を設定してください。");

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet && sheet.getLastRow() === 0) { writeHeaders_(sheet); return sheet; }
  if (sheet && !headersMatch_(sheet)) {
    sheet.setName(SHEET_NAME + "_旧形式_" + Utilities.formatDate(new Date(), tz_(), "yyyyMMdd_HHmm"));
    sheet = null;
  }
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME, 0); writeHeaders_(sheet); }
  return sheet;
}
function headersMatch_(sheet) {
  if (sheet.getLastColumn() < HEADERS.length) return false;
  const row = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  return HEADERS.every((h, i) => String(row[i]).trim() === h);   // 右側に先生が列を足していてもOK
}
function writeHeaders_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);
  setTextFormat_(sheet);
}
/** 文字の列を、列全体「書式なしテキスト」にする（すでに入っている値は変わりません） */
function setTextFormat_(sheet) {
  const rows = sheet.getMaxRows();
  TEXT_HEADERS.forEach(h => {
    const col = HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(1, col, rows, 1).setNumberFormat("@");
  });
}


/** 学生本人に「提出の控え」を送る。戻り値：sent / invalid / domain / quota / failed */
function sendReceipt_(d) {
  const to = d.email;
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return "invalid";
  if (!domainAllowed_(to)) return "domain";              // 許可していないドメイン：送らない（記録は済んでいる）
  try {
    if (MailApp.getRemainingDailyQuota() < 1) return "quota";
    const body =
      "以下の内容で提出を受け付けました。このメールは提出の控え（記録）です。\n" +
      "心当たりがない場合は破棄してください。\n\n" +
      "──────────────────────\n" +
      "氏名　　　　： " + d.name + "\n" +
      "学籍番号　　： " + d.studentId + "\n" +
      "課題名　　　： " + d.assignment + "\n" +
      "提出日時　　： " + Utilities.formatDate(new Date(), tz_(), "yyyy/MM/dd HH:mm:ss") + "\n" +
      "語数／文字数： " + d.wordCount + " 語 ／ " + d.charCount + " 文字\n" +
      "受付ID　　　： " + shortId_(d.submissionId) + "\n" +
      "──────────────────────\n\n" +
      "【提出した解答】\n" + d.answer + "\n" +
      "──────────────────────\n";
    MailApp.sendEmail({ to: to, subject: "【提出の控え】" + d.assignment, body: body, name: MAIL_SENDER_NAME });
    return "sent";
  } catch (err) {
    return "failed";
  }
}


/** メールアドレスのドメインが ALLOWED_EMAIL_DOMAINS に含まれるか（未設定なら常に true） */
function domainAllowed_(email) {
  const list = ALLOWED_EMAIL_DOMAINS.map(x => String(x).trim().toLowerCase().replace(/^@/, "")).filter(Boolean);
  if (!list.length) return true;
  const dom = String(email).toLowerCase().split("@").pop();
  return list.some(x => dom === x || dom.endsWith("." + x));
}


/* ===== 補助 ===== */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function tz_() { return Session.getScriptTimeZone() || "Asia/Tokyo"; }
function shortId_(id) { return String(id || "").replace(/-/g, "").slice(0, 8).toUpperCase(); }


/* ===== 先生がエディタから実行する関数（▶実行で使えるよう、名前の末尾に _ を付けていない） ===== */

/** 最初に一度実行：シートの準備と、スプレッドシート・メール送信の権限承認
 *  （すでに使っているシートでも、もう一度実行すると文字の列が「書式なしテキスト」になります） */
function setup() {
  const sheet = getSheet_();
  setTextFormat_(sheet);
  Logger.log("記録用シートの準備ができました：" + sheet.getName());
  Logger.log("今日送れる控えメールの残り：" + MailApp.getRemainingDailyQuota() + " 通");
}

/** 動作テスト：テスト行を1行記録し、自分宛てに控えメールを送る（確認後、テスト行は削除してください）
 *  課題名「8-10」と学籍番号「00000000」が、シートでそのままの文字で表示されることも確認できます。 */
function testSubmit() {
  const me = Session.getEffectiveUser().getEmail();
  Logger.log(JSON.stringify(process_({
    submissionId: "test-" + Date.now(),
    name: "テスト 花子", studentId: "00000000", assignment: "8-10", email: me,
    answer: "This is a test answer to check that the sheet records a row and the receipt email arrives.",
    durationSec: 120, keystrokes: 90, pasteBlocked: 0, blurCount: 0, blurLog: "", userAgent: "test"
  })));
}
