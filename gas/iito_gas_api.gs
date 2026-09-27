// ════════════════════════════════════════
// iito members 投票アプリ - GAS API
// ════════════════════════════════════════

const SHEET_VOTES   = '投票データ';
const SHEET_MEMBERS = '会員データ';
const SHEET_EVENTS  = 'イベント';
const SHEET_MODELS  = 'モデル';
const SHEET_GRANTS  = 'プラン付与履歴'; // その月に付与済みのプランを記録
const SECRET_KEY    = 'iito2026secret'; // HTMLと合わせること

// プランごとの基本持ち票（HTML側の PLANS と必ず一致させること）
const PLAN_VOTES = {
  supporter: 12,
  partner:   40,
  producer:  150
};
const DEFAULT_PLAN_VOTES = 40;

// 1セルに入れる最大文字数（GASの上限は約5万字。安全に4万で分割）
const CELL_CHUNK = 40000;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.key !== SECRET_KEY) {
      return corsRes({ ok: false, error: 'unauthorized' });
    }
    return corsRes(handleAction(data.action, data));
  } catch (err) {
    return corsRes({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.d) {
      const data = JSON.parse(decodeURIComponent(e.parameter.d));
      if (data.key !== SECRET_KEY) {
        return corsRes({ ok: false, error: 'unauthorized' });
      }
      return corsRes(handleAction(data.action, data));
    }
  } catch(err) {
    return corsRes({ ok: false, error: err.toString() });
  }
  return corsRes({ ok: true, msg: 'iito API running' });
}

// doPost / doGet 共通のアクション振り分け
function handleAction(action, data) {
  if (action === 'saveVote')    return saveVote(data);
  if (action === 'getResults')  return getResults(data);
  if (action === 'getTermPlans') return getTermPlans();
  if (action === 'saveMembers') return saveMembers(data);
  if (action === 'getMembers')  return getMembers();
  if (action === 'saveEvents')  return saveEvents(data);
  if (action === 'getEvents')   return getEvents();
  if (action === 'saveModels')  return saveModels(data);
  if (action === 'getModels')   return getModels();
  if (action === 'recordGrants') return recordGrants(data);
  if (action === 'getGrants')   return getGrants();
  if (action === 'getOshi')     return getOshi();
  if (action === 'ping')        return { ok: true, msg: 'pong' };
  return { ok: false, error: 'unknown action' };
}

function corsRes(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonRes(obj) { return corsRes(obj); }

// イベントIDから投票シートを取得（なければ作成）L列に「期」を追加
function getVoteSheet(evId, evName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const safeId = String(evId || 'noevent').replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 80);
  const sheetName = '投票_' + safeId;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue('イベント名：' + (evName || '') + '（ID: ' + safeId + '）');
    sheet.getRange(1, 1, 1, 12).merge().setFontWeight('bold').setBackground('#2b2b2b').setFontColor('#ffffff');
    sheet.getRange(2, 1, 1, 12).setValues([['タイムスタンプ', 'イベントID', 'イベント名', '会員番号', 'ニックネーム', 'プラン', 'メールアドレス', '推しモデル', '投票先モデルID', '投票先モデル名', '投票数', '期']]);
    sheet.getRange(2, 1, 1, 12).setFontWeight('bold').setBackground('#C9A84C').setFontColor('#ffffff');
    sheet.setFrozenRows(2);
  } else if (evName) {
    sheet.getRange(1, 1).setValue('イベント名：' + evName + '（ID: ' + safeId + '）');
  }
  return sheet;
}

// 会員が存在するか確認（ニックネーム一致）
function isMemberValid(nick) {
  if (!nick) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return rows.some(function(r){ return String(r[1]).trim() === String(nick).trim(); });
}

// 指定した「期(termId)」で、このニックネームが既に使った票の合計（全イベント横断）
function usedVotesInTerm(nick, termId) {
  if (!termId) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sum = 0;
  sheets.forEach(function(sh){
    const name = sh.getName();
    if (name.indexOf('投票_') !== 0) return;
    if (sh.getLastRow() <= 2) return;
    const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 12).getValues();
    rows.forEach(function(r){
      if (String(r[4]).trim() !== String(nick).trim()) return;
      if (String(r[11]).trim() !== String(termId).trim()) return;
      sum += Number(r[10]) || 0;
    });
  });
  return sum;
}

// 会員データから、指定ニックネームの現在のプランを取得
function getMemberPlan(nick) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(nick).trim()) {
      return String(rows[i][2]).trim(); // C列＝プラン
    }
  }
  return null;
}

// termId '2026年6月期' → 数値 202606（並べ替え用）
function termRankGAS(termId) {
  var m = String(termId || '').match(/^(\d{4})年(\d{1,2})月期$/);
  if (!m) return 0;
  return parseInt(m[1]) * 100 + parseInt(m[2]);
}

// このニックネームが投票したことのある全ての期を、投票記録から集める（年月順）
function listTermsForNick(nick) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  var set = {};
  sheets.forEach(function(sh){
    const name = sh.getName();
    if (name.indexOf('投票_') !== 0) return;
    if (sh.getLastRow() <= 2) return;
    const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 12).getValues();
    rows.forEach(function(r){
      if (String(r[4]).trim() !== String(nick).trim()) return;
      const t = String(r[11]).trim();
      if (t) set[t] = true;
    });
  });
  return Object.keys(set).sort(function(a,b){ return termRankGAS(a) - termRankGAS(b); });
}

// 【繰越判定用】このニックネームが「プラン付与を受けた期」の一覧を返す（＝その期に会員だった証拠）
//  投票記録ベースのlistTermsForNickだけだと「会員だが1票も使わなかった期」を拾えず、
//  未使用の票が繰り越されないバグになるため、付与履歴からも期を集める。
function listGrantTermsForNick(nick) {
  const sh = getGrantSheet_();
  var set = {};
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    rows.forEach(function(r){
      if (String(r[0]).trim() !== String(nick).trim()) return;
      const t = String(r[2]).trim();
      if (t) set[t] = true;
    });
  }
  return Object.keys(set);
}

// 【繰越判定用】「その期に会員だった」とみなせる期の一覧
//  ＝ 付与履歴のある期（未投票でも会員）∪ 投票記録のある期（履歴が無い移行期の救済）
function listMemberTermsForNick(nick) {
  var set = {};
  listGrantTermsForNick(nick).forEach(function(t){ set[t] = true; });
  listTermsForNick(nick).forEach(function(t){ set[t] = true; });
  return Object.keys(set).sort(function(a,b){ return termRankGAS(a) - termRankGAS(b); });
}

// 【繰越判定用】指定した期のちょうど1つ前の月の期IDを返す（1月なら前年12月）
//   例: '2026年8月期' → '2026年7月期' / '2026年1月期' → '2025年12月期'
function prevTermIdOfGAS(termId) {
  const mm = String(termId || '').match(/^(\d{4})年(\d{1,2})月期$/);
  if (!mm) return '';
  var y = parseInt(mm[1]), m = parseInt(mm[2]) - 1;
  if (m < 1) { m = 12; y -= 1; }
  return y + '年' + m + '月期';
}

// 前期にこのニックネームが使っていたプランを投票記録から取得（複数あれば上限が大きい方）
function getNickPlanInTerm(nick, termId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const rank = { supporter: 1, partner: 2, producer: 3 };
  var best = null;
  sheets.forEach(function(sh){
    const name = sh.getName();
    if (name.indexOf('投票_') !== 0) return;
    if (sh.getLastRow() <= 2) return;
    const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 12).getValues();
    rows.forEach(function(r){
      if (String(r[4]).trim() !== String(nick).trim()) return;
      if (String(r[11]).trim() !== String(termId).trim()) return;
      const p = String(r[5]).trim();
      if (p && (!best || (rank[p] || 0) > (rank[best] || 0))) best = p;
    });
  });
  return best;
}

// GAS側で「持ち票上限」を計算する（クライアントのlimitを信用しない）
//  新仕様：今月の付与合計（プラン変更で積み上がった票）＋ 前月繰越
//   - 今月付与合計 = プラン付与履歴から計算（サポーター12+パートナー40+... の合計）
//     履歴が無い場合は、現在のプランの基本票を最低保証とする（新規加入・移行期の安全策）
//   - 前月繰越 = 前月に「持っていた上限 − 使った票」の残り
// curTermId: 今回投票する期。curPlan: 今回のプラン（会員データの現プラン）
function calcServerLimit(nick, curTermId, curPlan) {
  const curBase = (curPlan && PLAN_VOTES[curPlan] != null) ? PLAN_VOTES[curPlan] : DEFAULT_PLAN_VOTES;
  if (!curTermId) return curBase; // 期が無ければ基本票のみ

  // ① 今月の付与合計（プラン変更で積み上がった票）
  var grantedThis = grantedVotesInTerm_(nick, curTermId);
  // 履歴が無い/現プランがまだ記録されていない場合は、現プランの基本票を最低保証
  if (grantedThis < curBase) grantedThis = curBase;

  // ② 前月・前々月繰越（直前2期分・連続会員が条件）
  //  ・前月繰越：直前期に会員なら (前月付与 − 前月使用)
  //  ・前々月繰越：前々期「と」前月の両方に会員（連続）なら (前々月付与 − 前々月使用)
  //    間の月に退会して会員記録が途切れると、前々月分は繰り越さない。
  //  会員実績の判定＝ 付与履歴のある期（未投票でも会員）∪ 投票記録のある期（移行期の救済）。
  const terms = listMemberTermsForNick(nick);
  const memberIn = function(t){ return terms.indexOf(t) >= 0; };
  const carryOfTerm = function(t){
    var g = grantedVotesInTerm_(nick, t);
    var u = usedVotesInTerm(nick, t);
    if (g === 0) {
      if (u > 0) { const p = getNickPlanInTerm(nick, t); g = (p && PLAN_VOTES[p] != null) ? PLAN_VOTES[p] : 0; }
      else { g = 0; }
    }
    return Math.max(0, g - u);
  };
  const prev1 = prevTermIdOfGAS(curTermId);            // 前月
  const prev2 = prev1 ? prevTermIdOfGAS(prev1) : '';   // 前々月
  var carry = 0;
  const inPrev1 = prev1 && memberIn(prev1);
  if (inPrev1) carry += carryOfTerm(prev1);
  if (inPrev1 && prev2 && memberIn(prev2)) carry += carryOfTerm(prev2); // 連続会員のときだけ前々月も
  return grantedThis + carry;
}

// Sheetsに保存されたイベントデータから、指定イベントの開始・終了時刻(ミリ秒)を取得
// 返り値: { start: ms|0, end: ms|0 } 見つからなければ両方0
function getEventWindow(evId) {
  try {
    const events = readBigJson(SHEET_EVENTS);
    if (!Array.isArray(events)) return { start: 0, end: 0 };
    const ev = events.filter(function(e){ return e && e.id === evId; })[0];
    if (!ev) return { start: 0, end: 0 };
    const start = (ev.startDate && ev.startTime) ? new Date(ev.startDate + 'T' + ev.startTime + ':00+09:00').getTime() : 0;
    const end   = (ev.endDate && ev.endTime)   ? new Date(ev.endDate   + 'T' + ev.endTime   + ':00+09:00').getTime() : 0;
    return { start: start, end: end };
  } catch(e) {
    return { start: 0, end: 0 };
  }
}

// 投票データを保存（会員確認＋締切チェック＋繰り越し込み上限チェック）
function saveVote(data) {
  const evId   = data.evId;
  const evName = data.evName;
  const nick   = data.nick;
  const plan   = data.plan;
  const mid    = data.mid;
  const email  = data.email;
  const oshi   = data.oshi;
  const votes  = data.votes || {};
  const termId = data.termId || '';
  const limit  = data.limit;

  // 1. 会員確認
  if (!isMemberValid(nick)) {
    return { ok: false, error: 'not_member', message: '会員として登録されていません' };
  }

  // 1.5 締切チェック（サーバー側）：Sheetsに保存されたイベントの開始・終了時刻で判定
  //     クライアントの申告時刻ではなく、GAS自身の現在時刻とSheetsのイベント設定を使う
  const win = getEventWindow(evId);
  const nowMs = Date.now();
  // グレース（締切ちょうどの送信タイミングずれを吸収）として10秒だけ許容
  if (win.end && nowMs > win.end + 10000) {
    return { ok: false, error: 'ended', message: 'この投票は終了しました' };
  }
  if (win.start && nowMs < win.start - 10000) {
    return { ok: false, error: 'before', message: 'この投票はまだ開始していません' };
  }

  // 2. 今回送信する票の合計
  let incoming = 0;
  Object.keys(votes).forEach(function(k){ incoming += Number(votes[k].count) || 0; });

  // 3. 上限チェック（繰り越し込み）
  //    ★セキュリティ：クライアントから送られる data.limit は信用せず、GAS側で計算する
  const curPlan = getMemberPlan(nick) || plan; // 会員データの現プランを正とする
  const effLimit = calcServerLimit(nick, termId, curPlan);

  let already = 0;
  if (termId) {
    already = usedVotesInTerm(nick, termId);
  } else {
    const sc = getVoteSheet(evId, evName);
    if (sc.getLastRow() > 2) {
      const rows = sc.getRange(3, 1, sc.getLastRow() - 2, 12).getValues();
      rows.forEach(function(r){ if (String(r[4]).trim() === String(nick).trim()) already += Number(r[10]) || 0; });
    }
  }

  if (already + incoming > effLimit) {
    return { ok: false, error: 'over_limit', message: '持ち票の上限を超えています', already: already, incoming: incoming, limit: effLimit };
  }

  // 3.5 イベント上限チェック（1候補あたりの上限票数）
  //     ★クライアントを回避した超過投票を防ぐため、サーバーでも必ず検証する。
  //     「確定済み(この候補) + 今回(この候補) > 上限」ならエラー。他候補には影響しない。
  const perCap = getEventPerModelCap(evId);
  if (perCap > 0) {
    // この会員がこの候補に既に入れている票数（投票シートから集計）
    const priorByModel = usedVotesByModel(evId, nick);
    var capError = null;
    Object.keys(votes).forEach(function(modelId){
      const add = Number(votes[modelId].count) || 0;
      const prior = priorByModel[modelId] || 0;
      if (prior + add > perCap) {
        capError = { modelId: modelId, name: votes[modelId].name || modelId, prior: prior, add: add };
      }
    });
    if (capError) {
      return { ok: false, error: 'over_model_cap', message: 'このイベントは1候補あたり' + perCap + '票までです', cap: perCap, detail: capError };
    }
  }

  // 4. 検証OK → 記録（差分を新しい行で追記）
  const sheet = getVoteSheet(evId, evName);
  const ts = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });

  Object.keys(votes).forEach(function(modelId){
    const v = votes[modelId];
    sheet.appendRow([ts, evId, evName || '', mid || '', nick, plan || '', email || '', oshi || '', modelId, v.name || modelId, v.count, termId]);
  });

  if (Object.keys(votes).length === 0) {
    sheet.appendRow([ts, evId, evName || '', mid || '', nick, plan || '', email || '', oshi || '', '', '(投票なし)', 0, termId]);
  }

  return { ok: true, saved: Object.keys(votes).length };
}

// イベントの「1候補あたりの上限票数」を取得（未設定/無効なら0＝制限なし）
function getEventPerModelCap(evId) {
  try {
    const events = readBigJson(SHEET_EVENTS);
    if (!Array.isArray(events)) return 0;
    const ev = events.filter(function(e){ return e && e.id === evId; })[0];
    if (!ev) return 0;
    const c = parseInt(ev.perModelCap);
    return (!isNaN(c) && c > 0) ? c : 0;
  } catch(e) { return 0; }
}

// 指定イベントで、この会員が「候補者ごとに」既に入れている票数 { modelId: 票数 }
function usedVotesByModel(evId, nick) {
  const result = {};
  const safeId = String(evId || 'noevent').replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 80);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('投票_' + safeId);
  if (!sheet || sheet.getLastRow() <= 2) return result;
  const rows = sheet.getRange(3, 1, sheet.getLastRow() - 2, 12).getValues();
  rows.forEach(function(r){
    if (String(r[4]).trim() !== String(nick).trim()) return; // 5列目=ニックネーム
    const modelId = String(r[8]).trim();                      // 9列目=投票先モデルID
    if (!modelId) return;
    result[modelId] = (result[modelId] || 0) + (Number(r[10]) || 0); // 11列目=投票数
  });
  return result;
}

// 集計データを取得
function getResults(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evId = data.evId;
  const safeId = String(evId || 'noevent').replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 80);
  const sheet = ss.getSheetByName('投票_' + safeId);
  if (!sheet || sheet.getLastRow() <= 2) {
    return { ok: true, votes: {} };
  }
  const rows = sheet.getRange(3, 1, sheet.getLastRow() - 2, 12).getValues();
  const voteMap = {};
  rows.forEach(function(row){
    const nick    = row[4];
    const modelId = row[8];
    const count   = Number(row[10]) || 0;
    if (!nick || !modelId) return;
    if (!voteMap[nick]) voteMap[nick] = {};
    voteMap[nick][modelId] = (voteMap[nick][modelId] || 0) + count;
  });
  return { ok: true, votes: voteMap };
}

// 各ニックネームが「各期(termId)」に投票時に使っていたプランを集計して返す
// 返り値: { ok:true, termPlans: { "ニックネーム": { "2026年6月期": "producer", ... }, ... } }
// 同じ期に複数プランの記録がある場合は、票数上限が最も大きいプラン（producer>partner>supporter）を採用する
function getTermPlans() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const rank = { supporter: 1, partner: 2, producer: 3 };
  const result = {}; // nick -> term -> plan
  sheets.forEach(function(sh){
    const name = sh.getName();
    if (name.indexOf('投票_') !== 0) return;
    if (sh.getLastRow() <= 2) return;
    const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 12).getValues();
    rows.forEach(function(r){
      const nick = String(r[4] || '').trim();
      const plan = String(r[5] || '').trim();  // F列＝プラン
      const term = String(r[11] || '').trim(); // L列＝期
      if (!nick || !plan || !term) return;
      if (!result[nick]) result[nick] = {};
      const cur = result[nick][term];
      // 同じ期に複数プラン記録があれば、上限の大きい方を採用（安全側）
      if (!cur || (rank[plan] || 0) > (rank[cur] || 0)) {
        result[nick][term] = plan;
      }
    });
  });
  return { ok: true, termPlans: result };
}

// 会員データを保存
function saveMembers(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) { sheet = ss.insertSheet(SHEET_MEMBERS); }
  sheet.clearContents();
  sheet.appendRow(['会員番号', 'ニックネーム', 'プラン', 'メールアドレス', 'ユーザーID']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#C9A84C').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  const members = data.members || [];
  members.forEach(function(m){
    sheet.appendRow([m.mid || '', m.nick || '', m.plan || '', m.email || '', m.uid || '']);
  });
  return { ok: true, saved: members.length };
}

// 会員データを取得
function getMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { ok: true, members: [] };
  }
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const members = rows
    .filter(function(r){ return r[1]; })
    .map(function(r){ return {
      mid:   r[0] || '',
      nick:  r[1] || '',
      plan:  r[2] || 'supporter',
      email: r[3] || '',
      uid:   r[4] || ''
    }; });
  return { ok: true, members };
}

// ════════════════════════════════════════
// 長いJSON文字列を複数セルに分割して保存／復元する共通ヘルパー
// A1に管理情報、A2以降にチャンクを縦に並べる
// ════════════════════════════════════════
function writeBigJson(sheetName, label, jsonStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); }
  sheet.clearContents();
  // チャンクに分割
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CELL_CHUNK) {
    chunks.push(jsonStr.substring(i, i + CELL_CHUNK));
  }
  // 1行目：ラベル＋チャンク数
  sheet.getRange(1, 1).setValue(label + '（JSON分割保存 / ' + chunks.length + 'チャンク）');
  sheet.getRange(1, 1).setFontWeight('bold').setBackground('#C9A84C').setFontColor('#ffffff');
  sheet.getRange(1, 2).setValue(chunks.length); // B1にチャンク数
  // 2行目以降：チャンクを縦に
  if (chunks.length > 0) {
    const values = chunks.map(function(c){ return [c]; });
    sheet.getRange(2, 1, chunks.length, 1).setValues(values);
  }
  return chunks.length;
}

function readBigJson(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const n = Number(sheet.getRange(1, 2).getValue()) || (sheet.getLastRow() - 1);
  if (n <= 0) return null;
  const rows = sheet.getRange(2, 1, n, 1).getValues();
  let str = '';
  rows.forEach(function(r){ str += (r[0] || ''); });
  if (!str) return null;
  try {
    return JSON.parse(str);
  } catch(e) {
    return null;
  }
}

// ════════════════════════════════════════
// イベントデータ（デバイス間同期）
// ════════════════════════════════════════
function saveEvents(data) {
  const events = data.events || [];
  writeBigJson(SHEET_EVENTS, 'イベントデータ', JSON.stringify(events));
  return { ok: true, saved: events.length };
}

function getEvents() {
  const events = readBigJson(SHEET_EVENTS);
  return { ok: true, events: Array.isArray(events) ? events : [] };
}

// ════════════════════════════════════════
// モデルデータ（デバイス間同期）※画像込みで大きくなるので分割保存
// ════════════════════════════════════════
function saveModels(data) {
  const models = data.models || {};
  writeBigJson(SHEET_MODELS, 'モデルデータ', JSON.stringify(models));
  const cnt = ((models.japan || []).length) + ((models.osaka || []).length);
  return { ok: true, saved: cnt };
}

function getModels() {
  const models = readBigJson(SHEET_MODELS);
  if (models && models.japan) {
    return { ok: true, models: models };
  }
  return { ok: true, models: null };
}

// ══════════════════════════════════════════
// プラン付与履歴：その月に「どのプランで投票券を付与したか」を記録
//   列: [ニックネーム, プラン, 期(YYYY年M月期), 記録日時]
//   同じ「ニックネーム×プラン×期」は1回だけ（月1回の付与ルール）
// ══════════════════════════════════════════

// 履歴シートを取得（無ければ作成）
function getGrantSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_GRANTS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GRANTS);
    sh.appendRow(['ニックネーム', 'プラン', '期', '記録日時']);
  }
  return sh;
}

// プラン付与を記録する（重複は自動でスキップ）
// data.grants = [{nick, plan, term}] の配列
function recordGrants(data) {
  const grants = data.grants || [];
  if (!grants.length) return { ok: true, added: 0 };
  const sh = getGrantSheet_();
  // 既存の「nick|plan|term」をセットにして重複チェック
  const existing = {};
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    rows.forEach(function(r){ existing[String(r[0]).trim()+'|'+String(r[1]).trim()+'|'+String(r[2]).trim()] = true; });
  }
  const now = new Date();
  const toAdd = [];
  grants.forEach(function(g){
    const nick = String(g.nick||'').trim();
    const plan = String(g.plan||'').trim();
    const term = String(g.term||'').trim();
    if (!nick || !plan || !term) return;
    const key = nick+'|'+plan+'|'+term;
    if (existing[key]) return; // 既に付与済み（月1回ルール）
    existing[key] = true;
    toAdd.push([nick, plan, term, now]);
  });
  if (toAdd.length) {
    sh.getRange(sh.getLastRow()+1, 1, toAdd.length, 4).setValues(toAdd);
  }
  return { ok: true, added: toAdd.length };
}

// 全付与履歴を取得 { nick: { term: [plan, plan, ...] } }
function getGrants() {
  const sh = getGrantSheet_();
  const result = {};
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    rows.forEach(function(r){
      const nick = String(r[0]).trim();
      const plan = String(r[1]).trim();
      const term = String(r[2]).trim();
      if (!nick || !plan || !term) return;
      if (!result[nick]) result[nick] = {};
      if (!result[nick][term]) result[nick][term] = [];
      if (result[nick][term].indexOf(plan) < 0) result[nick][term].push(plan);
    });
  }
  return { ok: true, grants: result };
}

// ════════════════════════════════════════
// 【管理画面用】各会員の推しモデルを投票シートから取得する
//   推しモデルは投票者のブラウザ(localStorage)にしか無いため、管理画面では表示できない。
//   ただし投票時に投票シートの8列目「推しモデル」へ保存されているので、そこから読み出す。
//   同じ人が複数回投票していれば、最も新しい行（＝最新の推し）を採用する。
//   返り値: { ok:true, oshi:{ ニックネーム: "モデルID,モデルID" } }
//   ※投票したことがある会員のみ対象（未投票の会員は取得できない）
// ════════════════════════════════════════
function getOshi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const latest = {}; // nick -> { ts: ミリ秒, oshi: '...' }
  sheets.forEach(function(sh){
    if (sh.getName().indexOf('投票_') !== 0) return;
    if (sh.getLastRow() <= 2) return;
    // 3行目以降が実データ（1行目=タイトル, 2行目=見出し）
    // 列: 1=タイムスタンプ, 5=ニックネーム, 8=推しモデル
    const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 8).getValues();
    rows.forEach(function(r){
      const nick = String(r[4]).trim();
      if (!nick) return;
      const oshi = String(r[7] || '').trim();
      if (!oshi) return; // 推し未設定の行は無視（過去の設定を消さない）
      var ts = 0;
      try { ts = (r[0] instanceof Date) ? r[0].getTime() : new Date(r[0]).getTime(); } catch (e) { ts = 0; }
      if (!ts || isNaN(ts)) ts = 0;
      if (!latest[nick] || ts >= latest[nick].ts) latest[nick] = { ts: ts, oshi: oshi };
    });
  });
  const result = {};
  Object.keys(latest).forEach(function(n){ result[n] = latest[n].oshi; });
  return { ok: true, oshi: result };
}

// その月にニックネームが付与された「プラン票の合計」を計算
//   例: サポーター(12)+パートナー(40)+プロデューサー(150) = 202
function grantedVotesInTerm_(nick, term) {
  const sh = getGrantSheet_();
  const PV = PLAN_VOTES; // {supporter:12, partner:40, producer:150}
  let sum = 0;
  const counted = {};
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    rows.forEach(function(r){
      if (String(r[0]).trim() !== String(nick).trim()) return;
      if (String(r[2]).trim() !== String(term).trim()) return;
      const plan = String(r[1]).trim();
      if (counted[plan]) return; // 同じプランは1回だけ
      counted[plan] = true;
      sum += (PV[plan] != null ? PV[plan] : 0);
    });
  }
  return sum;
}
