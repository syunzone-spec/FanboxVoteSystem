// ==================================================
// ⚙️ 設定エリア
// ==================================================
const SUPPORTER_MAX_VOTES = 5;       // 支援者の投票回数上限
const MONTHLY_PASSWORD 	= "";   // 今月の合言葉
const WEIGHT_SSR = 30; // ガチャ SSR
const WEIGHT_SR  = 10; // ガチャ SR
const WEIGHT_R   = 2;  // ガチャ R
const PROB_BORDER_SSR = 5;  // 5%
const PROB_BORDER_SR  = 20; // 15%

const SHEET_REQUESTS = 'Requests';
const SHEET_VOTES = 'Votes';
const SHEET_BLACKLIST = 'Blacklist';
const SHEET_LOGS = 'SystemLogs';
const SHEET_VARIANT_REQUESTS = 'VariantRequests';
const SHEET_VARIANT_VOTES    = 'VariantVotes';
const SHEET_CONFIG = 'Config';

// メンテナンス時のメッセージ
const MSG_MAINTENANCE = "現在メンテナンス中です。\n再開までしばらくお待ちください。";

// 対象イラストのプリセットリスト
const VARIANT_SUBJECTS = [
  "サンジェルマン"
];

const CURRENT_PERIOD_SALT = "2026_JAN_VOTE_V3"; 

// 列定義 (Requestsシート)
const COL_IDX_ID        = 0; // A列
const COL_IDX_CHARACTER = 3; // D列
const COL_IDX_THEME     = 4; // E列

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ==================================================
// 🌐 doGet (データ取得・認証)
// ==================================================
function doGet(e) {

	let sysKey = null;
  
  // 1. 一般投票/結果
  if (e.parameter.mode === 'public_vote' || e.parameter.mode === 'public_results') {
    sysKey = 'SYSTEM_PUBLIC';
  }
  // 2. 差分システム
  else if (e.parameter.mode === 'variant_init') {
    sysKey = 'SYSTEM_VARIANT';
  }
  // 3. 支援者投票 (pixiv_idがある場合)
  else if (e.parameter.pixiv_id) {
    sysKey = 'SYSTEM_SUPPORTER';
  }

  if (sysKey && !isSystemActive(sysKey)) {
    return createResponse({ status: 'maintenance', message: MSG_MAINTENANCE });
  }

  // 一般公開用データ
  if (e.parameter.mode === 'public_results' || e.parameter.mode === 'public_vote') {
    const candidates = getValidRequests();
    const results = aggregateResults(candidates);
    return createResponse({
      status: 'success',
      data: { candidates: candidates, results: results }
    });
  } else if (e.parameter.mode === 'variant_init') {
    return getVariantInitData(e); // 新設する関数へ丸投げ
  }


  // 2. 支援者ログイン
  // 支援者用データ
  const pixivId = e.parameter.pixiv_id;
  const password = e.parameter.password;

  if (!pixivId) return createResponse({ status: 'error', message: 'Pixiv ID is required' });
  if (password !== MONTHLY_PASSWORD) {
	return createResponse({ status: 'error', message: '合言葉が間違っています。\nFanbox記事をご確認ください。' });
  }

  // 現在の投票回数を取得
  const currentVoteCount = getVoteCountById(pixivId);
  const isFullyVoted = currentVoteCount >= SUPPORTER_MAX_VOTES;

  // 次のガチャ結果をプレビュー（現在の回数をシードにする）
  const nextGacha = calculateGachaSingle(pixivId, currentVoteCount);

  return createResponse({
    status: 'success',
    data: {
      user: {
        pixiv_id: pixivId,
        vote_count: currentVoteCount,     // 現在の投票数 (0〜5)
        max_votes: SUPPORTER_MAX_VOTES,   // 最大投票数 (5)
        is_fully_voted: isFullyVoted,     // 完了フラグ
        next_gacha: nextGacha             // 次回のガチャ結果(SSR等)
      },
      candidates: getValidRequests(),
      results: null // 投票画面では結果は見せない（リンクで誘導）
    }
  });
}

// ==================================================
// 📮 doPost (投票受付・分岐修正済み)
// ==================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const params = JSON.parse(e.postData.contents);

	const sysKey = getSystemKeyByMode(null, params.action);
    if (sysKey && !isSystemActive(sysKey)) {
      return createResponse({ status: 'maintenance', message: MSG_MAINTENANCE });
    }


    if (params.action === 'submit_request') {
      return processRequestSubmission(params);
    } else if (params.action === 'submit_vote_public') {
      return processPublicVote(params);
    } else if (params.action === 'submit_vote_supporter') {
      return processSupporterVote(params);
	} else if (params.action === 'submit_variant_request') {
      return processVariantRequest(params);
    } else if (params.action === 'submit_vote_variant') {
      return processVariantVote(params);
    } else {
      return createResponse({ status: 'error', message: 'Unknown action' });
    }
  } catch (error) {
    return createResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------
// 投票処理: 一般 (Public)
// ------------------------------------------
function processPublicVote(data) {
  if (isBlacklisted(data)) return createResponse({ status: 'success', message: 'Voted (Shadow)' });

  // 指紋生成
  const fingerprint = Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, 
    (data.ip_address || '') + (data.user_agent || '') + (data.screen_info || '')
  ));

  if (checkIfVotedFingerprint(fingerprint)) {
    return createResponse({ status: 'error', message: 'この端末からは既に投票済みです。' });
  }

  saveVote({
    id: Utilities.getUuid(),
    target_id: data.target_request_id,
    weight: 1, 
    voter_id: 'guest_' + fingerprint.substring(0, 8),
    ip: data.ip_address,
    ua: data.user_agent,
    uuid: data.device_uuid,
    note: 'Public Vote',
	pixivName: ''
  });

  return createResponse({ status: 'success', message: 'Voted' });
}


// ------------------------------------------
// 投票処理: 支援者 (Supporter)
// ------------------------------------------
function processSupporterVote(data) {
  if (isBlacklisted(data)) return createResponse({ status: 'success', message: 'Voted (Shadow)' });

  // 現在の投票回数を再確認
  const currentCount = getVoteCountById(data.pixiv_id);
  
  if (currentCount >= SUPPORTER_MAX_VOTES) {
    return createResponse({ status: 'error', message: '投票回数の上限に達しています。' });
  }

  // サーバー側でガチャ再計算（改ざん防止）
  // 渡された index (クライアント側で持っている回数) とサーバー側のカウントが一致するか確認
  // ※タイミングズレ防止のため、厳密にはサーバー側の currentCount を正とする
  const correctResult = calculateGachaSingle(data.pixiv_id, currentCount);

  saveVote({
    id: Utilities.getUuid(),
    target_id: data.target_request_id,
    weight: correctResult.weight,
    voter_id: data.pixiv_id,
    ip: data.ip_address,
    ua: data.user_agent,
    uuid: data.device_uuid,
    note: `Supporter Vote (${currentCount + 1}/${SUPPORTER_MAX_VOTES}): ${correctResult.rank}`,
	pixivName:data.user_name || ''
  });

  return createResponse({ status: 'success', message: 'Voted' });
}

// 共通: データ保存 (シート指定対応版)
function saveVote(p, sheetName) {
  const targetSheet = sheetName || SHEET_VOTES;
  const sheet = SS.getSheetByName(targetSheet);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  sheet.appendRow([
    p.id, p.target_id, p.weight, p.voter_id, 
    p.ip || '', p.ua || '', p.uuid || '', 
    timestamp, true, p.note, p.pixivName 
  ]);
}
// ------------------------------------------
// ユーティリティ
// ------------------------------------------



// ==================================================
// 🎲 ガチャ計算
// ==================================================

// 単発ガチャ計算 (Salt + ID + 回数インデックス)
function calculateGachaSingle(pixivId, index) {
  const input = String(pixivId) + CURRENT_PERIOD_SALT + "_" + index;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  let val = 0;
  for (let j = 0; j < digest.length; j++) { val += digest[j]; }
  const score = Math.abs(val) % 100;

  if (score < PROB_BORDER_SSR) return { rank: 'SSR', weight: WEIGHT_SSR };
  else if (score < PROB_BORDER_SR) return { rank: 'SR', weight: WEIGHT_SR };
  else return { rank: 'R', weight: WEIGHT_R };
}



// ==================================================
// 🛡️ 重複チェック・便利関数
// ==================================================


// ------------------------------------------
// 指紋による重複チェック（見た目通りの文字比較版）
// ------------------------------------------
function checkIfVotedFingerprint(fpHash) {
  const sheet = SS.getSheetByName(SHEET_VOTES);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const data = sheet.getRange(2, 4, lastRow - 1, 5).getDisplayValues(); 
  
  const targetId = 'guest_' + fpHash.substring(0, 8);
  // 今日の日付文字列（日本時間）
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  return data.some(row => {
    const id = row[0]; // A列
    const dateStrFull = row[4]; // H列 ("2026/01/20 13:47:11" という文字列)
    
    // 文字列の先頭10文字だけを切り取る
    const dateStr = dateStrFull.substring(0, 10); // "2026/01/20"
    
    // 文字同士で比較するので、絶対にズレません
    return (id === targetId && dateStr === todayStr);
  });
}

// 投票回数取得 (シート指定対応版)
// sheetName引数を省略した場合は、互換性のため SHEET_VOTES (本家) を参照します
function getVoteCountById(pixivId, sheetName) {
  const targetSheet = sheetName || SHEET_VOTES;
  const sheet = SS.getSheetByName(targetSheet);
  if (sheet.getLastRow() <= 1) return 0;
  
  // D列(voter_id)を取得
  const data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues().flat();
  const target = String(pixivId);
  
  // 一致する数をカウント
  return data.filter(id => String(id) === target).length;
}

function getValidRequests() {
  const sheet = SS.getSheetByName(SHEET_REQUESTS);
  if (sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][10] === true) { 
      list.push({ id: values[i][0], nickname: values[i][2], character: values[i][3], theme: values[i][4] });
    }
  }
  return list;
}

function aggregateResults(candidates) {
  const sheet = SS.getSheetByName(SHEET_VOTES);
  // データがなくても候補リストは返す（0票対応）
  const counts = {};
  
  if (sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][8] === true) {
        const targetId = values[i][1];
        const weight = Number(values[i][2]);
        if (!counts[targetId]) counts[targetId] = 0;
        counts[targetId] += weight;
      }
    }
  }

  // 全候補について票数をマッピング（0票も含む）
  return candidates.map(c => ({
    character: c.character, 
    theme: c.theme, 
    count: counts[c.id] || 0 
  })).sort((a, b) => b.count - a.count);
}


function processRequestSubmission(data) {
  if (isBlacklisted(data)) return createResponse({ status: 'success', id: Utilities.getUuid(), message: 'Received' });
  const sheet = SS.getSheetByName(SHEET_REQUESTS);
  const id = Utilities.getUuid();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([
    id, data.pixiv_id || '', data.nickname || '', data.character || '', data.theme || '', 
    '', data.ip_address || '', data.user_agent || '', data.device_uuid || '', data.screen_info || '', 
    false, timestamp, ''
  ]);
  return createResponse({ status: 'success', id: id, message: 'Received' });
}

function isBlacklisted(data) {
  const sheet = SS.getSheetByName(SHEET_BLACKLIST);
  if (sheet.getLastRow() <= 1) return false;
  const list = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const checkTargets = [
    { type: 'pixiv_id', value: String(data.pixiv_id || '') },
    { type: 'ip',       value: String(data.ip_address || '') },
    { type: 'uuid',     value: String(data.device_uuid || '') }
  ];
  return list.some(row => {
    const target = checkTargets.find(t => t.type === String(row[1]));
    return target && target.value === String(row[0]);
  });
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ==================================================
// 🆕 差分投票システム用ロジック
// ==================================================

// 初期化データ取得 (variant_init)
function getVariantInitData(e) {
  const pixivId = e.parameter.pixiv_id;
  const password = e.parameter.password;
  
  if (!pixivId || password !== MONTHLY_PASSWORD) {
    return createResponse({ status: 'error', message: '認証失敗: 合言葉またはIDが違います' });
  }

  // カウント確認（新シートを指定）
  const currentCount = getVoteCountById(pixivId, SHEET_VARIANT_VOTES);
  const isFullyVoted = currentCount >= SUPPORTER_MAX_VOTES; 
  const nextGacha = calculateGachaSingle(pixivId, currentCount);

  return createResponse({
    status: 'success',
    data: {
      user: {
        pixiv_id: pixivId,
        vote_count: currentCount,
        max_votes: SUPPORTER_MAX_VOTES,
        is_fully_voted: isFullyVoted,
        next_gacha: nextGacha
      },
      subjects: VARIANT_SUBJECTS,
      candidates: getValidVariantRequests()
    }
  });
}

// 差分リクエスト投稿 (セキュリティ情報込み)
function processVariantRequest(data) {
  if (isBlacklisted(data)) return createResponse({ status: 'success', message: 'Received' }); // Shadow Ban

  // 簡易認証
  if (data.password !== MONTHLY_PASSWORD) return createResponse({ status: 'error', message: 'Auth Failed' });
  if (!data.subject || !data.content) return createResponse({ status: 'error', message: '入力不足' });

  const sheet = SS.getSheetByName(SHEET_VARIANT_REQUESTS);
  const id = Utilities.getUuid();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  // Requestsシートと全く同じカラム順序で保存
  sheet.appendRow([
    id, 
    data.pixiv_id, 
    data.nickname, 
    data.subject, // character列として使用
    data.content, // theme列として使用
    '',           // attributes (予備)
    data.ip_address || '', 
    data.user_agent || '', 
    data.device_uuid || '', 
    data.screen_info || '', 
    false,         // is_valid (手動承認待ち)
    timestamp, 
    ''            // note
  ]);

  return createResponse({ status: 'success', message: 'Request Added' });
}

// 差分投票処理
function processVariantVote(data) {
  if (isBlacklisted(data)) return createResponse({ status: 'success', message: 'Voted (Shadow)' });

  // 投票上限チェック（新シートを指定）
  const currentCount = getVoteCountById(data.pixiv_id, SHEET_VARIANT_VOTES);
  if (currentCount >= SUPPORTER_MAX_VOTES) {
    return createResponse({ status: 'error', message: '投票回数の上限です。' });
  }

  // ガチャ結果計算
  const correctResult = calculateGachaSingle(data.pixiv_id, currentCount);

  // データ保存（新シートを指定して保存）
  saveVote({
    id: Utilities.getUuid(),
    target_id: data.target_request_id,
    weight: correctResult.weight,
    voter_id: data.pixiv_id,
    ip: data.ip_address,
    ua: data.user_agent,
    uuid: data.device_uuid,
    note: `Variant Vote (${currentCount + 1}): ${correctResult.rank}`,
    pixivName: data.user_name || ''
  }, SHEET_VARIANT_VOTES); // ★ここで新シートを指定

  return createResponse({ status: 'success', message: 'Voted' });
}

// 差分リクエスト一覧取得
function getValidVariantRequests() {
  const sheet = SS.getSheetByName(SHEET_VARIANT_REQUESTS);
  if (sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getValues();
  const list = [];
  
  for (let i = 1; i < values.length; i++) {
    // K列(index 10)が true なら有効
    if (values[i][10] === true) { 
      list.push({ 
        id: values[i][0], 
        nickname: values[i][2], 
        character: values[i][3], // subject
        theme: values[i][4]      // content
      });
    }
  }
  return list;
}

// ==================================================
// 🔧 メンテナンス制御
// ==================================================
function isSystemActive(systemKey) {
  const sheet = SS.getSheetByName(SHEET_CONFIG);
  if (!sheet) return true; // 設定シートが無ければ常に稼働とする（安全策）
  
  const values = sheet.getDataRange().getValues();
  // 1行目はヘッダーなのでスキップ
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === systemKey) {
      return values[i][2] === true; // C列がTRUEなら稼働
    }
  }
  return true; // キーが見つからない場合も稼働とする
}

function getSystemKeyByMode(mode, action) {
  // mode (GET) から判定
  if (mode === 'public_vote' || mode === 'public_results') return 'SYSTEM_PUBLIC';
  if (mode === 'variant_init') return 'SYSTEM_VARIANT';
  
  // action (POST) から判定
  if (action === 'submit_vote_public') return 'SYSTEM_PUBLIC';
  if (action === 'submit_vote_supporter') return 'SYSTEM_SUPPORTER';
  if (action && action.includes('variant')) return 'SYSTEM_VARIANT';
  
  // 支援者投票の初期化(doGet)はパラメータが特殊なので個別に判定が必要
  // ※doGet内で呼び出す際に手動で判定するため、ここは汎用的なもののみ
  return null; 
}


