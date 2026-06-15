// ==========================================
// FanboxVoteSystem 共通JavaScript
// 04_style-guide.md / 01_core-spec.md 準拠
// ==========================================

// GAS API URL（全画面共通）
const GAS_API_URL = "https://script.google.com/macros/s/AKfycbxlHCDSZgIV8kQ7XpjTg6CXwVNxLszhjOzdXXQXa1ffiRw_X1qfR256vG-ZSP7d4SLWUg/exec";

// ==========================================
// IPアドレス取得
// ==========================================
async function getIpAddress() {
  try {
    const res = await fetch('https://api.ipify.org?format=json');
    const data = await res.json();
    return data.ip;
  } catch (e) {
    return "unknown";
  }
}

// ==========================================
// UUID取得・生成（正式キー: gacha_device_uuid）
// ==========================================
function getOrCreateUUID() {
  const KEY = 'gacha_device_uuid';
  let uuid = localStorage.getItem(KEY);

  if (!uuid) {
    uuid = crypto.randomUUID
      ? crypto.randomUUID()
      : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
          var r = Math.random() * 16 | 0;
          var v = c === 'x' ? r : (r & 0x3 | 0x8);
          return v.toString(16);
        });
    localStorage.setItem(KEY, uuid);
  }

  return uuid;
}

// ==========================================
// HTMLエスケープ
// ==========================================
function esc(s) {
  if (!s) return "";
  return s.replace(/[&<>"']/g, function(c) {
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
  });
}

// ==========================================
// メンテナンス表示（ページ全体を置き換え）
// 見出しは「🚧 メンテナンス中」で全画面統一
// 英語の MAINTENANCE は使用しない
// ==========================================
function showMaintenance(message) {
  document.body.innerHTML =
    '<div class="maintenance-box">' +
      '<h2>\uD83D\uDEA7 メンテナンス中</h2>' +
      '<p>' + (message ? esc(message).replace(/\n/g, '<br>') : '現在メンテナンス中です。再開までしばらくお待ちください。') + '</p>' +
    '</div>';
}

// ==========================================
// 支援者認証情報の保存
// ==========================================
function saveSupporter(pid, pass, name) {
  localStorage.setItem('supporter_pid', pid);
  localStorage.setItem('supporter_pass', pass);
  if (name) localStorage.setItem('supporter_name', name);
}

// ==========================================
// 支援者認証情報の読み出し
// ==========================================
function loadSupporter() {
  return {
    pid: localStorage.getItem('supporter_pid'),
    pass: localStorage.getItem('supporter_pass'),
    name: localStorage.getItem('supporter_name')
  };
}

// ==========================================
// 認証失敗時の後始末（supporter_pass のみ削除）
// ID・名前は残し、合言葉の再入力だけで再ログインできるようにする
// ==========================================
function handleAuthFailure() {
  localStorage.removeItem('supporter_pass');
}

// ==========================================
// 支援者ログアウト（supporter_* のみ削除）
// localStorage.clear() は使用禁止（gacha_device_uuid等が消えるため）
// ==========================================
function logoutSupporter() {
  localStorage.removeItem('supporter_pid');
  localStorage.removeItem('supporter_pass');
  localStorage.removeItem('supporter_name');
}

// ==========================================
// 支援者ログアウト（確認ダイアログ付き）
// 画面ごとに logout() を乱立させず、この関数を使うこと
// ==========================================
function confirmLogoutSupporter(message) {
  if (confirm(message || 'ログアウトしますか？')) {
    logoutSupporter();
    location.reload();
  }
}

// ==========================================
// 共通API呼び出し（GET）
// maintenance応答時は自動でメンテナンス画面に切り替え、null を返す
// ==========================================
async function apiGet(params) {
  var query = new URLSearchParams(params).toString();
  var res = await fetch(GAS_API_URL + '?' + query);
  var json = await res.json();
  if (json.status === 'maintenance') {
    showMaintenance(json.message);
    return null;
  }
  return json;
}

// ==========================================
// 共通API呼び出し（POST）
// maintenance応答のハンドリングは呼び出し元に委ねる
// ==========================================
async function apiPost(data) {
  var res = await fetch(GAS_API_URL, { method: 'POST', body: JSON.stringify(data) });
  return await res.json();
}

// ==========================================
// 結果描画（result.html / result_variant.html 共通）
// 描画先は id="ranking-list" の要素
// ==========================================
function renderResults(ranking) {
  var container = document.getElementById('ranking-list');
  container.innerHTML = "";

  if (!ranking || ranking.length === 0) {
    container.innerHTML = "<p>まだ投票データがありません。</p>";
    return;
  }

  var maxVote = ranking[0].count > 0 ? ranking[0].count : 1;

  ranking.forEach(function(item, index) {
    var percent = (item.count / maxVote) * 100;
    var rank = index + 1;

    var badgeClass = "rank-badge";
    if (rank === 1) badgeClass += " rank-1";
    else if (rank === 2) badgeClass += " rank-2";
    else if (rank === 3) badgeClass += " rank-3";

    var div = document.createElement('div');
    div.className = 'result-row';
    div.innerHTML =
      '<div class="result-header">' +
        '<div>' +
          '<span class="' + badgeClass + '">' + rank + '</span>' +
          '<span class="char-name">' + esc(item.character) + '</span>' +
          '<span class="char-theme">' + esc(item.theme) + '</span>' +
        '</div>' +
        '<span class="vote-count">' + item.count + '<small class="vote-count-unit">票</small></span>' +
      '</div>' +
      '<div class="result-bar-bg">' +
        '<div class="result-bar-fill" style="width: ' + percent + '%;"></div>' +
      '</div>';
    container.appendChild(div);
  });
}

// ==========================================
// 期間表示用日時整形
// "2026/03/25 00:00:00" → "2026年03月25日"
// ==========================================
function formatScheduleDate(dateTimeStr) {
  if (!dateTimeStr) return '';
  var parts = dateTimeStr.split('/');
  if (parts.length < 3) return dateTimeStr;
  var day = parts[2].split(' ')[0];
  return parts[0] + '年' + parts[1] + '月' + day + '日';
}

// ==========================================
// schedule-info 描画（全画面共通）
// container: 描画先要素
// schedule: { request_end?, vote_start, vote_end }
// ==========================================
function renderSchedule(container, schedule) {
  if (!container || !schedule) return;
  var html = '';
  if (schedule.request_end) {
    html += '<p>📅 リクエスト締め切り日：' + formatScheduleDate(schedule.request_end) + '</p>';
  }
  if (schedule.vote_start && schedule.vote_end) {
    html += '<p>🗳️ 投票期間：' + formatScheduleDate(schedule.vote_start) + '～' + formatScheduleDate(schedule.vote_end) + '</p>';
  }
  container.innerHTML = html;
}

// ==========================================
// 一括抽選結果（bundle）描画（全画面共通）
// vote_support.html / vote_variant.html / vote_revote.html 共通
// listEl   : ドロー一覧の描画先（.draw-list を付与した要素）
// totalEl  : 合計票数の描画先（数値を innerText で表示する要素）
// bundle   : { draw_results: [{index, rank, weight}], total_weight }
// 表示用 class に使う rank はホワイトリストで制限する（class 注入余地を排除）
// ==========================================
var GACHA_ALLOWED_RANKS = ['R', 'SR', 'SSR', 'UR'];

function normalizeGachaRank(rank) {
  if (GACHA_ALLOWED_RANKS.indexOf(rank) !== -1) return rank;
  console.warn('Unexpected rank value, fallback to R:', rank);
  return 'R';
}

function renderGachaBundle(listEl, totalEl, bundle) {
  if (!listEl || !bundle || !bundle.draw_results) return;
  listEl.innerHTML = '';

  bundle.draw_results.forEach(function(r) {
    var rank = normalizeGachaRank(r.rank);

    var row = document.createElement('div');
    row.className = 'draw-row';

    var label = document.createElement('span');
    label.className = 'draw-label';

    var ball = document.createElement('span');
    ball.className = 'gacha-ball draw-ball';
    ball.classList.add('ball-' + rank);
    ball.textContent = rank;

    var idx = document.createElement('span');
    idx.textContent = (r.index + 1) + '回目';

    label.appendChild(ball);
    label.appendChild(idx);

    var weight = document.createElement('span');
    weight.textContent = r.weight + '票';

    row.appendChild(label);
    row.appendChild(weight);
    listEl.appendChild(row);
  });

  if (totalEl) totalEl.innerText = bundle.total_weight;
}

// ==========================================
// 配列シャッフル（Fisher-Yates）
// ==========================================
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = array[i];
    array[i] = array[j];
    array[j] = tmp;
  }
  return array;
}
