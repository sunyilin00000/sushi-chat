// meeting-features.js
// 业务功能：字幕 / 屏幕共享 / 匿名投票 / 强制禁言 / 会议文件 / 文本聊天 + 苏轼发言

// ★ 管理员登录账号（发禁言/投票指令时其他客户端靠它识别管理员）
const ADMIN_USERS = ['admin'];

const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";
let token = null, trtc = null, TRTC = null, myUserId = null, isAdmin = false;

function apiCall(action, dataParams) {
  const body = { action, token };
  if (dataParams) body.data = dataParams;
  return fetch(TEAM_API, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    .then(r => r.json()).then(d => { if (d.error) throw new Error(d.error); return d; });
}
function sendMsg(cmdId, obj) {
  try { trtc.sendCustomMessage({ cmdId, data: new TextEncoder().encode(JSON.stringify(obj)).buffer }); }
  catch (e) { console.warn('[信令] 发送失败:', e); }
}
// 字幕条（画面底部，8秒淡出）
function showSubLine(text) {
  const bar = document.getElementById('subtitle-bar');
  const line = document.createElement('div');
  line.className = 'sub-line'; line.textContent = text;
  bar.appendChild(line);
  setTimeout(() => line.remove(), 8200);
  while (bar.children.length > 3) bar.firstChild.remove();
}

window.addEventListener('meeting-ready', function () {
  trtc = window.meeting.trtc; TRTC = window.meeting.TRTC;
  myUserId = window.meeting.userId; token = window.meeting.token;
  try { isAdmin = (JSON.parse(localStorage.getItem('xh_user') || '{}').role === 'admin'); } catch (e) {}
  if (isAdmin) ['btnVote', 'btnMuteAll', 'btnUnmute'].forEach(id => document.getElementById(id).classList.remove('hidden'));

  initCustomMessages();
  initSubtitleButton();
  initScreenShare();
  initVote();
  initMuteButtons();
  initFiles();
  initChat();   // ★ 新增：文本聊天
  document.getElementById('btnLeave').onclick = () => window.meeting.leave();
});

// ---------- 自定义消息 ----------
function initCustomMessages() {
  trtc.on(TRTC.EVENT.CUSTOM_MESSAGE, function (event) {
    let msg = {};
    try { msg = JSON.parse(new TextDecoder().decode(event.data)); } catch (e) { return; }
    const from = event.userId;

    if (msg.t === 'subtitle') showSubLine(msg.text);
    else if (msg.t === 'vote_start') openVoteModal(msg);
    else if (msg.t === 'vote_end') { closeVoteModal(); showVoteResult(msg.q, msg.opts, msg.counts); }
    else if (msg.t === 'mute_all') {
      if (ADMIN_USERS.includes(from)) { window.meeting.muteLocalAudio(); showSubLine('🔇 管理员已开启全体禁言'); }
    } else if (msg.t === 'unmute_all') {
      if (ADMIN_USERS.includes(from)) { window.meeting.unmuteLocalAudio(); showSubLine('🔊 管理员已解除全体禁言'); }
    } else if (msg.t === 'file') {
      showSubLine('📎 ' + from + ' 上传了文件：' + msg.name);
      refreshFileList();
    }
    // ★ 聊天消息：别人说的话直接上屏；苏轼的回答上屏+朗读（自己发的已在本地处理，避免重复）
    else if (msg.t === 'chat') { if (from !== myUserId) showChat(from, msg.text, 'other'); }
    else if (msg.t === 'sushi') { if (from !== myUserId) { showChat('苏轼', msg.text, 'sushi'); speakSushi(msg.text); } }
  });
}

// ---------- ★ 文本聊天 + 苏轼发言 ----------
function initChat() {
  const btn = document.getElementById('btnChat');
  btn.onclick = function () {
    const m = document.getElementById('chatModal');
    m.style.display = (m.style.display === 'flex') ? 'none' : 'flex';
  };
  document.getElementById('btnSendChat').onclick = sendChat;
  document.getElementById('chatInput').addEventListener('keydown', e => { if (e.key === 'Enter') sendChat(); });
}
function sendChat() {
  const input = document.getElementById('chatInput');
  const text = input.value.trim();
  if (!text) return;
  input.value = '';
  showChat(myUserId, text, 'me');
  sendMsg(7, { t: 'chat', text });
  // ★ @苏轼 提问：本地调后端 AI 代理，回答后广播全房间
  if (text.startsWith('@苏轼')) {
    const q = text.slice(3).trim();
    if (q) askSushi(q);
  }
}
async function askSushi(question) {
  showChat('苏轼', '（思索中…）', 'sushi');
  try {
    const d = await apiCall('sushi_chat', { question });
    showChat('苏轼', d.answer, 'sushi');   // 自己先看
    sendMsg(8, { t: 'sushi', text: d.answer });   // 广播给其他人显示+朗读
    speakSushi(d.answer);                  // 自己也朗读
  } catch (e) {
    showChat('系统', '苏轼接口调用失败: ' + e.message, 'sys');
  }
}
// 浏览器本地 TTS 朗读（免费，不用后端）
function speakSushi(text) {
  try {
    speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(text);
    u.lang = 'zh-CN'; u.rate = 0.95;
    speechSynthesis.speak(u);
  } catch (e) { console.warn('[朗读] 不支持:', e); }
}
// 聊天面板渲染一条消息
function showChat(who, text, cls) {
  const list = document.getElementById('chatList');
  const div = document.createElement('div');
  div.className = 'chat-msg ' + cls;
  div.innerHTML = '<span class="chat-who">' + who + '：</span>' + escapeHtml(text);
  list.appendChild(div);
  list.scrollTop = list.scrollHeight;
  while (list.children.length > 100) list.firstChild.remove();
}
function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------- 字幕 ----------
let recognition = null, subtitleOn = false;
function initSubtitleButton() {
  const btn = document.getElementById('btnSubtitle');
  btn.onclick = function () {
    if (subtitleOn) { recognition && recognition.stop(); subtitleOn = false; btn.innerText = '字幕：关'; return; }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { alert('当前浏览器不支持语音识别，请用 Chrome / Edge'); return; }
    recognition = new SR();
    recognition.lang = 'zh-CN'; recognition.continuous = true; recognition.interimResults = false;
    recognition.onresult = function (e) {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) {
          const text = '【' + myUserId + '】' + e.results[i][0].transcript.trim();
          if (text) { showSubLine(text); sendMsg(1, { t: 'subtitle', text }); }
        }
      }
    };
    recognition.onerror = function (e) { console.warn('[字幕] 识别错误:', e.error); };
    recognition.start();
    subtitleOn = true; btn.innerText = '字幕：开';
  };
}

// ---------- 屏幕共享 ----------
let sharing = false;
function initScreenShare() {
  const btn = document.getElementById('btnShare');
  const big = document.getElementById('bigscreen');
  btn.onclick = async function () {
    if (!sharing) {
      try {
        await trtc.startScreenShare({ view: 'screen-share-view' });
        sharing = true; big.dataset.sharer = myUserId; btn.innerText = '停止共享';
      } catch (e) { console.warn('[共享] 已取消或不支持:', e); }
    } else { await trtc.stopScreenShare(); }
  };
  trtc.on(TRTC.EVENT.SCREEN_SHARE_STOPPED, function () {
    sharing = false; big.style.display = 'none'; delete big.dataset.sharer; btn.innerText = '共享屏幕';
  });
}

// ---------- 匿名投票 ----------
let currentVoteId = null, votePollTimer = null, currentQ = '', currentOpts = [];
function initVote() {
  document.getElementById('btnVote').onclick = function () {
    document.getElementById('voteModal').style.display = 'flex';
    document.getElementById('voteAdminCreate').classList.remove('hidden');
    document.getElementById('voteParticipate').classList.add('hidden');
    document.getElementById('voteAdminPanel').classList.add('hidden');
    document.getElementById('voteModalTitle').innerText = '发起投票';
  };
  document.getElementById('btnStartVote').onclick = function () {
    const q = document.getElementById('voteQ').value.trim();
    const opts = [1, 2, 3, 4].map(i => document.getElementById('voteOpt' + i).value.trim()).filter(Boolean);
    if (!q || opts.length < 2) { alert('请填写问题和至少两个选项'); return; }
    currentVoteId = 'vote_' + Date.now(); currentQ = q; currentOpts = opts;
    sendMsg(2, { t: 'vote_start', vote_id: currentVoteId, q, opts });
    document.getElementById('voteAdminCreate').classList.add('hidden');
    document.getElementById('voteAdminPanel').classList.remove('hidden');
    showVoteMask(q, '投票进行中…');
    votePollTimer = setInterval(() => pollVoteCount(q), 2000);
    pollVoteCount(q);
  };
  document.getElementById('btnSubmitVote').onclick = async function () {
    const checked = document.querySelector('input[name=voteOpt]:checked');
    if (!checked) { alert('请先选择一个选项'); return; }
    try {
      await apiCall('meeting_vote_cast', { vote_id: currentVoteId, choice: Number(checked.value) });
      closeVoteModal();
      showVoteMask(currentQ, '已提交，等待公布结果…');
      showSubLine('🗳️ 你已投票');
    } catch (e) { alert(e.message); }
  };
  document.getElementById('btnEndVote').onclick = async function () {
    try {
      const d = await apiCall('meeting_vote_status', { vote_id: currentVoteId });
      sendMsg(2, { t: 'vote_end', q: currentQ, opts: currentOpts, counts: d.counts });
      clearInterval(votePollTimer);
      closeVoteModal();
      showVoteResult(currentQ, currentOpts, d.counts);
    } catch (e) { alert(e.message); }
  };
}
function pollVoteCount(q) {
  apiCall('meeting_vote_status', { vote_id: currentVoteId }).then(d => {
    const total = d.counts.reduce((a, b) => a + b, 0);
    document.getElementById('voteLiveCount').innerText =
      '已收到 ' + total + ' 票\n' + currentOpts.map((o, i) => o + '：' + (d.counts[i] || 0) + ' 票').join('\n');
    updateVoteMask(q, '投票进行中，已投 ' + total + ' 人');
  }).catch(() => {});
}
function openVoteModal(msg) {
  currentVoteId = msg.vote_id; currentQ = msg.q; currentOpts = msg.opts;
  document.getElementById('voteModal').style.display = 'flex';
  document.getElementById('voteModalTitle').innerText = '匿名投票';
  document.getElementById('voteAdminCreate').classList.add('hidden');
  document.getElementById('voteAdminPanel').classList.add('hidden');
  document.getElementById('voteParticipate').classList.remove('hidden');
  document.getElementById('voteQuestionText').innerText = msg.q;
  document.getElementById('voteOptions').innerHTML = msg.opts.map((o, i) =>
    `<label class="opt-item"><input type="radio" name="voteOpt" value="${i}"> ${o}</label>`).join('');
  showVoteMask(msg.q, '投票进行中…');
}
function closeVoteModal() { document.getElementById('voteModal').style.display = 'none'; clearInterval(votePollTimer); }
function showVoteMask(q, sub) {
  document.getElementById('vote-mask').style.display = 'flex';
  updateVoteMask(q, sub);
}
function updateVoteMask(q, sub) {
  document.getElementById('voteMaskContent').innerHTML =
    '<h2>🗳️ ' + escapeHtml(q) + '</h2><div class="mask-sub">' + sub + '<br>（投票期间，共享内容已遮罩）</div>';
}
function showVoteResult(q, opts, counts) {
  const total = counts.reduce((a, b) => a + b, 0);
  let html = '<h2>📊 投票结果：' + escapeHtml(q) + '</h2><div class="mask-sub">';
  opts.forEach((o, i) => {
    const n = counts[i] || 0, pct = total ? Math.round(n / total * 100) : 0;
    html += escapeHtml(o) + '：<b>' + n + '</b> 票（' + pct + '%）<br>';
  });
  html += '共 ' + total + ' 票</div>';
  document.getElementById('voteMaskContent').innerHTML = html;
  document.getElementById('vote-mask').style.display = 'flex';
  setTimeout(() => { document.getElementById('vote-mask').style.display = 'none'; }, 10000);
}

// ---------- 强制禁言 ----------
function initMuteButtons() {
  document.getElementById('btnMuteAll').onclick = function () {
    sendMsg(4, { t: 'mute_all' });
    window.meeting.muteRemoteAudio();
    showSubLine('🔇 你已开启全体禁言');
  };
  document.getElementById('btnUnmute').onclick = function () {
    sendMsg(5, { t: 'unmute_all' });
    window.meeting.unmuteRemoteAudio();
    showSubLine('🔊 你已解除全体禁言');
  };
}

// ---------- 会议文件 ----------
function initFiles() {
  document.getElementById('btnFiles').onclick = function () {
    document.getElementById('fileModal').style.display = 'flex';
    refreshFileList();
  };
  document.getElementById('btnUploadFile').onclick = async function () {
    const f = document.getElementById('fileInput').files[0];
    if (!f) { alert('请先选择文件'); return; }
    if (f.size > 700 * 1024) { alert('文件超过 700KB，请压缩后上传'); return; }
    const b64 = await new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result.split(',')[1]); r.readAsDataURL(f); });
    try {
      await apiCall('meeting_file_upload', { filename: f.name, content_b64: b64 });
      showSubLine('📎 文件上传成功：' + f.name);
      sendMsg(6, { t: 'file', name: f.name });
      refreshFileList();
    } catch (e) { alert('上传失败: ' + e.message); }
  };
  document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m) m.style.display = 'none'; }));
}
async function refreshFileList() {
  const box = document.getElementById('fileList');
  try {
    const d = await apiCall('meeting_file_list');
    if (!d.files || !d.files.length) { box.innerHTML = '<p style="color:#94a3b8;font-size:13px;">暂无文件</p>'; return; }
    box.innerHTML = d.files.map(f =>
      `<div class="file-item"><span>${escapeHtml(f.name)}（${Math.round(f.size / 1024)}KB）</span><button onclick="downloadFile('${f.path}', '${f.name}')">下载</button></div>`).join('');
  } catch (e) { box.innerHTML = '<p style="color:#ef4444;font-size:13px;">加载失败: ' + e.message + '</p>'; }
}
async function downloadFile(path, name) {
  try {
    const d = await apiCall('meeting_file_download', { path });
    const bin = atob(d.content_b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([arr]));
    a.download = name; a.click();
  } catch (e) { alert('下载失败: ' + e.message); }
}
