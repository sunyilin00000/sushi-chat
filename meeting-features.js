// meeting-features.js
// 业务功能：字幕 / 屏幕共享 / 匿名表决 / 全员禁言 / 会议文件 / 文字聊天（@补全）/ 苏轼感知发言 / 协同批注

const ADMIN_USERS = ['admin'];

const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";
let token = null, trtc = null, TRTC = null, myUserId = null, isAdmin = false;

// 会议讨论缓存（苏轼上下文）
const CONV = [];
function pushConv(who, text) {
  CONV.push(who + '：' + text);
  if (CONV.length > 30) CONV.shift();
}

// ★ 参会名单（@补全用）：自己 + 进出房事件维护
const roster = new Set();

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
  roster.add(myUserId);
  try { isAdmin = (JSON.parse(localStorage.getItem('xh_user') || '{}').role === 'admin'); } catch (e) {}
  if (isAdmin) ['btnVote', 'btnMuteAll', 'btnUnmute'].forEach(id => document.getElementById(id).classList.remove('hidden'));

  initCustomMessages();
  initSubtitleButton();
  initScreenShare();
  initVote();
  initMuteButtons();
  initFiles();
  initChat();
  initBoard();
  document.getElementById('btnLeave').onclick = () => window.meeting.leave();
  // 维护名单
  trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, e => roster.add(e.userId));
  trtc.on(TRTC.EVENT.REMOTE_USER_LEAVE, e => roster.delete(e.userId));
});

// ---------- 信令 ----------
function initCustomMessages() {
  trtc.on(TRTC.EVENT.CUSTOM_MESSAGE, function (event) {
    let msg = {};
    try { msg = JSON.parse(new TextDecoder().decode(event.data)); } catch (e) { return; }
    const from = event.userId;

    if (msg.t === 'subtitle') {
      showSubLine(msg.text);
      if (from !== myUserId) pushConv(from, stripName(msg.text));
    } else if (msg.t === 'vote_start') openVoteModal(msg);
    else if (msg.t === 'vote_end') { closeVoteModal(); showVoteResult(msg.q, msg.opts, msg.counts); }
    else if (msg.t === 'mute_all') {
      if (ADMIN_USERS.includes(from)) { window.meeting.muteLocalAudio(); showSubLine('【系统】管理员已开启全员禁言'); }
    } else if (msg.t === 'unmute_all') {
      if (ADMIN_USERS.includes(from)) { window.meeting.muteLocalAudio(); showSubLine('【系统】管理员已解除全员禁言'); }
    } else if (msg.t === 'file') {
      showSubLine('【系统】' + from + ' 上传了会议文件：' + msg.name);
      refreshFileList();
    } else if (msg.t === 'chat') {
      if (from !== myUserId) { showChat(from, msg.text, 'other'); pushConv(from, msg.text); }
    } else if (msg.t === 'sushi') {
      if (from !== myUserId) { showChat('苏轼', msg.text, 'sushi'); pushConv('苏轼', msg.text); speakSushi(msg.text); }
    }
    // ★ 批注同步
    else if (msg.t === 'draw') { if (from !== myUserId) { applyDraw(msg); markActive(from); } }
    else if (msg.t === 'btext') { if (from !== myUserId) { applyBoardText(msg); markActive(from); } }
    else if (msg.t === 'bclear') { if (from !== myUserId && ADMIN_USERS.includes(from)) clearBoardCanvas(); }
    else if (msg.t === 'bbg') { if (from !== myUserId) setBoardBg(msg.dataUrl); }
    else if (msg.t === 'cursor') { if (from !== myUserId) showCursorTag(from, msg.x, msg.y); }
  });
}
function stripName(t) { return t.replace(/^【.+?】/, ''); }

// ---------- 聊天 + @补全 ----------
function initChat() {
  document.getElementById('btnChat').onclick = function () {
    document.getElementById('chatPanel').classList.toggle('collapsed');
  };
  document.getElementById('btnChatToggle').onclick = function () {
    document.getElementById('chatPanel').classList.toggle('collapsed');
  };
  document.getElementById('btnSendChat').onclick = sendChat;
  const input = document.getElementById('chatInput');
  input.addEventListener('input', updateAtDropdown);
  input.addEventListener('keydown', function (e) {
    const dd = document.getElementById('atDropdown');
    if (dd.style.display === 'block') {
      if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); pickAt(dd.querySelector('div')); return; }
      if (e.key === 'Escape') { dd.style.display = 'none'; return; }
    }
    if (e.key === 'Enter') sendChat();
  });
  document.getElementById('btnSushiSpeak').onclick = function () {
    askSushi('请根据以上会议讨论记录，以你的视角发表一段简短意见或点评。');
  };
}
// ★ 输入 @ 显示成员下拉
function updateAtDropdown() {
  const input = document.getElementById('chatInput');
  const dd = document.getElementById('atDropdown');
  const m = input.value.match(/@([A-Za-z0-9_一-龥]*)$/);
  if (!m) { dd.style.display = 'none'; return; }
  const prefix = m[1].toLowerCase();
  const names = [...roster, '苏轼'].filter(n => n.toLowerCase().includes(prefix));
  if (!names.length) { dd.style.display = 'none'; return; }
  dd.innerHTML = names.map(n =>
    `<div data-name="${escapeHtml(n)}">${escapeHtml(n)}${n === '苏轼' ? '<span class="ai-tag">AI 成员</span>' : ''}</div>`).join('');
  dd.querySelectorAll('div').forEach(d => d.onmousedown = function (e) { e.preventDefault(); pickAt(d); });
  dd.style.display = 'block';
}
function pickAt(el) {
  if (!el) return;
  const input = document.getElementById('chatInput');
  input.value = input.value.replace(/@[A-Za-z0-9_一-龥]*$/, '@' + el.dataset.name + ' ');
  document.getElementById('atDropdown').style.display = 'none';
  input.focus();
}
function sendChat() {
  const input = document.getElementById('chatInput');
  const text = input.value.trim();
  if (!text) return;
  input.value = '';
  document.getElementById('atDropdown').style.display = 'none';
  showChat(myUserId, text, 'me');
  sendMsg(7, { t: 'chat', text });
  pushConv(myUserId, text);
  if (text.startsWith('@苏轼')) {
    const q = text.slice(3).trim();
    if (q) askSushi(q);
  }
}
async function askSushi(question) {
  showChat('苏轼', '（思索中）', 'sushi');
  try {
    const d = await apiCall('sushi_chat', { question, context: CONV.join('\n') });
    showChat('苏轼', d.answer, 'sushi');
    pushConv('苏轼', d.answer);
    sendMsg(8, { t: 'sushi', text: d.answer });
    speakSushi(d.answer);
  } catch (e) { showChat('系统', '苏轼接口调用失败：' + e.message, 'sys'); }
}
function speakSushi(text) {
  try {
    speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(text);
    u.lang = 'zh-CN'; u.rate = 0.95;
    speechSynthesis.speak(u);
  } catch (e) {}
}
function showChat(who, text, cls) {
  const list = document.getElementById('chatList');
  const div = document.createElement('div');
  div.className = 'chat-msg ' + cls;
  div.innerHTML = '<span class="chat-who">' + escapeHtml(who) + '：</span>' + highlightAt(text);
  list.appendChild(div);
  list.scrollTop = list.scrollHeight;
  while (list.children.length > 120) list.firstChild.remove();
}
function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function highlightAt(s) {
  return escapeHtml(s).replace(/@([A-Za-z0-9_一-龥]+)/g, '<span class="at">@$1</span>');
}

// ---------- 字幕 ----------
let recognition = null, subtitleOn = false;
function initSubtitleButton() {
  const btn = document.getElementById('btnSubtitle');
  btn.onclick = function () {
    if (subtitleOn) { recognition && recognition.stop(); subtitleOn = false; btn.innerText = '实时字幕'; return; }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { alert('当前浏览器不支持语音识别，请使用 Chrome 或 Edge'); return; }
    recognition = new SR();
    recognition.lang = 'zh-CN'; recognition.continuous = true; recognition.interimResults = false;
    recognition.onresult = function (e) {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) {
          const raw = e.results[i][0].transcript.trim();
          if (raw) {
            showSubLine('【' + myUserId + '】' + raw);
            pushConv(myUserId, raw);
            sendMsg(1, { t: 'subtitle', text: '【' + myUserId + '】' + raw });
          }
        }
      }
    };
    recognition.onerror = function (e) { console.warn('[字幕] 识别错误:', e.error); };
    recognition.start();
    subtitleOn = true; btn.innerText = '字幕开启中';
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
    sharing = false; big.style.display = 'none'; delete big.dataset.sharer; btn.innerText = '屏幕共享';
  });
}

// ---------- 协同批注 ----------
let boardCtx = null, boardTool = 'pen', drawing = false, curPts = [], lastCursorSent = 0;
function initBoard() {
  const modal = document.getElementById('boardModal');
  document.getElementById('btnBoard').onclick = () => { modal.style.display = 'flex'; resizeBoard(); };
  modal.addEventListener('click', e => { if (e.target === modal) modal.style.display = 'none'; });
  boardCtx = document.getElementById('boardCanvas').getContext('2d');
  // 工具切换
  document.querySelectorAll('.board-toolbar .tool[data-tool]').forEach(b => {
    b.onclick = function () {
      document.querySelectorAll('.board-toolbar .tool[data-tool]').forEach(x => x.classList.remove('active'));
      b.classList.add('active');
      boardTool = b.dataset.tool;
    };
  });
  document.getElementById('btnBoardClear').onclick = function () {
    clearBoardCanvas();
    sendMsg(10, { t: 'bclear' });
  };
  document.getElementById('btnBoardSave').onclick = saveBoard;
  document.getElementById('boardBgInput').onchange = uploadBoardBg;
  // 画布事件
  const cv = document.getElementById('boardCanvas');
  cv.addEventListener('pointerdown', boardDown);
  cv.addEventListener('pointermove', boardMove);
  cv.addEventListener('pointerup', boardUp);
  cv.addEventListener('pointerleave', boardUp);
  window.addEventListener('resize', resizeBoard);
}
function resizeBoard() {
  const wrap = document.getElementById('boardWrap');
  const cv = document.getElementById('boardCanvas');
  const img = cv.toDataURL();               // 保留已有笔迹
  cv.width = wrap.clientWidth; cv.height = wrap.clientHeight;
  const im = new Image();
  im.onload = () => boardCtx.drawImage(im, 0, 0, cv.width, cv.height);
  im.src = img;
}
function getPos(e) {
  const cv = document.getElementById('boardCanvas');
  const r = cv.getBoundingClientRect();
  return { x: (e.clientX - r.left) / r.width, y: (e.clientY - r.top) / r.height };
}
function boardDown(e) {
  const p = getPos(e);
  if (boardTool === 'text') {
    const text = prompt('输入文字内容：');
    if (text && text.trim()) {
      const msg = { t: 'btext', x: p.x, y: p.y, text: text.trim(), color: document.getElementById('penColor').value };
      applyBoardText(msg);
      sendMsg(10, msg);
    }
    return;
  }
  drawing = true; curPts = [[p.x, p.y]];
  document.getElementById('boardCanvas').setPointerCapture(e.pointerId);
}
function boardMove(e) {
  if (!drawing) return;
  const p = getPos(e);
  curPts.push([p.x, p.y]);
  // 本地实时画一小段
  applyDraw({ mode: boardTool === 'eraser' ? 'erase' : 'pen', color: document.getElementById('penColor').value, width: 3, pts: curPts.slice(-2) });
  // 批量发：累计 40 个点发一次（1KB 限制内）
  if (curPts.length >= 40) { flushStroke(); curPts = [[p.x, p.y]]; }
  // 光标位置节流发送，用于显示修改者名字
  const now = Date.now();
  if (now - lastCursorSent > 120) {
    lastCursorSent = now;
    sendMsg(11, { t: 'cursor', x: p.x, y: p.y });
  }
}
function boardUp() {
  if (!drawing) return;
  drawing = false;
  flushStroke();
}
function flushStroke() {
  if (curPts.length < 2) return;
  sendMsg(10, {
    t: 'draw', mode: boardTool === 'eraser' ? 'erase' : 'pen',
    color: document.getElementById('penColor').value, width: 3,
    pts: curPts.map(p => [Math.round(p[0] * 1000) / 1000, Math.round(p[1] * 1000) / 1000])
  });
}
function applyDraw(msg) {
  const cv = document.getElementById('boardCanvas');
  const c = boardCtx;
  c.save();
  if (msg.mode === 'erase') { c.globalCompositeOperation = 'destination-out'; c.strokeStyle = '#000'; c.lineWidth = msg.width * 4; }
  else { c.globalCompositeOperation = 'source-over'; c.strokeStyle = msg.color; c.lineWidth = msg.width; }
  c.lineCap = 'round'; c.lineJoin = 'round';
  c.beginPath();
  msg.pts.forEach((p, i) => { const x = p[0] * cv.width, y = p[1] * cv.height; i === 0 ? c.moveTo(x, y) : c.lineTo(x, y); });
  c.stroke();
  c.restore();
}
function applyBoardText(msg) {
  const cv = document.getElementById('boardCanvas');
  boardCtx.save();
  boardCtx.fillStyle = msg.color;
  boardCtx.font = 'bold 20px "PingFang SC", sans-serif';
  boardCtx.fillText(msg.text, msg.x * cv.width, msg.y * cv.height);
  boardCtx.restore();
}
function clearBoardCanvas() {
  const cv = document.getElementById('boardCanvas');
  boardCtx.save();
  boardCtx.globalCompositeOperation = 'source-over';
  boardCtx.clearRect(0, 0, cv.width, cv.height);
  boardCtx.restore();
}
function setBoardBg(dataUrl) {
  clearBoardCanvas();
  const cv = document.getElementById('boardCanvas');
  const im = new Image();
  im.onload = () => boardCtx.drawImage(im, 0, 0, cv.width, cv.height);
  im.src = dataUrl;
}
function uploadBoardBg() {
  const f = document.getElementById('boardBgInput').files[0];
  if (!f) return;
  if (f.size > 700 * 1024) { alert('底图超过 700KB，请压缩后再传'); return; }
  const r = new FileReader();
  r.onload = () => {
    setBoardBg(r.result);
    sendMsg(9, { t: 'bbg', dataUrl: r.result });
    showSubLine('【系统】' + myUserId + ' 更新了批注底图');
  };
  r.readAsDataURL(f);
}
async function saveBoard() {
  const cv = document.getElementById('boardCanvas');
  const dataUrl = cv.toDataURL('image/jpeg', 0.7);
  const b64 = dataUrl.split(',')[1];
  if (b64.length > 950000) { alert('批注图过大，请清空部分笔迹后再保存'); return; }
  try {
    await apiCall('meeting_file_upload', { filename: 'board_' + Date.now() + '.jpg', content_b64: b64 });
    showSubLine('【系统】批注已保存到会议文件');
    sendMsg(6, { t: 'file', name: '协同批注快照.jpg' });
  } catch (e) { alert('保存失败：' + e.message); }
}
// ★ 显示"谁正在改"：横幅 + 光标姓名标签
let activeTimer = null;
function markActive(user) {
  const st = document.getElementById('boardStatus');
  st.innerText = user + ' 正在批注';
  st.style.display = 'block';
  clearTimeout(activeTimer);
  activeTimer = setTimeout(() => st.style.display = 'none', 2500);
}
const tagTimers = {};
function showCursorTag(user, x, y) {
  const wrap = document.getElementById('boardWrap');
  let tag = document.getElementById('tag-' + user);
  if (!tag) {
    tag = document.createElement('div');
    tag.className = 'cursor-tag';
    tag.id = 'tag-' + user;
    wrap.appendChild(tag);
  }
  tag.innerText = user;
  tag.style.left = (x * 100) + '%';
  tag.style.top = (y * 100) + '%';
  tag.style.display = 'block';
  markActive(user);
  clearTimeout(tagTimers[user]);
  tagTimers[user] = setTimeout(() => tag.style.display = 'none', 2000);
}

// ---------- 表决 ----------
let currentVoteId = null, votePollTimer = null, currentQ = '', currentOpts = [];
function initVote() {
  document.getElementById('btnVote').onclick = function () {
    document.getElementById('voteModal').style.display = 'flex';
    document.getElementById('voteAdminCreate').classList.remove('hidden');
    document.getElementById('voteParticipate').classList.add('hidden');
    document.getElementById('voteAdminPanel').classList.add('hidden');
  };
  document.getElementById('btnStartVote').onclick = function () {
    const q = document.getElementById('voteQ').value.trim();
    const opts = [1, 2, 3, 4].map(i => document.getElementById('voteOpt' + i).value.trim()).filter(Boolean);
    if (!q || opts.length < 2) { alert('请填写议题和至少两个选项'); return; }
    currentVoteId = 'vote_' + Date.now(); currentQ = q; currentOpts = opts;
    sendMsg(2, { t: 'vote_start', vote_id: currentVoteId, q, opts });
    document.getElementById('voteAdminCreate').classList.add('hidden');
    document.getElementById('voteAdminPanel').classList.remove('hidden');
    showVoteMask(q, '表决进行中');
    votePollTimer = setInterval(() => pollVoteCount(q), 2000);
    pollVoteCount(q);
  };
  document.getElementById('btnSubmitVote').onclick = async function () {
    const checked = document.querySelector('input[name=voteOpt]:checked');
    if (!checked) { alert('请先选择一个选项'); return; }
    try {
      await apiCall('meeting_vote_cast', { vote_id: currentVoteId, choice: Number(checked.value) });
      closeVoteModal();
      showVoteMask(currentQ, '已提交，等待公布结果');
      showSubLine('【系统】表决已提交');
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
    updateVoteMask(q, '表决进行中，已投 ' + total + ' 人');
  }).catch(() => {});
}
function openVoteModal(msg) {
  currentVoteId = msg.vote_id; currentQ = msg.q; currentOpts = msg.opts;
  document.getElementById('voteModal').style.display = 'flex';
  document.getElementById('voteAdminCreate').classList.add('hidden');
  document.getElementById('voteAdminPanel').classList.add('hidden');
  document.getElementById('voteParticipate').classList.remove('hidden');
  document.getElementById('voteQuestionText').innerText = msg.q;
  document.getElementById('voteOptions').innerHTML = msg.opts.map((o, i) =>
    `<label class="opt-item"><input type="radio" name="voteOpt" value="${i}"> ${escapeHtml(o)}</label>`).join('');
  showVoteMask(msg.q, '表决进行中');
}
function closeVoteModal() { document.getElementById('voteModal').style.display = 'none'; clearInterval(votePollTimer); }
function showVoteMask(q, sub) {
  document.getElementById('vote-mask').style.display = 'flex';
  updateVoteMask(q, sub);
}
function updateVoteMask(q, sub) {
  document.getElementById('voteMaskContent').innerHTML =
    '<h2>表决 · ' + escapeHtml(q) + '</h2><div class="mask-sub">' + sub + '<br>表决期间，共享内容已遮蔽</div>';
}
function showVoteResult(q, opts, counts) {
  const total = counts.reduce((a, b) => a + b, 0);
  let html = '<h2>表决结果 · ' + escapeHtml(q) + '</h2><div class="mask-sub">';
  opts.forEach((o, i) => {
    const n = counts[i] || 0, pct = total ? Math.round(n / total * 100) : 0;
    html += escapeHtml(o) + '：<b>' + n + '</b> 票（' + pct + '%）<br>';
  });
  html += '总计 ' + total + ' 票</div>';
  document.getElementById('voteMaskContent').innerHTML = html;
  document.getElementById('vote-mask').style.display = 'flex';
  setTimeout(() => { document.getElementById('vote-mask').style.display = 'none'; }, 10000);
}

// ---------- 全员禁言 ----------
function initMuteButtons() {
  document.getElementById('btnMuteAll').onclick = function () {
    sendMsg(4, { t: 'mute_all' });
    window.meeting.muteRemoteAudio();
    showSubLine('【系统】你已开启全员禁言');
  };
  document.getElementById('btnUnmute').onclick = function () {
    sendMsg(5, { t: 'unmute_all' });
    window.meeting.unmuteRemoteAudio();
    showSubLine('【系统】你已解除全员禁言');
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
      showSubLine('【系统】文件上传成功：' + f.name);
      sendMsg(6, { t: 'file', name: f.name });
      refreshFileList();
    } catch (e) { alert('上传失败：' + e.message); }
  };
  document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m) m.style.display = 'none'; }));
}
async function refreshFileList() {
  const box = document.getElementById('fileList');
  try {
    const d = await apiCall('meeting_file_list');
    if (!d.files || !d.files.length) { box.innerHTML = '<p style="color:var(--muted);font-size:13px;">暂无文件</p>'; return; }
    box.innerHTML = d.files.map(f =>
      `<div class="file-item"><span>${escapeHtml(f.name)}（${Math.round(f.size / 1024)}KB）</span><button onclick="downloadFile('${f.path}', '${escapeHtml(f.name)}')">下载</button></div>`).join('');
  } catch (e) { box.innerHTML = '<p style="color:#d98a6a;font-size:13px;">加载失败：' + e.message + '</p>'; }
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
  } catch (e) { alert('下载失败：' + e.message); }
}
