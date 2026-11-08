// meeting-features.js
// 字幕 / 屏幕共享 / 匿名表决 / 全员禁言 / 单人屏蔽禁言 / 会议文件(分片上传) / 文字聊天 / 苏轼 / 协同批注 / 本地录屏

const ADMIN_USERS = ['syl'];

const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";
let token = null, trtc = null, TRTC = null, myUserId = null, isAdmin = false;

const CONV = [];
function pushConv(who, text) {
  CONV.push(who + '：' + text);
  if (CONV.length > 30) CONV.shift();
}
const roster = new Set();

// ★ 单人屏蔽禁言名单（管理员维护，全员本地执行屏蔽）
const mutedSet = new Set();

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
  if (isAdmin) {
    ['btnVote', 'btnMuteAll', 'btnUnmute'].forEach(id => document.getElementById(id).classList.remove('hidden'));
    document.querySelectorAll('.mute-btn').forEach(b => b.style.display = 'block');
  }

  initCustomMessages();
  initSubtitleButton();
  initScreenShare();
  initVote();
  initMuteButtons();
  initFiles();
  initChat();
  initBoard();
  initRecord();
  // ★ 离会：存档 + 苏轼纪要
  document.getElementById('btnLeave').onclick = async function () {
    let summary = '';
    try {
      if (CONV.length) {
        showSubLine('【系统】会议存档中，苏轼正在生成纪要…');
        const d = await apiCall('meeting_record_save', {
          room: String(window.meeting.roomId),
          name: document.getElementById('meetingName').innerText || '',
          lines: CONV.slice(-200)
        });
        summary = d.summary || '';
      }
    } catch (e) { console.warn('会议存档失败:', e); }
    if (summary) alert('【苏轼 · 会议纪要】\n\n' + summary);
    window.meeting.leave();
  };
  // 名单维护 + 管理员向新加入者同步禁言名单
  trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, function (e) {
    roster.add(e.userId);
    if (isAdmin && mutedSet.size) sendMsg(4, { t: 'mute_list', list: [...mutedSet] });
  });
  trtc.on(TRTC.EVENT.REMOTE_USER_LEAVE, e => roster.delete(e.userId));
  // ★ 管理员点击格子上禁言按钮（事件委托，含后加入者）
  document.getElementById('remote-videos').addEventListener('click', function (e) {
    const btn = e.target.closest('.mute-btn');
    if (!btn || !isAdmin) return;
    const uid = btn.dataset.uid;
    const muted = !mutedSet.has(uid);
    mutedSet[muted ? 'add' : 'delete'](uid);
    sendMsg(4, { t: 'mute_target', target: uid, muted: muted, list: [...mutedSet] });
    applyLocalMute(uid, muted);
  });
});

// ★ 本地执行屏蔽/恢复某人的声音
function applyLocalMute(uid, muted) {
  if (uid !== myUserId) {
    try { trtc.muteRemoteAudio(uid, muted); } catch (e) {}
  }
  const btn = document.querySelector('.mute-btn[data-uid="' + uid + '"]');
  if (btn) {
    btn.innerText = muted ? '已禁言' : '禁言';
    btn.classList.toggle('muted', muted);
  }
  if (uid === myUserId) {
    showSubLine(muted ? '【系统】管理员已屏蔽你的发言，其他人听不到你' : '【系统】管理员已恢复你的发言');
  }
}

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
      if (ADMIN_USERS.includes(from)) { window.meeting.unmuteLocalAudio(); showSubLine('【系统】管理员已解除全员禁言'); }
    }
    // ★ 单人禁言信令
    else if (msg.t === 'mute_target' && ADMIN_USERS.includes(from)) {
      if (Array.isArray(msg.list)) {
        const next = new Set(msg.list);
        mutedSet.forEach(u => { if (!next.has(u)) applyLocalMute(u, false); });   // 名单里被移除的恢复
        next.forEach(u => { if (!mutedSet.has(u)) applyLocalMute(u, true); });
        mutedSet.clear(); next.forEach(u => mutedSet.add(u));
      }
      if (msg.target) applyLocalMute(msg.target, msg.muted);
    } else if (msg.t === 'mute_list' && ADMIN_USERS.includes(from)) {
      if (Array.isArray(msg.list)) {
        const next = new Set(msg.list);
        mutedSet.forEach(u => { if (!next.has(u)) applyLocalMute(u, false); });
        next.forEach(u => { if (!mutedSet.has(u)) applyLocalMute(u, true); });
        mutedSet.clear(); next.forEach(u => mutedSet.add(u));
      }
    }
    else if (msg.t === 'file') {
      showSubLine('【系统】' + from + ' 上传了会议文件：' + msg.name);
      refreshFileList();
    } else if (msg.t === 'chat') {
      if (from !== myUserId) { showChat(from, msg.text, 'other'); pushConv(from, msg.text); }
    } else if (msg.t === 'sushi') {
      if (from !== myUserId) { showChat('苏轼', msg.text, 'sushi'); pushConv('苏轼', msg.text); speakSushi(msg.text); }
    }
    else if (msg.t === 'draw') { if (from !== myUserId) { applyDraw(msg); markActive(from); } }
    else if (msg.t === 'btext') { if (from !== myUserId) { applyBoardText(msg); markActive(from); } }
    else if (msg.t === 'bclear') { if (from !== myUserId && ADMIN_USERS.includes(from)) clearBoardCanvas(); }
    else if (msg.t === 'bbg') { if (from !== myUserId) setBoardBg(msg.dataUrl); }
    else if (msg.t === 'cursor') { if (from !== myUserId) showCursorTag(from, msg.x, msg.y); }
  });
}
function stripName(t) { return t.replace(/^【.+?】/, ''); }

// ---------- ★ 分片上传（突破 700KB：每片 600KB 单独走接口，下载时自动拼接） ----------
const CHUNK_SIZE = 600 * 1024;
function blobToB64(blob) {
  return new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result.split(',')[1]); r.readAsDataURL(blob); });
}
async function uploadBigFile(file) {
  const n = Math.max(1, Math.ceil(file.size / CHUNK_SIZE));
  const upid = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const safeBase = file.name.replace(/[^\w.\-一-龥]/g, '_').slice(0, 40);
  for (let i = 0; i < n; i++) {
    const b64 = await blobToB64(file.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE));
    await apiCall('meeting_file_upload', { filename: `${i}of${n}__${upid}__${safeBase}`, content_b64: b64 });
  }
  return { name: safeBase, path: null, upid: upid, parts: n };   // path 由下载方按 upid 查
}
function parsePartName(name) {
  const m = String(name).match(/^(\d+)of(\d+)__(\w+)__/);
  return m ? { i: Number(m[1]), n: Number(m[2]), upid: m[3], base: name.replace(/^\d+of\d+__\w+__/, '') } : null;
}
// 下载：单片直下；分片自动找齐拼合（上限 30MB）
async function downloadAny(filePayload) {
  if (!filePayload.upid) return downloadSingle(filePayload.path, filePayload.name);
  const d = await apiCall('meeting_file_list');
  const parts = (d.files || []).map(f => ({ f, p: parsePartName(f.name) }))
    .filter(x => x.p && x.p.upid === filePayload.upid)
    .sort((a, b) => a.p.i - b.p.i);
  if (!parts.length) throw new Error('找不到文件分片');
  const blobs = [];
  for (const x of parts) {
    const dd = await apiCall('meeting_file_download', { path: x.f.path });
    const bin = atob(dd.content_b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    blobs.push(arr);
  }
  const total = blobs.reduce((a, b) => a + b.length, 0);
  const all = new Uint8Array(total);
  let off = 0;
  blobs.forEach(b => { all.set(b, off); off += b.length; });
  triggerDownload(new Blob([all]), filePayload.name);
}
function triggerDownload(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name; a.click();
}
async function downloadSingle(path, name) {
  const d = await apiCall('meeting_file_download', { path });
  const bin = atob(d.content_b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  triggerDownload(new Blob([arr]), name);
}

// ---------- 本地录屏 ----------
let mediaRecorder = null, recChunks = [];
function initRecord() {
  const btn = document.getElementById('btnRecord');
  btn.onclick = async function () {
    if (mediaRecorder && mediaRecorder.state === 'recording') { mediaRecorder.stop(); return; }
    try {
      const stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: true });
      recChunks = [];
      mediaRecorder = new MediaRecorder(stream);
      mediaRecorder.ondataavailable = e => { if (e.data.size) recChunks.push(e.data); };
      mediaRecorder.onstop = function () {
        stream.getTracks().forEach(t => t.stop());
        const blob = new Blob(recChunks, { type: 'video/webm' });
        triggerDownload(blob, 'meeting_record_' + Date.now() + '.webm');
        btn.innerText = '录制'; btn.classList.remove('recording');
        showSubLine('【系统】录屏已保存到本地');
      };
      mediaRecorder.start();
      btn.innerText = '停止录制'; btn.classList.add('recording');
      showSubLine('【系统】录屏中（共享整个屏幕以获得完整会议画面）');
    } catch (e) { showSubLine('【系统】录屏已取消'); }
  };
}

// ---------- 聊天 + @补全 ----------
function initChat() {
  document.getElementById('btnChat').onclick = function () {
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

// ---------- 屏幕共享（5fps 1080p，共享文档足够且省带宽） ----------
let sharing = false;
function initScreenShare() {
  const btn = document.getElementById('btnShare');
  const big = document.getElementById('bigscreen');
  btn.onclick = async function () {
    if (!sharing) {
      try {
        await trtc.startScreenShare({
          view: 'screen-share-view',
          option: { profile: '1080p', qosPreference: TRTC.TYPE.QOS_PREFERENCE_SMOOTH }
        });
        sharing = true; big.dataset.sharer = myUserId; btn.innerText = '停止共享';
      } catch (e) { console.warn('[共享] 已取消或不支持:', e); }
    } else { await trtc.stopScreenShare(); }
  };
  trtc.on(TRTC.EVENT.SCREEN_SHARE_STOPPED, function () {
    sharing = false; big.style.display = 'none'; delete big.dataset.sharer; btn.innerText = '屏幕共享';
  });
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

// ---------- 会议文件（分片上传，单文件可达 30MB） ----------
function initFiles() {
  document.getElementById('btnFiles').onclick = function () {
    document.getElementById('fileModal').style.display = 'flex';
    refreshFileList();
  };
  document.getElementById('btnUploadFile').onclick = async function () {
    const f = document.getElementById('fileInput').files[0];
    if (!f) { alert('请先选择文件'); return; }
    if (f.size > 30 * 1024 * 1024) { alert('文件超过 30MB，请压缩后上传'); return; }
    try {
      if (f.size > CHUNK_SIZE) {
        const r = await uploadBigFile(f);
        showSubLine('【系统】文件上传成功（分片）：' + f.name);
        sendMsg(6, { t: 'file', name: f.name });
      } else {
        const b64 = await blobToB64(f);
        const up = await apiCall('meeting_file_upload', { filename: f.name, content_b64: b64 });
        showSubLine('【系统】文件上传成功：' + f.name);
        sendMsg(6, { t: 'file', name: f.name });
      }
      refreshFileList();
    } catch (e) { alert('上传失败：' + e.message); }
  };
  document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m) m.style.display = 'none'; }));
}
async function refreshFileList() {
  const box = document.getElementById('fileList');
  try {
    const d = await apiCall('meeting_file_list');
    const visible = (d.files || []).filter(f => { const p = parsePartName(f.name); return !p || p.i === 0; });
    if (!visible.length) { box.innerHTML = '<p style="color:var(--muted);font-size:13px;">暂无文件</p>'; return; }
    box.innerHTML = visible.map(f => {
      const p = parsePartName(f.name);
      const disp = p ? p.base : f.name;
      const payload = p ? { name: p.base, path: f.path, upid: p.upid, parts: p.n } : { name: f.name, path: f.path };
      return `<div class="file-item"><span>${escapeHtml(disp)}${p ? '（分片）' : ''}</span><button onclick='downloadAny(${JSON.stringify(payload)})'>下载</button></div>`;
    }).join('');
  } catch (e) { box.innerHTML = '<p style="color:#d98a6a;font-size:13px;">加载失败：' + e.message + '</p>'; }
}

// ---------- 协同批注 ----------
let boardCtx = null, boardTool = 'pen', drawing = false, curPts = [], lastCursorSent = 0;
function initBoard() {
  const modal = document.getElementById('boardModal');
  document.getElementById('btnBoard').onclick = () => { modal.style.display = 'flex'; resizeBoard(); };
  modal.addEventListener('click', e => { if (e.target === modal) modal.style.display = 'none'; });
  boardCtx = document.getElementById('boardCanvas').getContext('2d');
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
  const img = cv.toDataURL();
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
  applyDraw({ mode: boardTool === 'eraser' ? 'erase' : 'pen', color: document.getElementById('penColor').value, width: 3, pts: curPts.slice(-2) });
  if (curPts.length >= 40) { flushStroke(); curPts = [[p.x, p.y]]; }
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
  if (f.size > 2 * 1024 * 1024) { alert('底图请控制在 2MB 以内'); return; }
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
