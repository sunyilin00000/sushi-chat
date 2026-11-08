// meeting-features.js
// 作用：会议业务功能——实时字幕 / 屏幕共享开关 / 匿名投票（后端聚合）/ 管理员强制禁言 / 会议文件
// 依赖：必须在 meeting-boot.js 之后引入（等 meeting-ready 事件后才操作 window.meeting）

// ★★★ 要改的唯一一处：管理员登录账号（发禁言/投票指令时，其他客户端靠这个名单识别管理员）★★★
const ADMIN_USERS = ['admin'];

const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";
let token = null, trtc = null, TRTC = null, myUserId = null, isAdmin = false;

// ---------- 工具函数 ----------
function apiCall(action, dataParams) {
  const body = { action, token };
  if (dataParams) body.data = dataParams;
  return fetch(TEAM_API, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    .then(r => r.json()).then(d => { if (d.error) throw new Error(d.error); return d; });
}
// 通过 TRTC 自定义消息广播一个 JSON 对象（cmdId 取值 1~10）
function sendMsg(cmdId, obj) {
  try {
    trtc.sendCustomMessage({ cmdId, data: new TextEncoder().encode(JSON.stringify(obj)).buffer });
  } catch (e) { console.warn('[信令] 发送失败:', e); }
}
// 字幕条显示一行字，8 秒后自动淡出
function showSubLine(text) {
  const bar = document.getElementById('subtitle-bar');
  const line = document.createElement('div');
  line.className = 'sub-line';
  line.textContent = text;
  bar.appendChild(line);
  setTimeout(() => line.remove(), 8200);
  while (bar.children.length > 3) bar.firstChild.remove();
}

// ---------- 等初始化完成 ----------
window.addEventListener('meeting-ready', function () {
  trtc = window.meeting.trtc;
  TRTC = window.meeting.TRTC;
  myUserId = window.meeting.userId;
  token = window.meeting.token;

  // 从本地登录信息判断是不是管理员（控制按钮显隐）
  try { isAdmin = (JSON.parse(localStorage.getItem('xh_user') || '{}').role === 'admin'); } catch (e) {}
  if (isAdmin) ['btnVote', 'btnMuteAll', 'btnUnmute'].forEach(id => document.getElementById(id).classList.remove('hidden'));

  initCustomMessages();
  initSubtitleButton();
  initScreenShare();
  initVote();
  initMuteButtons();
  initFiles();

  document.getElementById('btnLeave').onclick = () => window.meeting.leave();
});

// ---------- 自定义消息接收（所有功能的信令通道） ----------
function initCustomMessages() {
  trtc.on(TRTC.EVENT.CUSTOM_MESSAGE, function (event) {
    let msg = {};
    try { msg = JSON.parse(new TextDecoder().decode(event.data)); } catch (e) { return; }
    const from = event.userId;
    console.log('[信令]', msg.t, '来自', from);

    if (msg.t === 'subtitle') {
      showSubLine(msg.text);
    } else if (msg.t === 'vote_start') {
      openVoteModal(msg);
    } else if (msg.t === 'vote_end') {
      closeVoteModal();
      showVoteResult(msg.q, msg.opts, msg.counts);
    } else if (msg.t === 'mute_all') {
      if (ADMIN_USERS.includes(from)) {   // 只认管理员发的禁言指令
        window.meeting.muteLocalAudio();
        showSubLine('🔇 管理员已开启全体禁言');
      }
    } else if (msg.t === 'unmute_all') {
      if (ADMIN_USERS.includes(from)) {
        window.meeting.unmuteLocalAudio();
        showSubLine('🔊 管理员已解除全体禁言');
      }
    } else if (msg.t === 'file') {
      showSubLine('📎 ' + from + ' 上传了文件：' + msg.name);
      refreshFileList();
    }
  });
}

// ---------- 1. 实时字幕（浏览器自带语音识别，免费，无需后端） ----------
let recognition = null, subtitleOn = false;
function initSubtitleButton() {
  const btn = document.getElementById('btnSubtitle');
  btn.onclick = function () {
    if (subtitleOn) {
      recognition && recognition.stop();
      subtitleOn = false;
      btn.innerText = '字幕：关';
      return;
    }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { alert('当前浏览器不支持语音识别，请用 Chrome / Edge'); return; }
    recognition = new SR();
    recognition.lang = 'zh-CN';
    recognition.continuous = true;       // 持续识别
    recognition.interimResults = false;  // 只要最终结果，减少噪音
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
    subtitleOn = true;
    btn.innerText = '字幕：开';
  };
}

// ---------- 2. 屏幕共享 ----------
let sharing = false;
function initScreenShare() {
  const btn = document.getElementById('btnShare');
  const big = document.getElementById('bigscreen');
  btn.onclick = async function () {
    if (!sharing) {
      try {
        await trtc.startScreenShare({ view: 'screen-share-view' });  // 本地预览也放进大屏
        sharing = true;
        big.dataset.sharer = myUserId;
        btn.innerText = '停止共享';
      } catch (e) { console.warn('[共享] 已取消或不支持:', e); }
    } else {
      await trtc.stopScreenShare();
    }
  };
  // 用户点了浏览器自带的"停止共享"按钮时同步状态
  trtc.on(TRTC.EVENT.SCREEN_SHARE_STOPPED, function () {
    sharing = false;
    big.style.display = 'none';
    delete big.dataset.sharer;
    btn.innerText = '共享屏幕';
  });
}

// ---------- 3. 匿名投票（计数在后端聚合，任何人看不到谁投了啥） ----------
let currentVoteId = null, votePollTimer = null;
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
    currentVoteId = 'vote_' + Date.now();
    sendMsg(2, { t: 'vote_start', vote_id: currentVoteId, q, opts });
    // 管理员自己也要能看结果面板（管理员可不投，只看实时计数）
    document.getElementById('voteAdminCreate').classList.add('hidden');
    document.getElementById('voteAdminPanel').classList.remove('hidden');
    showVoteMask(q, '投票进行中…');
    // 每 2 秒拉一次后端计数以更新大屏遮罩
    votePollTimer = setInterval(() => pollVoteCount(q), 2000);
    pollVoteCount(q);
  };
  document.getElementById('btnSubmitVote').onclick = async function () {
    const checked = document.querySelector('input[name=voteOpt]:checked');
    if (!checked) { alert('请先选择一个选项'); return; }
    try {
      await apiCall('meeting_vote_cast', { vote_id: currentVoteId, choice: Number(checked.value) });
      closeVoteModal();
      showVoteMask(document.getElementById('voteQuestionText').innerText, '已提交，等待公布结果…');
      showSubLine('🗳️ 你已投票');
    } catch (e) { alert(e.message); }
  };
  document.getElementById('btnEndVote').onclick = async function () {
    try {
      const d = await apiCall('meeting_vote_status', { vote_id: currentVoteId });
      const msg = { t: 'vote_end', q: currentQ, opts: currentOpts, counts: d.counts };
      sendMsg(2, msg);
      clearInterval(votePollTimer);
      closeVoteModal();
      showVoteResult(msg.q, msg.opts, msg.counts);
    } catch (e) { alert(e.message); }
  };
}

let currentQ = '', currentOpts = [];
function pollVoteCount(q) {
  apiCall('meeting_vote_status', { vote_id: currentVoteId }).then(d => {
    const total = d.counts.reduce((a, b) => a + b, 0);
    document.getElementById('voteLiveCount').innerText =
      '已收到 ' + total + ' 票\n' + currentOpts.map((o, i) => o + '：' + (d.counts[i] || 0) + ' 票').join('\n');
    updateVoteMask(q, '投票进行中，已投 ' + total + ' 人');
  }).catch(() => {});
}
// 普通成员收到 vote_start：打开投票弹窗
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
function closeVoteModal() {
  document.getElementById('voteModal').style.display = 'none';
  clearInterval(votePollTimer);
}
// 大屏遮罩显示/更新/结果
function showVoteMask(q, sub) {
  const mask = document.getElementById('vote-mask');
  mask.style.display = 'flex';
  updateVoteMask(q, sub);
}
function updateVoteMask(q, sub) {
  document.getElementById('voteMaskContent').innerHTML =
    '<h2>🗳️ ' + q + '</h2><div class="mask-sub">' + sub + '<br>（投票期间，共享内容已遮罩）</div>';
}
function showVoteResult(q, opts, counts) {
  const total = counts.reduce((a, b) => a + b, 0);
  let html = '<h2>📊 投票结果：' + q + '</h2><div class="mask-sub">';
  opts.forEach((o, i) => {
    const n = counts[i] || 0, pct = total ? Math.round(n / total * 100) : 0;
    html += o + '：<b>' + n + '</b> 票（' + pct + '%）<br>';
  });
  html += '共 ' + total + ' 票</div>';
  document.getElementById('voteMaskContent').innerHTML = html;
  document.getElementById('vote-mask').style.display = 'flex';
  setTimeout(() => { document.getElementById('vote-mask').style.display = 'none'; }, 10000);
}

// ---------- 4. 管理员强制禁言（信令广播，各客户端自己闭麦） ----------
function initMuteButtons() {
  document.getElementById('btnMuteAll').onclick = function () {
    sendMsg(4, { t: 'mute_all' });
    window.meeting.muteRemoteAudio();   // 管理员自己也静音远端（不影响别人）
    showSubLine('🔇 你已开启全体禁言');
  };
  document.getElementById('btnUnmute').onclick = function () {
    sendMsg(5, { t: 'unmute_all' });
    window.meeting.unmuteRemoteAudio();
    showSubLine('🔊 你已解除全体禁言');
  };
}

// ---------- 5. 会议文件（存 xinghuan-data 仓库 files/ 目录，单文件 ≤700KB） ----------
function initFiles() {
  document.getElementById('btnFiles').onclick = function () {
    document.getElementById('fileModal').style.display = 'flex';
    refreshFileList();
  };
  document.getElementById('btnUploadFile').onclick = async function () {
    const f = document.getElementById('fileInput').files[0];
    if (!f) { alert('请先选择文件'); return; }
    if (f.size > 700 * 1024) { alert('文件超过 700KB，GitHub 存储接口放不下，请压缩后上传'); return; }
    const b64 = await new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result.split(',')[1]); r.readAsDataURL(f); });
    try {
      await apiCall('meeting_file_upload', { filename: f.name, content_b64: b64 });
      showSubLine('📎 文件上传成功：' + f.name);
      sendMsg(6, { t: 'file', name: f.name });   // 通知其他人刷新列表
      refreshFileList();
    } catch (e) { alert('上传失败: ' + e.message); }
  };
  // 点击遮罩空白处关闭弹窗
  document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => { if (e.target === m) m.style.display = 'none'; }));
}
async function refreshFileList() {
  const box = document.getElementById('fileList');
  try {
    const d = await apiCall('meeting_file_list');
    if (!d.files || d.files.length === 0) { box.innerHTML = '<p style="color:#94a3b8;font-size:13px;">暂无文件</p>'; return; }
    box.innerHTML = d.files.map(f =>
      `<div class="file-item"><span>${f.name}（${Math.round(f.size / 1024)}KB）</span><button onclick="downloadFile('${f.path}', '${f.name}')">下载</button></div>`).join('');
  } catch (e) { box.innerHTML = '<p style="color:#ef4444;font-size:13px;">加载失败: ' + e.message + '</p>'; }
}
async function downloadFile(path, name) {
  try {
    const d = await apiCall('meeting_file_download', { path });
    const bin = atob(d.content_b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    const blob = new Blob([arr]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
  } catch (e) { alert('下载失败: ' + e.message); }
}
