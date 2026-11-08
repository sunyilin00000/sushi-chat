// meeting-boot.js
// 会议初始化——登录检查 → 会议密钥 → 加载SDK → 拿UserSig → 进房 → 开摄像头麦(480p+双流) → 广播 ready
document.addEventListener('DOMContentLoaded', async function () {
  const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";

  try {
    const token = localStorage.getItem('xh_token');
    if (!token) { location.href = 'member.html'; return; }

    const roomId = new URLSearchParams(location.search).get('room') || '101';
    const roomKey = prompt('请输入房间 ' + roomId + ' 的会议密钥（未设置则留空直接确定）', '') || '';

    const TRTC = await window.loadTRTC();
    if (!(await TRTC.isSupported())) {
      alert('当前浏览器不支持 WebRTC，请使用最新版 Chrome 或 Edge');
      return;
    }

    const resp = await fetch(TEAM_API, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'get_trtc_sig', token: token, data: { room: roomId, room_key: roomKey } })
    });
    if (!resp.ok) throw new Error((resp.status === 403) ? '会议密钥错误，请刷新页面重新输入' : '获取 UserSig 失败: HTTP ' + resp.status);
    const sigData = await resp.json();
    if (sigData.error) throw new Error('获取 UserSig 失败: ' + sigData.error);

    const roomName = sigData.roomName || '';
    if (roomName) {
      document.getElementById('meetingName').innerText = roomName;
      document.title = roomName + ' · 星环终端';
    }

    const trtc = TRTC.create();

    // 远端用户进房：建格子（含单人禁言按钮，默认隐藏，管理员身份在 features 里显示）
    trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, function (event) {
      const wall = document.getElementById('remote-videos');
      if (wall && !document.getElementById('cell-' + event.userId)) {
        const cell = document.createElement('div');
        cell.className = 'remote-cell';
        cell.id = 'cell-' + event.userId;
        cell.innerHTML = '<button class="mute-btn" data-uid="' + event.userId + '" style="display:none">禁言</button>';
        wall.appendChild(cell);
      }
    });

    // 双流：默认订阅小流；点击格子切大流
    const bigStreamUsers = new Set();
    trtc.on(TRTC.EVENT.REMOTE_VIDEO_AVAILABLE, function (event) {
      if (event.streamType === TRTC.TYPE.STREAM_TYPE_SUB) {
        const big = document.getElementById('bigscreen');
        big.style.display = 'block';
        big.dataset.sharer = event.userId;
        trtc.startRemoteVideo({ userId: event.userId, streamType: event.streamType, view: 'screen-share-view', option: { fillMode: 'contain' } });
      } else {
        trtc.startRemoteVideo({
          userId: event.userId, streamType: event.streamType, view: 'cell-' + event.userId,
          option: { small: !bigStreamUsers.has(event.userId) }   // ★ 默认小流
        });
      }
    });

    trtc.on(TRTC.EVENT.REMOTE_VIDEO_UNAVAILABLE, function (event) {
      if (event.streamType === TRTC.TYPE.STREAM_TYPE_SUB) {
        const big = document.getElementById('bigscreen');
        if (big.dataset.sharer === event.userId) big.style.display = 'none';
      }
    });

    trtc.on(TRTC.EVENT.REMOTE_USER_LEAVE, function (event) {
      const cell = document.getElementById('cell-' + event.userId);
      if (cell) cell.remove();
      bigStreamUsers.delete(event.userId);
      const big = document.getElementById('bigscreen');
      if (big.dataset.sharer === event.userId) big.style.display = 'none';
    });

    trtc.on(TRTC.EVENT.KICKED_OUT, function () {
      alert('你已被请出房间（可能在其他设备用同一账号进入了会议）');
      location.href = 'member.html';
    });

    await trtc.enterRoom({
      sdkAppId: sigData.sdkAppId,
      userId: sigData.userId,
      userSig: sigData.userSig,
      roomId: Number(roomId)
    });

    // ★ 480p 推流 + 同时推小流（供他人低带宽观看）
    try {
      await trtc.startLocalVideo({
        view: 'local-video',
        option: { profile: '480p', small: true, qosPreference: TRTC.TYPE.QOS_PREFERENCE_SMOOTH }
      });
      await trtc.startLocalAudio();
    } catch (e) {
      alert('摄像头/麦克风授权被拒绝，无法加入会议');
      await trtc.exitRoom();
      return;
    }

    window.meeting = {
      trtc: trtc, TRTC: TRTC,
      userId: sigData.userId, roomId: Number(roomId),
      TEAM_API: TEAM_API, token: token,
      muteRemoteAudio: function () { trtc.muteRemoteAudio('*', true); },
      unmuteRemoteAudio: function () { trtc.muteRemoteAudio('*', false); },
      muteLocalAudio: function () { trtc.updateLocalAudio({ mute: true }); },
      unmuteLocalAudio: function () { trtc.updateLocalAudio({ mute: false }); },
      leave: async function () { await trtc.exitRoom(); location.href = 'member.html'; }
    };

    // ★ 点击成员格子：切换该成员大流/小流
    document.getElementById('remote-videos').addEventListener('click', async function (e) {
      const btn = e.target.closest('.mute-btn');
      if (btn) return;   // 禁言按钮由 features 处理
      const cell = e.target.closest('.remote-cell');
      if (!cell) return;
      const uid = cell.id.replace('cell-', '');
      try {
        if (bigStreamUsers.has(uid)) {
          bigStreamUsers.delete(uid);
          await trtc.updateRemoteVideo({ userId: uid, streamType: TRTC.TYPE.STREAM_TYPE_MAIN, option: { small: true } });
        } else {
          bigStreamUsers.add(uid);
          await trtc.updateRemoteVideo({ userId: uid, streamType: TRTC.TYPE.STREAM_TYPE_MAIN, option: { small: false } });
        }
      } catch (err) { console.warn('[双流切换]', err); }
    });

    console.log('[会议] 初始化完成');
    window.dispatchEvent(new Event('meeting-ready'));
  } catch (err) {
    console.error('[会议] 初始化失败:', err);
    alert('会议初始化失败: ' + err.message);
  }
});
