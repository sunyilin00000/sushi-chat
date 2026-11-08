// meeting-boot.js
// 会议初始化——登录检查 → 会议密钥 → 加载SDK → 校验并拿UserSig → 进房 → 开摄像头麦 → 广播 ready
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

    // ★ 显示会议名称（管理台设置，经 get_trtc_sig 返回）
    const roomName = sigData.roomName || '';
    if (roomName) {
      document.getElementById('meetingName').innerText = roomName;
      document.title = roomName + ' · 星环终端';
    }

    const trtc = TRTC.create();

    trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, function (event) {
      const wall = document.getElementById('remote-videos');
      if (wall && !document.getElementById('cell-' + event.userId)) {
        const cell = document.createElement('div');
        cell.className = 'remote-cell';
        cell.id = 'cell-' + event.userId;
        wall.appendChild(cell);
      }
    });

    trtc.on(TRTC.EVENT.REMOTE_VIDEO_AVAILABLE, function (event) {
      if (event.streamType === TRTC.TYPE.STREAM_TYPE_SUB) {
        const big = document.getElementById('bigscreen');
        big.style.display = 'block';
        big.dataset.sharer = event.userId;
        trtc.startRemoteVideo({ userId: event.userId, streamType: event.streamType, view: 'screen-share-view' });
      } else {
        trtc.startRemoteVideo({ userId: event.userId, streamType: event.streamType, view: 'cell-' + event.userId });
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

    try {
      await trtc.startLocalVideo({ view: 'local-video' });
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
    console.log('[会议] 初始化完成');
    window.dispatchEvent(new Event('meeting-ready'));
  } catch (err) {
    console.error('[会议] 初始化失败:', err);
    alert('会议初始化失败: ' + err.message);
  }
});
