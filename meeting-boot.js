// meeting-boot.js
// 作用：会议初始化主流程——登录检查 → 加载SDK → 拿UserSig → 进房 → 开摄像头麦 → 广播 ready 事件
// 本文件只管"进房和流"，业务功能（投票/禁言/字幕/文件/共享）在 meeting-features.js
document.addEventListener('DOMContentLoaded', async function () {
  const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";

  try {
    // 检查登录
    const token = localStorage.getItem('xh_token');
    if (!token) { location.href = 'member.html'; return; }

    // 容错加载 SDK
    const TRTC = await window.loadTRTC();
    if (!(await TRTC.isSupported())) {
      alert('当前浏览器不支持 WebRTC，请使用最新版 Chrome 或 Edge');
      return;
    }

    // 获取 UserSig
    const resp = await fetch(TEAM_API, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'get_trtc_sig', token: token })
    });
    if (!resp.ok) throw new Error('获取 UserSig 失败: HTTP ' + resp.status);
    const sigData = await resp.json();
    if (sigData.error) throw new Error('获取 UserSig 失败: ' + sigData.error);

    const trtc = TRTC.create();

    // 远端用户进房：普通摄像头流建格子；屏幕共享流由 VIDEO_AVAILABLE 事件处理
    trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, function (event) {
      const wall = document.getElementById('remote-videos');
      if (wall && !document.getElementById('cell-' + event.userId)) {
        const cell = document.createElement('div');
        cell.className = 'remote-cell';
        cell.id = 'cell-' + event.userId;
        wall.appendChild(cell);
      }
    });

    // ★ 核心：区分普通视频流和屏幕共享流（sub 流）
    trtc.on(TRTC.EVENT.REMOTE_VIDEO_AVAILABLE, function (event) {
      if (event.streamType === TRTC.TYPE.STREAM_TYPE_SUB) {
        // 屏幕共享：渲染到大屏
        const big = document.getElementById('bigscreen');
        big.style.display = 'block';
        big.dataset.sharer = event.userId;
        trtc.startRemoteVideo({
          userId: event.userId,
          streamType: event.streamType,
          view: 'screen-share-view'
        });
      } else {
        // 普通摄像头：渲染到成员格子
        trtc.startRemoteVideo({
          userId: event.userId,
          streamType: event.streamType,
          view: 'cell-' + event.userId
        });
      }
    });

    // 远端停止屏幕共享：收起大屏
    trtc.on(TRTC.EVENT.REMOTE_VIDEO_UNAVAILABLE, function (event) {
      if (event.streamType === TRTC.TYPE.STREAM_TYPE_SUB) {
        const big = document.getElementById('bigscreen');
        if (big.dataset.sharer === event.userId) big.style.display = 'none';
      }
    });

    // 远端退房：删格子 + 若是共享者则收大屏
    trtc.on(TRTC.EVENT.REMOTE_USER_LEAVE, function (event) {
      const cell = document.getElementById('cell-' + event.userId);
      if (cell) cell.remove();
      const big = document.getElementById('bigscreen');
      if (big.dataset.sharer === event.userId) big.style.display = 'none';
    });

    // ★ 被踢下线（同账号在别处进房等）：给出提示并返回工作台
    trtc.on(TRTC.EVENT.KICKED_OUT, function (event) {
      alert('你已被请出房间（可能在其他设备用同一账号进入了会议）');
      location.href = 'member.html';
    });

    // 进房
    const roomId = new URLSearchParams(location.search).get('room') || '101';
    await trtc.enterRoom({
      sdkAppId: sigData.sdkAppId,
      userId: sigData.userId,
      userSig: sigData.userSig,
      roomId: Number(roomId)
    });

    // 强制摄像头/麦克风
    try {
      await trtc.startLocalVideo({ view: 'local-video' });
      await trtc.startLocalAudio();
    } catch (e) {
      alert('摄像头/麦克风授权被拒绝，无法加入会议');
      await trtc.exitRoom();
      return;
    }

    // 暴露控制接口
    window.meeting = {
      trtc: trtc,
      TRTC: TRTC,
      userId: sigData.userId,
      roomId: Number(roomId),
      TEAM_API: TEAM_API,
      token: token,
      // 本端静音所有远端声音（只影响自己听）
      muteRemoteAudio: function () { trtc.muteRemoteAudio('*', true); },
      unmuteRemoteAudio: function () { trtc.muteRemoteAudio('*', false); },
      // 本端麦克风静音/恢复（被管理员禁言时调用）
      muteLocalAudio: function () { trtc.updateLocalAudio({ mute: true }); },
      unmuteLocalAudio: function () { trtc.updateLocalAudio({ mute: false }); },
      leave: async function () { await trtc.exitRoom(); location.href = 'member.html'; }
    };
    console.log('[会议] 初始化完成');
    // 通知业务功能模块可以开工了
    window.dispatchEvent(new Event('meeting-ready'));
  } catch (err) {
    console.error('[会议] 初始化失败:', err);
    alert('会议初始化失败: ' + err.message);
  }
});
