// meeting-boot.js
// 作用：会议页完整初始化流程——检查登录 → 容错加载SDK → 向后端要UserSig → 进房 → 开摄像头/麦克风
document.addEventListener('DOMContentLoaded', async function () {
  // teamApi 接口地址（与 member.html 里的 TEAM_API 保持一致）
  const TEAM_API = "https://1502956590-0kdly2qr1v.ap-nanjing.tencentscf.com";

  try {
    // ============ 第 0 步：检查是否已登录（member.html 登录后存 xh_token） ============
    const token = localStorage.getItem('xh_token');
    if (!token) {
      // 没登录就直接回工作台登录
      location.href = 'member.html';
      return;
    }

    // ============ 第 1 步：容错加载 SDK（多源自动切换，本地优先） ============
    const TRTC = await window.loadTRTC();

    // ============ 第 2 步：浏览器兼容性检查 ============
    if (!(await TRTC.isSupported())) {
      alert('当前浏览器不支持 WebRTC，请使用最新版 Chrome 或 Edge');
      return;
    }

    // ============ 第 3 步：从 teamApi 获取 UserSig ============
    // 后端 action = get_trtc_sig，必须带 token 鉴权
    // 返回 { sdkAppId, userId, userSig }，userId 必须用后端返回的（它和签名是绑定的）
    const resp = await fetch(TEAM_API, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'get_trtc_sig', token: token })
    });
    if (!resp.ok) throw new Error('获取 UserSig 失败: HTTP ' + resp.status);
    const sigData = await resp.json();
    if (sigData.error) throw new Error('获取 UserSig 失败: ' + sigData.error);

    // ============ 第 4 步：创建 TRTC 实例、注册远端用户事件 ============
    const trtc = TRTC.create();

    // 远端用户进房
    trtc.on(TRTC.EVENT.REMOTE_USER_ENTER, function (event) {
      console.log('[会议] 远端用户进房:', event.userId);
    });

    // 远端用户发布音视频：订阅并播放（页面里需要有 <div id="video-远端用户ID"></div>）
    trtc.on(TRTC.EVENT.REMOTE_VIDEO_AVAILABLE, function (event) {
      trtc.startRemoteVideo({
        userId: event.userId,
        streamType: event.streamType,
        view: 'video-' + event.userId
      });
    });

    // 远端用户退房
    trtc.on(TRTC.EVENT.REMOTE_USER_LEAVE, function (event) {
      console.log('[会议] 远端用户退房:', event.userId);
    });

    // ============ 第 5 步：进房（房间号从网址读，如 meeting.html?room=101） ============
    const roomId = new URLSearchParams(location.search).get('room') || '101';
    await trtc.enterRoom({
      sdkAppId: sigData.sdkAppId,
      userId: sigData.userId,   // 必须用后端返回的，不能自己拼
      userSig: sigData.userSig,
      roomId: Number(roomId)
    });

    // ============ 第 6 步：强制申请摄像头/麦克风（你的需求：强制授权） ============
    try {
      await trtc.startLocalVideo({ view: 'local-video' }); // 页面需有 <div id="local-video">
      await trtc.startLocalAudio();
    } catch (e) {
      alert('摄像头/麦克风授权被拒绝，无法加入会议，请刷新页面并重试');
      await trtc.exitRoom();
      return;
    }

    // ============ 第 7 步：暴露控制接口给后续业务代码（投票/字幕/禁言会用） ============
    window.meeting = {
      trtc: trtc,
      userId: sigData.userId,
      roomId: Number(roomId),
      muteAllRemoteAudio: function () { trtc.muteAllRemoteAudio(); },     // 管理员一键全体禁言（本端静音所有远端）
      unmuteAllRemoteAudio: function () { trtc.unmuteAllRemoteAudio(); }, // 解除
      stopLocalAudio: function () { trtc.stopLocalAudio(); },             // 被禁言的客户端执行
      startLocalAudio: function () { trtc.startLocalAudio(); },
      leave: async function () { await trtc.exitRoom(); location.href = 'member.html'; }
    };
    console.log('[会议] 初始化完成，window.meeting 已可用');
  } catch (err) {
    console.error('[会议] 初始化失败:', err);
    alert('会议初始化失败: ' + err.message);
  }
});
