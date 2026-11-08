// trtc-loader.js
// 作用：按优先级依次尝试多个 SDK 源，任一成功即返回全局 TRTC 对象。
// 每个源都有超时保护——"被屏蔽"的域名通常不是报错而是永远挂起，必须靠超时跳过。
(function (global) {
  'use strict';

  // 候选源列表，按顺序尝试。
  // 本地文件排第一：GitHub Pages 页面能打开，同源文件就 100% 能加载，这是最可靠的源。
  var SOURCES = [
    './trtc.js',                                                          // 兜底：仓库根目录本地副本
    'https://cdn.jsdelivr.net/npm/trtc-sdk-v5@5.20.2/trtc.js',            // jsDelivr（npm 源）
    'https://unpkg.com/trtc-sdk-v5@5.20.2/trtc.js',                       // unpkg（npm 源）
    'https://cdn.jsdelivr.net/gh/sunyilin00000/sushi-chat@main/trtc.js',  // jsDelivr（GitHub 源）
    'https://web.sdk.qcloud.com/trtc/webrtc/v5/dist/trtc.js'              // 腾讯云官方 CDN
  ];

  var TIMEOUT_MS = 8000; // 单个源最长等待 8 秒，防挂起

  // 加载单个源，返回 Promise；只有 window.TRTC 真的出现才算成功
  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var script = document.createElement('script');
      var timer = setTimeout(function () {
        script.src = '';           // 超时：主动放弃这个源的加载
        reject(new Error('timeout'));
      }, TIMEOUT_MS);

      script.onload = function () {
        clearTimeout(timer);
        if (global.TRTC) {
          resolve(src);
        } else {
          reject(new Error('loaded but TRTC undefined')); // 文件下载了但不是 SDK
        }
      };
      script.onerror = function () {
        clearTimeout(timer);
        reject(new Error('network error')); // 404 / 断网
      };

      script.src = src;
      script.async = true;
      document.head.appendChild(script);
    });
  }

  // 对外暴露的唯一入口：window.loadTRTC()，返回 Promise<TRTC>
  global.loadTRTC = function () {
    var chain = Promise.reject(new Error('start'));
    SOURCES.forEach(function (src) {
      chain = chain.catch(function () {
        console.warn('[TRTC] 尝试加载源:', src);
        return loadScript(src);
      });
    });
    return chain.then(function (src) {
      console.log('[TRTC] 加载成功，来源:', src);
      return global.TRTC;
    });
  };
})(window);
