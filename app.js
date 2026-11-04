// ==================== 配置区 ====================
const API_BASE = "https://service-xxx.gz.apigw.tencentcs.com/release"; // 你的腾讯云函数地址
const ROLE = "sushi"; // 默认角色

// ==================== 用户注册模块 ====================
function register() {
  const nickname = document.getElementById("nickname").value.trim();
  const grade = document.getElementById("grade").value;
  const interest = document.getElementById("interest").value.trim();

  if (!nickname) {
    alert("请填写你的称呼");
    return;
  }

  // 构建用户档案
  const userProfile = {
    nickname: nickname,
    grade: grade || "未填写",
    interest: interest || "未填写",
    registerTime: new Date().toISOString(),
    conversationCount: 0,
    memory: [] // 存储关键记忆点
  };

  // 存入本地存储
  localStorage.setItem("userProfile", JSON.stringify(userProfile));

  // 显示欢迎页
  document.getElementById("registerCard").classList.add("hidden");
  document.getElementById("welcomeCard").classList.remove("hidden");
  document.getElementById("welcomeText").innerHTML = 
    `你好，<strong>${nickname}</strong>。苏轼已准备好与你对话。`;

  // 3秒后自动跳转
  setTimeout(goChat, 1500);
}

function goChat() {
  window.location.href = "chat.html";
}

// ==================== 对话模块（含记忆） ====================
let conversationHistory = [];

function initChat() {
  // 检查是否注册
  const profile = localStorage.getItem("userProfile");
  if (!profile) {
    window.location.href = "index.html";
    return;
  }

  const user = JSON.parse(profile);
  
  // 显示用户信息
  document.getElementById("userInfo").innerHTML = 
    `${user.nickname} | ${user.grade}`;

  // 加载历史记忆
  if (user.memory && user.memory.length > 0) {
    conversationHistory = user.memory.slice(-10); // 最近10条
  }

  // 显示开场白（根据记忆个性化）
  const greeting = generateGreeting(user);
  addMessage("assistant", greeting);
}

function generateGreeting(user) {
  const hour = new Date().getHours();
  let timeGreeting = "";
  
  if (hour < 6) timeGreeting = "夜深了，";
  else if (hour < 12) timeGreeting = "晨安，";
  else if (hour < 18) timeGreeting = "午后好，";
  else timeGreeting = "傍晚好，";

  // 根据记忆生成个性化问候
  if (user.conversationCount > 5) {
    return `${timeGreeting}${user.nickname}。许久不见，别来无恙？今日有何想问老夫？`;
  } else if (user.interest && user.interest !== "未填写") {
    return `${timeGreeting}${user.nickname}。听闻你对${user.interest}颇有兴趣，我们不妨聊聊。`;
  } else {
    return `${timeGreeting}${user.nickname}。呵呵，今日想聊些什么？`;
  }
}

function addMessage(role, content) {
  const chatBox = document.getElementById("chatBox");
  const div = document.createElement("div");
  div.className = `message ${role}`;
  
  // 检查是否包含记忆标记
  let displayContent = content;
  let memoryTag = "";
  if (content.includes("[记忆]")) {
    displayContent = content.replace("[记忆]", "");
    memoryTag = `<div class="memory-tag">已记入记忆</div>`;
  }

  div.innerHTML = `
    <div class="bubble">
      ${escapeHtml(displayContent)}
      ${memoryTag}
    </div>
  `;
  
  chatBox.appendChild(div);
  chatBox.scrollTop = chatBox.scrollHeight;
}

async function sendMessage() {
  const input = document.getElementById("userInput");
  const text = input.value.trim();
  if (!text) return;

  // 显示用户消息
  addMessage("user", text);
  input.value = "";

  // 更新本地计数
  const profile = JSON.parse(localStorage.getItem("userProfile"));
  profile.conversationCount++;
  localStorage.setItem("userProfile", JSON.stringify(profile));

  // 构建发送历史（含记忆）
  const messages = conversationHistory.map(m => ({
    role: m.role,
    content: m.content
  }));
  messages.push({ role: "user", content: text });

  try {
    // 调用后端
    const response = await fetch(`${API_BASE}/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        role: ROLE,
        messages: messages,
        userProfile: {
          nickname: profile.nickname,
          grade: profile.grade,
          interest: profile.interest
        }
      })
    });

    const data = await response.json();
    
    // 显示回复
    addMessage("assistant", data.reply);

    // 更新记忆（如果后端返回了记忆点）
    if (data.memory) {
      conversationHistory.push(
        { role: "user", content: text },
        { role: "assistant", content: data.reply }
      );
      
      // 限制历史长度（防止 token 爆炸）
      if (conversationHistory.length > 20) {
        conversationHistory = conversationHistory.slice(-20);
      }

      // 保存到本地
      profile.memory = conversationHistory;
      localStorage.setItem("userProfile", JSON.stringify(profile));
    }

  } catch (error) {
    addMessage("assistant", "网络不畅，请稍后再试。");
    console.error("API Error:", error);
  }
}

// ==================== 工具函数 ====================
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
