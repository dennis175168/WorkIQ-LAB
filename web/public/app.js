const $ = (sel) => document.querySelector(sel);

const loginPanel = $("#login-panel");
const chatPanel = $("#chat-panel");
const messagesEl = $("#messages");
const accountName = $("#account-name");
const endpointBadge = $("#endpoint-badge");

// ---- Boot ----
init();

async function init() {
  const status = await fetch("/api/status").then((r) => r.json());
  endpointBadge.textContent = status.endpoint;
  if (status.authenticated) {
    showChat(status.account);
  } else {
    showLogin();
  }
}

function showLogin() {
  loginPanel.hidden = false;
  loginPanel.style.display = "grid";
  chatPanel.hidden = true;
  chatPanel.style.display = "none";
}

// ---- Sign in (device code) ----
$("#login-btn").addEventListener("click", startLogin);

async function startLogin() {
  const btn = $("#login-btn");
  const errEl = $("#login-error");
  errEl.hidden = true;
  btn.disabled = true;
  btn.textContent = "取得登入代碼中…";

  const flow = await fetch("/api/login", { method: "POST" }).then((r) => r.json());
  if (flow.state === "error") {
    errEl.textContent = flow.error;
    errEl.hidden = false;
    btn.disabled = false;
    btn.textContent = "登入";
    return;
  }

  // Show the device code
  $("#device-code").hidden = false;
  $("#device-uri").textContent = flow.verificationUri;
  $("#device-uri").href = flow.verificationUri;
  $("#device-user-code").textContent = flow.userCode;
  btn.textContent = "等待登入完成…";
  pollLogin();
}

$("#copy-code").addEventListener("click", () => {
  navigator.clipboard.writeText($("#device-user-code").textContent);
  $("#copy-code").textContent = "已複製";
  setTimeout(() => ($("#copy-code").textContent = "複製"), 1500);
});

async function pollLogin() {
  const errEl = $("#login-error");
  const timer = setInterval(async () => {
    const flow = await fetch("/api/login/poll").then((r) => r.json());
    if (flow.state === "done") {
      clearInterval(timer);
      const status = await fetch("/api/status").then((r) => r.json());
      showChat(status.account);
    } else if (flow.state === "error") {
      clearInterval(timer);
      errEl.textContent = flow.error;
      errEl.hidden = false;
      $("#login-btn").disabled = false;
      $("#login-btn").textContent = "重試";
      $("#device-code").hidden = true;
    }
  }, 2500);
}

// ---- Chat ----
function showChat(account) {
  loginPanel.hidden = true;
  loginPanel.style.display = "none";
  chatPanel.hidden = false;
  chatPanel.style.display = "flex";
  accountName.textContent = account ?? "";
  $("#new-btn").hidden = false;
  $("#logout-btn").hidden = false;
  $("#chat-input").focus();
}

$("#chat-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const input = $("#chat-input");
  const text = input.value.trim();
  if (text) {
    input.value = "";
    sendMessage(text);
  }
});

messagesEl.addEventListener("click", (e) => {
  if (e.target.classList.contains("chip")) {
    sendMessage(e.target.textContent);
  }
});

async function sendMessage(text) {
  const hint = messagesEl.querySelector(".hint");
  if (hint) hint.remove();

  addMessage(text, "user");
  const thinking = addMessage("Work IQ 思考中…", "bot", true);
  setSending(true);

  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, webSearch: $("#web-search").checked }),
    });
    const data = await res.json();
    thinking.remove();
    if (!res.ok) {
      addMessage("⚠️ " + (data.error || "發生錯誤"), "bot");
    } else {
      addMessage(data.text, "bot");
    }
  } catch (err) {
    thinking.remove();
    addMessage("⚠️ " + err.message, "bot");
  } finally {
    setSending(false);
    $("#chat-input").focus();
  }
}

function addMessage(text, role, thinking = false) {
  const el = document.createElement("div");
  el.className = `msg ${role}` + (thinking ? " thinking" : "");
  if (role === "bot" && !thinking) {
    el.innerHTML = marked.parse(text);
    el.querySelectorAll("a").forEach((a) => (a.target = "_blank"));
  } else {
    el.textContent = text;
  }
  messagesEl.appendChild(el);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return el;
}

function setSending(on) {
  $("#send-btn").disabled = on;
  $("#chat-input").disabled = on;
}

// ---- Header actions ----
$("#new-btn").addEventListener("click", async () => {
  await fetch("/api/new-conversation", { method: "POST" });
  messagesEl.innerHTML =
    '<div class="hint">已開始新對話。試著問問看您的工作內容吧！</div>';
});

$("#logout-btn").addEventListener("click", async () => {
  await fetch("/api/logout", { method: "POST" });
  location.reload();
});
