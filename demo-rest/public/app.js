const $ = (sel) => document.querySelector(sel);

const loginPanel = $("#login-panel");
const chatPanel = $("#chat-panel");
const messagesEl = $("#messages");
const accountName = $("#account-name");
const providerBadge = $("#provider-badge");
const chatInput = $("#chat-input");

// ---- Boot ----
init();

async function init() {
  const status = await fetch("/api/status").then((r) => r.json());
  if (status.endpoint) {
    providerBadge.textContent = status.endpoint;
    providerBadge.hidden = false;
  }
  if (status.authenticated) showChat(status.account);
  else showLogin();
}

function showLogin() {
  loginPanel.hidden = false;
  loginPanel.style.display = "grid";
  chatPanel.hidden = true;
  chatPanel.style.display = "none";
}

function showChat(account) {
  loginPanel.hidden = true;
  loginPanel.style.display = "none";
  chatPanel.hidden = false;
  chatPanel.style.display = "flex";
  accountName.textContent = account ?? "";
  $("#new-btn").hidden = false;
  $("#logout-btn").hidden = false;
  chatInput.focus();
}

// ---- Sign in (device code) ----
$("#login-btn").addEventListener("click", startLogin);

async function startLogin() {
  const btn = $("#login-btn");
  const errEl = $("#login-error");
  errEl.hidden = true;
  btn.disabled = true;
  btn.textContent = "Getting sign-in code…";

  const flow = await fetch("/api/login", { method: "POST" }).then((r) => r.json());
  if (flow.state === "error") {
    errEl.textContent = flow.error;
    errEl.hidden = false;
    btn.disabled = false;
    btn.textContent = "Sign in";
    return;
  }

  $("#device-code").hidden = false;
  $("#device-uri").textContent = flow.verificationUri;
  $("#device-uri").href = flow.verificationUri;
  $("#device-user-code").textContent = flow.userCode;
  btn.textContent = "Waiting for sign-in…";
  pollLogin();
}

$("#copy-code").addEventListener("click", () => {
  navigator.clipboard.writeText($("#device-user-code").textContent);
  $("#copy-code").textContent = "Copied";
  setTimeout(() => ($("#copy-code").textContent = "Copy"), 1500);
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
      $("#login-btn").textContent = "Retry";
      $("#device-code").hidden = true;
    }
  }, 2500);
}

// ---- Chat ----
$("#chat-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const text = chatInput.value.trim();
  if (text) {
    chatInput.value = "";
    updateSendState();
    sendMessage(text);
  }
});

// Toggle the send button icon (waveform vs arrow) based on input content.
chatInput.addEventListener("input", updateSendState);
function updateSendState() {
  chatPanel.classList.toggle("can-send", chatInput.value.trim().length > 0);
}

// Suggestion chips
$("#suggestions").addEventListener("click", (e) => {
  const btn = e.target.closest(".suggestion");
  if (btn) sendMessage(btn.querySelector("span").textContent);
});

async function sendMessage(text) {
  chatPanel.classList.add("has-messages");

  addMessage(text, "user");
  const thinking = addMessage("Work IQ is thinking…", "bot", true);
  setSending(true);

  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    const data = await res.json();
    thinking.remove();
    if (!res.ok) addMessage("⚠️ " + (data.error || "An error occurred"), "bot");
    else addMessage(data.text, "bot");
  } catch (err) {
    thinking.remove();
    addMessage("⚠️ " + err.message, "bot");
  } finally {
    setSending(false);
    chatInput.focus();
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
  chatInput.disabled = on;
}

// ---- Header actions ----
$("#new-btn").addEventListener("click", async () => {
  await fetch("/api/new-conversation", { method: "POST" });
  messagesEl.innerHTML = "";
  chatPanel.classList.remove("has-messages", "can-send");
  chatInput.value = "";
  chatInput.focus();
});

$("#logout-btn").addEventListener("click", async () => {
  await fetch("/api/logout", { method: "POST" });
  location.reload();
});
