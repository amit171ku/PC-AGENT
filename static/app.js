const $ = (id) => document.getElementById(id);
const chat = $("chat");
const welcome = $("welcome");
const input = $("input");
const sendBtn = $("send");
const working = $("working");
const workingText = document.querySelector("#working span:last-child");
const dot = $("dot");
const statusText = $("statusText");
const meta = $("meta");
const themeBtn = $("theme");
const sidebar = $("sidebar");
const scrim = $("scrim");
const filesEl = $("files");
const fileInput = $("fileInput");
const dropEl = $("drop");
let ws;

// ---------- Theme ----------
function applyTheme(t) {
  document.documentElement.setAttribute("data-theme", t);
  themeBtn.textContent = t === "dark" ? "Light mode" : "Dark mode";
}
let savedTheme = null;
try { savedTheme = localStorage.getItem("theme"); } catch (e) {}
applyTheme(savedTheme || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"));

themeBtn.onclick = () => {
  const next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
  applyTheme(next);
  try { localStorage.setItem("theme", next); } catch (e) {}
};

// ---------- Sidebar (mobile) ----------
function toggleSidebar(open) {
  sidebar.classList.toggle("open", open);
  scrim.classList.toggle("show", open);
}
$("menu").onclick = () => toggleSidebar(true);
scrim.onclick = () => toggleSidebar(false);

// ---------- Helpers ----------
function hideWelcome() { welcome.classList.add("hidden"); }
function scrollDown() { chat.scrollTop = chat.scrollHeight; }
function timeNow() {
  return new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

// SVG icon from the sprite in index.html (no HTML injection)
function icon(id, cls) {
  const NS = "http://www.w3.org/2000/svg";
  const s = document.createElementNS(NS, "svg");
  s.setAttribute("class", "ico" + (cls ? " " + cls : ""));
  s.setAttribute("aria-hidden", "true");
  const u = document.createElementNS(NS, "use");
  u.setAttribute("href", "#i-" + id);
  s.appendChild(u);
  return s;
}

// Safe lightweight formatting: code blocks, `inline code`, **bold**.
// Everything is inserted as text nodes, never as HTML.
function inline(el, text) {
  text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g).forEach((tok) => {
    if (!tok) return;
    if (tok.length > 4 && tok.startsWith("**") && tok.endsWith("**")) {
      const b = document.createElement("strong");
      b.textContent = tok.slice(2, -2);
      el.appendChild(b);
    } else if (tok.length > 2 && tok.startsWith("`") && tok.endsWith("`")) {
      const c = document.createElement("code");
      c.className = "inline";
      c.textContent = tok.slice(1, -1);
      el.appendChild(c);
    } else {
      el.appendChild(document.createTextNode(tok));
    }
  });
}

function renderRich(el, text) {
  text.split(/```(?:\w+)?\n?([\s\S]*?)```/g).forEach((part, i) => {
    if (i % 2 === 1) {
      const pre = document.createElement("pre");
      const code = document.createElement("code");
      code.textContent = part.replace(/\n$/, "");
      pre.appendChild(code);
      el.appendChild(pre);
    } else if (part.trim()) {
      const p = document.createElement("div");
      p.className = "para";
      inline(p, part.trim());
      el.appendChild(p);
    }
  });
}

// ---------- Messages ----------
function addMessage(text, kind) {
  hideWelcome();
  const row = document.createElement("div");
  row.className = "msg " + kind;

  if (kind !== "user") {
    const av = document.createElement("div");
    av.className = "avatar";
    av.appendChild(icon("bot"));
    row.appendChild(av);
  }

  const wrap = document.createElement("div");
  wrap.className = "body-wrap";

  const bubble = document.createElement("div");
  bubble.className = "bubble";
  if (kind === "agent" && typeof window.renderMarkdown === "function") {
    window.renderMarkdown(bubble, text);
} else if (kind === "agent") {
    renderRich(bubble, text);
} else {
    bubble.textContent = text;
}
  wrap.appendChild(bubble);

  // Download buttons for any files from reports/ mentioned in the reply
  if (kind === "agent") {
    const found = [...new Set(text.match(/reports\/[^\s`*"'()]+\.(?:docx|xlsx|pptx|pdf)/g) || [])];
    if (found.length) {
      const dls = document.createElement("div");
      dls.className = "downloads";
      found.forEach((p) => {
        const a = document.createElement("a");
        a.className = "dl";
        a.href = "/download?path=" + encodeURIComponent(p);
        a.setAttribute("download", "");
        a.textContent = "Download " + p.split("/").pop();
        dls.appendChild(a);
      });
      wrap.appendChild(dls);
    }
  }

  const metaRow = document.createElement("div");
  metaRow.className = "meta-row";
  const t = document.createElement("span");
  t.textContent = timeNow();
  metaRow.appendChild(t);

  if (kind === "agent") {
    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "copy";
    copy.textContent = "Copy";
    copy.onclick = () => {
      if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(() => {
          copy.textContent = "Copied";
          setTimeout(() => (copy.textContent = "Copy"), 1500);
        });
      }
    };
    metaRow.appendChild(copy);
  }

  wrap.appendChild(metaRow);
  row.appendChild(wrap);
  chat.appendChild(row);
  scrollDown();
}

function addNote(text) {
  hideWelcome();
  const n = document.createElement("div");
  n.className = "note";
  n.textContent = text;
  chat.appendChild(n);
  scrollDown();
}

function addConfirm(id, text) {
  hideWelcome();
  const box = document.createElement("div");
  box.className = "confirm";

  const head = document.createElement("div");
  head.className = "c-head";
  head.textContent = "Approval needed";

  const body = document.createElement("div");
  body.className = "c-body";
  body.textContent = text;

  const actions = document.createElement("div");
  actions.className = "c-actions";

  const yes = document.createElement("button");
  yes.type = "button";
  yes.className = "btn primary";
  yes.textContent = "Allow";

  const no = document.createElement("button");
  no.type = "button";
  no.className = "btn";
  no.textContent = "Deny";

  function answer(ok) {
    if (ws && ws.readyState === 1) {
      ws.send(JSON.stringify({ type: "confirm", id: id, ok: ok }));
    }
    actions.remove();
    head.textContent = ok ? "Allowed" : "Denied";
    box.classList.add("done");
  }
  yes.onclick = () => answer(true);
  no.onclick = () => answer(false);

  actions.append(yes, no);
  box.append(head, body, actions);
  chat.appendChild(box);
  scrollDown();
}

// Question card: the agent asks, you type the answer
function addAsk(id, text) {
  hideWelcome();
  const box = document.createElement("div");
  box.className = "ask";

  const head = document.createElement("div");
  head.className = "a-head";
  head.textContent = "The agent needs some details";

  const body = document.createElement("div");
  body.className = "a-body";
  body.textContent = text;

  const ta = document.createElement("textarea");
  ta.className = "a-input";
  ta.rows = 4;
  ta.placeholder = "Type your answers here. Number them to match the questions.";

  const actions = document.createElement("div");
  actions.className = "c-actions";

  const send = document.createElement("button");
  send.type = "button";
  send.className = "btn primary";
  send.textContent = "Send answers";

  const skip = document.createElement("button");
  skip.type = "button";
  skip.className = "btn";
  skip.textContent = "Skip (use defaults)";

  function reply(txt) {
    if (ws && ws.readyState === 1) {
      ws.send(JSON.stringify({ type: "answer", id: id, text: txt }));
    }
    ta.remove();
    actions.remove();
    head.textContent = txt ? "Answered" : "Skipped";
    box.classList.add("done");
    if (txt) addMessage(txt, "user");
  }
  send.onclick = () => reply(ta.value.trim());
  skip.onclick = () => reply("");
  ta.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      reply(ta.value.trim());
    }
  });

  actions.append(send, skip);
  box.append(head, body, ta, actions);
  chat.appendChild(box);
  scrollDown();
  ta.focus();
}

// ---------- Busy state + friendly status ----------
function setBusy(busy) {
  sendBtn.disabled = busy;
  working.classList.toggle("hidden", !busy);
  if (busy && workingText) workingText.textContent = "Agent is working...";
}

function statusFor(text) {
  const name = text.split("(")[0];
  const map = {
    search_web: "Searching the web...",
    open_url: "Opening the page...",
    get_page_text: "Reading the page...",
    click: "Clicking on the page...",
    type_text: "Typing on the page...",
    list_dir: "Looking at your files...",
    read_file: "Reading a file...",
    read_document: "Reading your template...",
    search_files: "Searching your files...",
    write_file: "Writing a file...",
    move_file: "Moving a file...",
    delete_file: "Removing a file...",
    ask_user: "Waiting for your answers...",
    report_start: "Creating the Word report...",
    report_add_content: "Writing report content...",
    report_add_chart: "Adding a chart...",
    report_add_image: "Adding an image...",
    replace_text: "Filling in the template...",
    excel_start: "Creating the Excel workbook...",
    excel_add_sheet: "Building an Excel sheet...",
    excel_set_cells: "Adding formulas...",
    excel_add_chart: "Adding an Excel chart...",
    excel_read: "Reading the workbook...",
    excel_format: "Formatting cells...",
    excel_conditional_format: "Applying conditional formatting...",
    excel_layout: "Adjusting the sheet layout...",
    excel_design_report: "Designing and verifying the workbook...",
    ppt_start: "Creating the presentation...",
    ppt_add_slides: "Building slides...",
    ppt_add_table_slide: "Adding a table slide...",
    ppt_add_chart_slide: "Adding a chart slide...",
    export_pdf: "Exporting to PDF...",
    rate_limit: "Waiting for Groq rate limit, retrying...",
  };
  return map[name] || "Agent is working...";
}

// ---------- Workspace files ----------
function fmtSize(n) {
  if (n == null) return "";
  if (n < 1024) return n + " B";
  if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
  return (n / 1048576).toFixed(1) + " MB";
}

function fileLine(text, cls) {
  const li = document.createElement("li");
  li.className = cls || "";
  li.textContent = text;
  filesEl.appendChild(li);
}

async function loadFiles() {
  try {
    const r = await fetch("/files");
    if (!r.ok) throw new Error("no route");
    const items = await r.json();
    filesEl.textContent = "";
    if (!items.length) { fileLine("Workspace is empty", "empty"); return; }
    items.forEach((f) => {
      const depth = f.path.split("/").length - 1;
      const li = document.createElement("li");
      li.style.paddingLeft = 8 + depth * 14 + "px";
      li.appendChild(icon(f.dir ? "folder" : "file", f.dir ? "file-ico dir" : "file-ico"));

      let name;
      if (f.dir) {
        name = document.createElement("span");
      } else {
        name = document.createElement("a");
        name.href = "/download?path=" + encodeURIComponent(f.path);
        name.setAttribute("download", "");
      }
      name.className = "f-name";
      name.textContent = f.name;
      name.title = f.path;
      li.appendChild(name);

      const size = document.createElement("span");
      size.className = "f-size";
      size.textContent = fmtSize(f.size);
      li.appendChild(size);
      filesEl.appendChild(li);
    });
  } catch (e) {
    filesEl.textContent = "";
    fileLine("File list unavailable", "empty");
  }
}
$("refresh").onclick = loadFiles;

// ---------- Upload a template / sample (button, paperclip, drag and drop) ----------
async function uploadFiles(list) {
  for (const f of list) {
    const fd = new FormData();
    fd.append("file", f);
    try {
      const r = await fetch("/upload", { method: "POST", body: fd });
      const j = await r.json();
      if (!r.ok) throw new Error(j.detail || "Upload failed");
      addNote("Uploaded " + j.saved + ". Ask me to use it for your document.");
    } catch (e) {
      addNote("Upload failed (" + f.name + "): " + e.message);
    }
  }
  loadFiles();
}

document.querySelectorAll(".js-upload").forEach((b) => {
  b.onclick = () => { fileInput.click(); toggleSidebar(false); };
});
fileInput.onchange = () => {
  if (fileInput.files.length) uploadFiles([...fileInput.files]);
  fileInput.value = "";
};

let dragDepth = 0;
function hasFiles(e) {
  return e.dataTransfer && [...e.dataTransfer.types].includes("Files");
}
window.addEventListener("dragenter", (e) => {
  if (!hasFiles(e)) return;
  dragDepth++;
  dropEl.classList.remove("hidden");
});
window.addEventListener("dragleave", () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) dropEl.classList.add("hidden");
});
window.addEventListener("dragover", (e) => { if (hasFiles(e)) e.preventDefault(); });
window.addEventListener("drop", (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  dropEl.classList.add("hidden");
  if (e.dataTransfer.files.length) uploadFiles([...e.dataTransfer.files]);
});

// ---------- WebSocket ----------
function connect() {
  ws = new WebSocket("ws://" + location.host + "/ws");

  ws.onopen = () => {
    dot.className = "dot on";
    statusText.textContent = "Connected";
  };

  ws.onclose = () => {
    dot.className = "dot off";
    statusText.textContent = "Reconnecting...";
    setBusy(false);
    setTimeout(connect, 2000);
  };

  ws.onmessage = (e) => {
    const m = JSON.parse(e.data);
    if (m.type === "tool") {
      // Tool steps are hidden from the chat; only the status line updates.
      if (workingText) workingText.textContent = statusFor(m.text);
    } else if (m.type === "agent") {
      addMessage(m.text, m.text.startsWith("Error") ? "error" : "agent");
    } else if (m.type === "confirm") {
      addConfirm(m.id, m.text);
    } else if (m.type === "model") {
      addNote(m.note);
      loadModels();
    } else if (m.type === "ask") {
      addAsk(m.id, m.text);
    } else if (m.type === "done") {
      setBusy(false);
      loadFiles();
      input.focus();
    }
  };
}

// ---------- Sending ----------
function autoGrow() {
  input.style.height = "auto";
  input.style.height = Math.min(input.scrollHeight, 160) + "px";
}

function send() {
  const text = input.value.trim();
  if (!text || !ws || ws.readyState !== 1 || sendBtn.disabled) return;
  addMessage(text, "user");
  ws.send(JSON.stringify({ type: "message", text: text }));
  input.value = "";
  autoGrow();
  setBusy(true);
}

sendBtn.onclick = send;
input.addEventListener("input", autoGrow);
input.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    send();
  }
});

document.querySelectorAll("[data-prompt]").forEach((el) => {
  el.onclick = () => {
    input.value = el.dataset.prompt;
    autoGrow();
    input.focus();
    toggleSidebar(false);
  };
});

$("reset").onclick = () => {
  if (ws && ws.readyState === 1) ws.send(JSON.stringify({ type: "reset" }));
  chat.querySelectorAll(".msg, .confirm, .ask, .note").forEach((n) => n.remove());
  welcome.classList.remove("hidden");
  setBusy(false);
  toggleSidebar(false);
};


// ---------- Model status list (priority order, live) ----------
function renderModels(d) {
  meta.textContent = "";
  const head = document.createElement("div");
  head.className = "mhead";
  head.textContent = "Models";
  meta.appendChild(head);
  d.models.forEach((m) => {
    const row = document.createElement("div");
    row.className = "mrow " + m.state + (m.name === d.current ? " cur" : "");
    row.title = m.name;
    const dotEl = document.createElement("span");
    dotEl.className = "mdot";
    const name = document.createElement("span");
    name.className = "mname";
    name.textContent = m.name.split("/").pop();
    const note = document.createElement("span");
    note.className = "mnote";
    note.textContent = m.state === "limited" ? "limit " + m.wait + "s"
      : m.state === "unavailable" ? "unavailable"
      : m.name === d.current ? "active" : "ready";
    row.append(dotEl, name, note);
    meta.appendChild(row);
  });
}

async function loadModels() {
  try {
    const r = await fetch("/models");
    if (r.ok) renderModels(await r.json());
  } catch (e) {}
}
setInterval(loadModels, 5000);
loadModels();

loadFiles();
connect();