let API = null;
const state = { path: null, dealId: null, dealName: null, queue: [], qIndex: 0, lastAnalyze: null };

const $ = (id) => document.getElementById(id);
const overlay = (on, txt) => { $("overlay").classList.toggle("hidden", !on); if (txt) $("overlay-text").textContent = txt; };

const SCREENS = ["settings", "file", "deal", "property", "questions", "review", "done"];
const PHASE = { file: "setup", deal: "setup", property: "setup", questions: "questions", review: "review", done: "review" };

function showScreen(name) {
  SCREENS.forEach(s => $("screen-" + s).classList.toggle("hidden", s !== name));
  const order = { setup: 0, questions: 1, review: 2 };
  const cur = order[PHASE[name]];
  document.querySelectorAll(".steps .step").forEach(el => {
    const i = order[el.dataset.step];
    el.classList.toggle("active", i === cur);
    el.classList.toggle("done", cur !== undefined && i < cur);
  });
  $("steps").style.display = name === "settings" ? "none" : "flex";
}

window.addEventListener("pywebviewready", async () => {
  API = window.pywebview.api;
  const s = await API.check_setup();
  if (s.configured) { setConn(true); showScreen("file"); }
  else { showScreen("settings"); }
});

function setConn(ok) {
  $("conn").textContent = ok ? "Connected" : "Not connected";
  $("conn").className = "conn " + (ok ? "ok" : "bad");
}
function msg(id, text, err) { const el = $(id); el.textContent = text; el.className = "msg " + (text ? (err ? "err" : "ok") : ""); }

// ---------- Settings ----------
$("s-save").onclick = async () => {
  const u = $("s-user").value.trim(), p = $("s-pass").value, t = $("s-token").value.trim(), d = $("s-domain").value.trim();
  if (!u || !p) { msg("s-msg", "Username and password are required.", true); return; }
  overlay(true, "Connecting to Salesforce…");
  const r = await API.save_settings(u, p, t, d);
  overlay(false);
  if (r.ok) { setConn(true); showScreen("file"); }
  else { msg("s-msg", r.error || "Could not connect.", true); }
};

// ---------- Step 1a: file ----------
$("pick").onclick = async () => {
  overlay(true, "Reading file…");
  const r = await API.pick_excel();
  overlay(false);
  if (r.cancelled) return;
  if (!r.ok) { msg("file-msg", r.error, true); return; }
  state.path = r.path;
  $("sheet-row").classList.remove("hidden");
  await reloadSheet();
};
async function reloadSheet() {
  if (!state.path) return;
  overlay(true, "Reading sheet…");
  const r = await API.load_excel_sheet(state.path, $("sheet").value.trim() || null);
  overlay(false);
  if (!r.ok) { msg("file-msg", r.error, true); $("to-deal").disabled = true; return; }
  msg("file-msg", "", false);
  $("file-label").textContent = r.file;
  $("file-sub").textContent = `${r.rows} rows · ${r.columns.join(", ")}`;
  $("to-deal").disabled = false;
}
$("sheet").onchange = reloadSheet;
$("to-deal").onclick = () => showScreen("deal");

// ---------- Step 1b: deal ----------
let dealTimer = null;
$("deal-q").oninput = () => { clearTimeout(dealTimer); dealTimer = setTimeout(searchDeals, 250); };
async function searchDeals() {
  const q = $("deal-q").value.trim();
  const box = $("deal-results");
  if (q.length < 2) { box.innerHTML = ""; return; }
  const deals = await API.search_deals(q);
  box.innerHTML = "";
  deals.forEach(d => {
    const el = document.createElement("div");
    el.className = "item"; el.textContent = d.name;
    el.onclick = () => chooseDeal(d);
    box.appendChild(el);
  });
}
function chooseDeal(d) {
  state.dealId = d.id; state.dealName = d.name;
  $("deal-results").innerHTML = ""; $("deal-q").value = "";
  const c = $("deal-chosen");
  c.classList.remove("hidden");
  c.innerHTML = `<span>${d.name}</span><span class="badge">✓ matched</span>`;
  $("to-property").disabled = false;
}
$("deal-back").onclick = () => showScreen("file");
$("to-property").onclick = () => showScreen("property");

// ---------- Step 1c: property ----------
$("property").oninput = () => {
  const p = $("property").value.trim();
  $("name-preview").textContent = p ? `→ e.g. “REIT - ${p}”` : "";
  $("analyze").disabled = !p;
};
$("prop-back").onclick = () => showScreen("deal");
$("analyze").onclick = runAnalyze;

async function runAnalyze() {
  overlay(true, "Matching accounts & contacts…");
  const r = await API.analyze(state.dealId, state.dealName, $("property").value.trim());
  overlay(false);
  state.lastAnalyze = r;
  if (r.questions && r.questions.length) { state.queue = r.questions; state.qIndex = 0; showScreen("questions"); renderQuestion(); }
  else { showReview(r); }
}

// ---------- Step 2: questions ----------
function renderQuestion() {
  const q = state.queue[state.qIndex];
  $("q-progress").textContent = `Question ${state.qIndex + 1} of ${state.queue.length}`;
  $("q-title").textContent = q.kind === "account"
    ? `Which account is “${q.typed}”?`
    : `Which person is “${q.typed}”?` + (q.row ? ` (row ${q.row})` : "");
  let sel = null;
  const card = $("q-card"); card.innerHTML = "";
  q.candidates.forEach(c => {
    const el = document.createElement("div");
    el.className = "qopt";
    el.innerHTML = `<span class="radio"></span><span>${c.name}</span>` + (c.score != null ? `<span class="score">${c.score}%</span>` : "");
    el.onclick = () => { sel = c.id; [...card.children].forEach(x => x.classList.remove("sel")); el.classList.add("sel"); $("q-confirm").disabled = false; };
    card.appendChild(el);
  });
  const foot = document.createElement("div");
  foot.className = "qfoot";
  foot.innerHTML = `<button id="q-skip" class="ghost">Skip (leave blank)</button>
    <label class="remember"><input id="q-remember" type="checkbox" checked/> Remember this choice</label>
    <button id="q-confirm" class="primary" disabled>Confirm</button>`;
  card.appendChild(foot);
  $("q-skip").onclick = () => answer(q, null);
  $("q-confirm").onclick = () => { if (sel) answer(q, sel); };
}
async function answer(q, choiceId) {
  overlay(true, "Saving…");
  await API.answer(q.id, choiceId, $("q-remember") ? $("q-remember").checked : true);
  overlay(false);
  if (state.qIndex + 1 < state.queue.length) { state.qIndex++; renderQuestion(); }
  else { runAnalyze(); }
}

// ---------- Step 3: review ----------
function showReview(r) {
  $("deal-line").textContent = `Deal: ${r.deal}`;
  const m = r.summary;
  $("metrics").innerHTML = [
    ["Ready", m.ready], ["Auto-fixed", m.auto_fixed], ["You chose", m.you_chose], ["Blank contact", m.blank_contacts]
  ].map(([k, v]) => `<div class="metric"><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");
  const rows = r.preview.map(p =>
    `<tr><td>${p.name}</td><td>${p.interest || ""}</td>
     <td><span class="tag ${p.account ? "yes" : "no"}">${p.account ? "set" : "blank"}</span></td>
     <td><span class="tag ${p.contact ? "yes" : "no"}">${p.contact ? "set" : "blank"}</span></td></tr>`).join("");
  $("preview").innerHTML = `<table><tr><th>Name</th><th>Interest</th><th>Account</th><th>Contact</th></tr>${rows}</table>
     <div class="muted small" style="margin-top:6px">showing first ${r.preview.length} of ${m.ready}</div>`;
  $("upload").textContent = `Upload ${m.ready} records to Salesforce`;
  showScreen("review");
}
$("back-setup").onclick = resetToStart;

$("upload").onclick = async () => {
  overlay(true, "Uploading to Salesforce… this can take a minute.");
  const r = await API.upload();
  overlay(false);
  const box = $("done-box");
  const check = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 12l5 5L20 6"/></svg>`;
  const alert = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 8v5M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>`;
  if (r.ok) {
    box.className = "donebox";
    box.innerHTML = `<div class="big">${check}</div><h2>${r.created} records created</h2>
      <p class="muted">on ${r.deal}${r.failed ? ` · ${r.failed} failed` : ""}</p>` +
      (r.failed ? `<div class="muted small" style="margin-top:8px">${r.errors.map(e => `row ${e.row}: ${e.error}`).join("<br>")}</div>` : "");
  } else {
    box.className = "donebox fail";
    box.innerHTML = `<div class="big">${alert}</div><h2>Upload failed</h2><p class="muted">${r.error}</p>`;
  }
  showScreen("done");
};
$("another").onclick = resetToStart;

function resetToStart() {
  state.dealId = state.dealName = null;
  $("deal-chosen").classList.add("hidden"); $("to-property").disabled = true;
  $("property").value = ""; $("name-preview").textContent = ""; $("analyze").disabled = true;
  showScreen("file");
}
