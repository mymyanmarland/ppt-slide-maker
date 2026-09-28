"use strict";

const I18N = {
  my: {
    brand: "PPT Slide Maker", brandSub: "AI ဆလိုက်ဒ်ဖန်တီးစက်",
    settings: "⚙ ဆက်တင်", newDeck: "ဆလိုက်အသစ်ဖန်တီးရန်",
    topic: "အကြောင်းအရာ *", topicPh: "ဥပမာ — မြန်မာ့ရိုးရာအစားအစာ",
    detail: "အသေးစိတ် (ရွေးချယ်ရန်)", detailPh: "ပရိသတ်၊ ရည်ရွယ်ချက်၊ ထည့်ချင်တဲ့အချက်များ…",
    slideCount: "ဆလိုက်အရေအတွက်", lang: "ဘာသာစကား", theme: "ဒီဇိုင်း", model: "AI မော်ဒယ်",
    generate: "✨ ဆလိုက်ထုတ်ရန်", preview: "အစမ်းကြည့်ရှုခြင်း",
    downloadPptx: "PPTX ဒေါင်းလုဒ်",
    previewEmpty: "အကြောင်းအရာရိုက်ပြီး \"ဆလိုက်ထုတ်ရန်\" ကိုနှိပ်ပါ — AI က အကြောင်းအရာရေးပေးပြီး တကယ့် .pptx ဖိုင်အဖြစ်ရမယ်။",
    gallery: "သိမ်းထားသော ဆလိုက်များ",
    noKey: "API key မရှိသေးပါ — ဆက်တင်မှာ သင့် gateway key ကိုထည့်ပါ။",
    openSettings: "ဆက်တင်ဖွင့်",
    ephemeral: "APP_SECRET မသတ်မှတ်ထားလို့ key က restart ပြန်စရင် ပျောက်မယ်။",
    settingsTitle: "Gateway ဆက်တင်", baseUrl: "Gateway URL", apiKey: "API Key",
    apiKeyPh: "ထည့်ပြီးသားရှိရင် အလွတ်ထားနိုင်", defaultModel: "ပုံသေမော်ဒယ်",
    testConn: "စမ်းသပ်ရန်", close: "ပိတ်ရန်", save: "သိမ်းရန်",
    generating: "AI က ရေးနေတယ်… ခဏစောင့်ပါ ⏳",
    genDone: "ပြီးပါပြီ ✅ — အစမ်းကြည့်ပြီး PPTX ဒေါင်းလုဒ်လုပ်နိုင်ပါပြီ။",
    errTopic: "အကြောင်းအရာ အရင်ရိုက်ပါ။",
    errKey: "API key မရှိသေးပါ။ ဆက်တင်မှာ ထည့်ပါ။",
    errGen: "ထုတ်လုပ်ရာမှာ အမှားဖြစ်တယ်",
    saved: "သိမ်းပြီးပါပြီ ✅", testOk: "ချိတ်ဆက်မှုအောင်မြင်တယ် ✅", testFail: "မအောင်မြင်ပါ",
    download: "ဒေါင်းလုဒ်", del: "ဖျက်", confirmDel: "ဒီဆလိုက်ကို ဖျက်မှာလား?",
    emptyGallery: "မရှိသေးပါ — အပေါ်ကနေ အသစ်ဖန်တီးပါ။",
    kicker: "AI ဆလိုက်ဒ်", closingQ: "မေးခွန်းများ နှင့် ဆွေးနွေးခန်း",
  },
  en: {
    brand: "PPT Slide Maker", brandSub: "AI slide generator",
    settings: "⚙ Settings", newDeck: "Create new slides",
    topic: "Topic *", topicPh: "e.g. Traditional Myanmar food",
    detail: "Details (optional)", detailPh: "Audience, purpose, points to include…",
    slideCount: "Slide count", lang: "Language", theme: "Theme", model: "AI model",
    generate: "✨ Generate slides", preview: "Preview",
    downloadPptx: "Download PPTX",
    previewEmpty: "Type a topic and hit Generate — the AI writes the content and you get a real .pptx file.",
    gallery: "Saved decks",
    noKey: "No API key yet — add your gateway key in Settings.",
    openSettings: "Open settings",
    ephemeral: "APP_SECRET is not set, so the key will be lost on restart.",
    settingsTitle: "Gateway settings", baseUrl: "Gateway URL", apiKey: "API Key",
    apiKeyPh: "Leave blank to keep the existing one", defaultModel: "Default model",
    testConn: "Test connection", close: "Close", save: "Save",
    generating: "AI is writing… please wait ⏳",
    genDone: "Done ✅ — preview it and download the PPTX.",
    errTopic: "Please enter a topic first.",
    errKey: "No API key. Add it in Settings.",
    errGen: "Generation failed",
    saved: "Saved ✅", testOk: "Connection OK ✅", testFail: "Failed",
    download: "Download", del: "Delete", confirmDel: "Delete this deck?",
    emptyGallery: "Nothing yet — create one above.",
    kicker: "AI SLIDE DECK", closingQ: "Questions & Discussion",
  },
};

let uiLang = localStorage.getItem("pptm-ui") || "my";
let THEMES = [];
let selectedTheme = "navy-gold";
let currentDeckId = null;

const $ = (id) => document.getElementById(id);
const t = (k) => (I18N[uiLang] && I18N[uiLang][k]) || I18N.en[k] || k;

function applyI18n() {
  document.documentElement.lang = uiLang === "my" ? "my" : "en";
  document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll("[data-i18n-ph]").forEach((el) => { el.placeholder = t(el.dataset.i18nPh); });
  $("langMy").classList.toggle("active", uiLang === "my");
  $("langEn").classList.toggle("active", uiLang === "en");
  renderThemes();
  renderGallery();
  if (window._lastDeck) renderPreview(window._lastDeck.deck, window._lastDeck.theme, window._lastDeck.lang);
}

function themeByKey(k) { return THEMES.find((x) => x.key === k) || THEMES[0]; }

function renderThemes() {
  const g = $("themeGrid");
  g.innerHTML = "";
  THEMES.forEach((th) => {
    const c = th.colors;
    const card = document.createElement("button");
    card.type = "button";
    card.className = "theme-card" + (th.key === selectedTheme ? " selected" : "");
    card.innerHTML =
      `<div class="theme-swatch" style="background:#${c.bg}"><div class="dot" style="background:#${c.accent}"></div></div>` +
      `<div class="theme-name">${uiLang === "my" ? th.nameMy : th.name}</div>`;
    card.onclick = () => { selectedTheme = th.key; renderThemes(); };
    g.appendChild(card);
  });
}

function esc(s) {
  return String(s || "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

function slideHtml(item, i, total, c, ui) {
  const bullets = (item.bullets || []).map((b) => `<li>${esc(b)}</li>`).join("");
  return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">
    <div class="snum" style="color:#${c.footer}">${i + 1} / ${total}</div>
    <div style="width:44px;height:4px;background:#${c.accent};border-radius:2px;margin-bottom:12px"></div>
    <h3 style="color:#${c.title}">${esc(item.heading)}</h3>
    <ul style="color:#${c.text}">${bullets}</ul>
  </div>`;
}

function renderPreview(deck, themeKey, lang) {
  window._lastDeck = { deck, theme: themeKey, lang };
  const th = themeByKey(themeKey);
  const c = th.colors;
  const total = deck.slides.length + 2;
  let html =
    `<div class="slide-card title-slide" style="background:#${c.bgDeep};color:#${c.text}">
      <div class="kicker" style="color:#${c.accent}">${esc(t("kicker"))}</div>
      <h1 style="color:#${c.title}">${esc(deck.title)}</h1>
      ${deck.subtitle ? `<div class="subtitle" style="color:#${c.muted}">${esc(deck.subtitle)}</div>` : ""}
    </div>`;
  deck.slides.forEach((s, i) => { html += slideHtml(s, i + 1, total, c, uiLang); });
  html +=
    `<div class="slide-card closing" style="background:#${c.bgDeep};color:#${c.text}">
      <h2 style="color:#${c.title}">${lang === "en" ? "Thank You" : "ကျေးဇူးတင်ပါတယ်"}</h2>
      <div style="color:#${c.muted}">${esc(deck.title)}</div>
      <div style="margin-top:10px;color:#${c.accentSoft};font-style:italic">${esc(t("closingQ"))}</div>
    </div>`;
  $("previewSlides").innerHTML = html;
  $("previewEmpty").classList.add("hidden");
  $("downloadBtn").classList.remove("hidden");
}

async function loadStatus() {
  const r = await fetch("/api/status");
  const s = await r.json();
  THEMES = s.themes || [];
  if (!THEMES.find((x) => x.key === selectedTheme) && THEMES.length) selectedTheme = THEMES[0].key;
  renderThemes();
  $("keyBanner").classList.toggle("hidden", s.gatewayConfigured);
  $("ephemeralBanner").classList.toggle("hidden", !s.ephemeralKey);
  const ms = $("modelSel");
  ms.innerHTML = "";
  try {
    const mr = await fetch("/api/models");
    if (mr.ok) {
      const { models } = await mr.json();
      (models || []).forEach((m) => {
        const o = document.createElement("option");
        o.value = m; o.textContent = m;
        if (m === s.model) o.selected = true;
        ms.appendChild(o);
      });
    }
  } catch { /* offline until key is set */ }
  if (!ms.options.length) {
    const o = document.createElement("option");
    o.value = s.model; o.textContent = s.model;
    ms.appendChild(o);
  }
}

async function generate() {
  const topic = $("topic").value.trim();
  if (!topic) { setStatus(t("errTopic"), "error"); return; }
  const btn = $("generateBtn");
  btn.disabled = true;
  setStatus(t("generating"), "");
  try {
    const r = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        topic,
        detail: $("detail").value.trim(),
        slides: Number($("slideCount").value),
        theme: selectedTheme,
        lang: $("deckLang").value,
        model: $("modelSel").value,
      }),
    });
    const d = await r.json();
    if (!r.ok) {
      setStatus((d.error === "no-key" ? t("errKey") : t("errGen") + ": " + (d.error || "")) + (d.detail ? " — " + d.detail : ""), "error");
      return;
    }
    currentDeckId = d.id;
    renderPreview(d.deck, d.theme, d.lang);
    setStatus(t("genDone"), "ok");
    renderGallery();
  } catch (e) {
    setStatus(t("errGen") + ": " + String(e.message || e).slice(0, 200), "error");
  } finally {
    btn.disabled = false;
  }
}

function setStatus(msg, cls) {
  const el = $("genStatus");
  el.textContent = msg;
  el.className = "status " + (cls || "");
}

async function renderGallery() {
  const g = $("gallery");
  try {
    const r = await fetch("/api/decks");
    const { items } = await r.json();
    if (!items.length) { g.innerHTML = `<div class="empty" style="padding:20px"><p>${esc(t("emptyGallery"))}</p></div>`; return; }
    g.innerHTML = "";
    items.forEach((d) => {
      const th = themeByKey(d.theme);
      const card = document.createElement("div");
      card.className = "deck-card";
      const dt = new Date(d.created_at).toLocaleString(uiLang === "my" ? "my-MM" : "en-US");
      card.innerHTML =
        `<div style="height:6px;border-radius:3px;background:#${th.colors.accent}"></div>` +
        `<div class="t">${esc(d.title)}</div>` +
        `<div class="m">${esc(d.topic)} · ${uiLang === "my" ? th.nameMy : th.name} · ${esc(dt)}</div>` +
        `<div class="actions">
           <a class="btn small primary" style="text-decoration:none" href="/api/decks/${d.id}/download">⬇ ${esc(t("download"))}</a>
           <button class="btn small danger" data-del="${d.id}">${esc(t("del"))}</button>
         </div>`;
      g.appendChild(card);
    });
    g.querySelectorAll("[data-del]").forEach((b) => {
      b.onclick = async () => {
        if (!confirm(t("confirmDel"))) return;
        await fetch("/api/decks/" + b.dataset.del, { method: "DELETE" });
        renderGallery();
      };
    });
  } catch { /* ignore */ }
}

// ---- settings modal ----
async function openSettings() {
  const r = await fetch("/api/settings");
  const s = await r.json();
  $("setBaseUrl").value = s.baseUrl || "";
  $("setApiKey").value = "";
  $("setModel").value = s.model || "";
  $("testResult").textContent = "";
  $("settingsModal").classList.remove("hidden");
}
async function saveSettings() {
  const body = {
    baseUrl: $("setBaseUrl").value.trim(),
    model: $("setModel").value.trim(),
  };
  if ($("setApiKey").value) body.apiKey = $("setApiKey").value;
  await fetch("/api/settings", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  $("settingsModal").classList.add("hidden");
  $("testResult").textContent = t("saved");
  loadStatus();
}
async function testConn() {
  const el = $("testResult");
  el.textContent = "…";
  try {
    const r = await fetch("/api/settings/test", { method: "POST" });
    const d = await r.json();
    el.textContent = d.ok ? `${t("testOk")} (${d.count})` : t("testFail") + ": " + (d.error || "");
    el.className = "status " + (d.ok ? "ok" : "error");
  } catch (e) {
    el.textContent = t("testFail");
    el.className = "status error";
  }
}

$("generateBtn").onclick = generate;
$("downloadBtn").onclick = () => {
  if (currentDeckId) window.location.href = "/api/decks/" + currentDeckId + "/download";
};
$("settingsBtn").onclick = openSettings;
$("keyBannerBtn").onclick = openSettings;
$("settingsClose").onclick = () => $("settingsModal").classList.add("hidden");
$("settingsSave").onclick = saveSettings;
$("testConnBtn").onclick = testConn;
$("langMy").onclick = () => { uiLang = "my"; localStorage.setItem("pptm-ui", "my"); applyI18n(); };
$("langEn").onclick = () => { uiLang = "en"; localStorage.setItem("pptm-ui", "en"); applyI18n(); };

applyI18n();
loadStatus();
renderGallery();
