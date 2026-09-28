"use strict";

const I18N = {
  my: {
    brand: "PPT Slide Maker", brandSub: "AI ဆလိုက်ဒ်ဖန်တီးစက်",
    settings: "⚙ ဆက်တင်", newDeck: "ဆလိုက်အသစ်ဖန်တီးရန်",
    topic: "အကြောင်းအရာ *", topicPh: "ဥပမာ — မြန်မာ့ရိုးရာအစားအစာ",
    detail: "အသေးစိတ် (ရွေးချယ်ရန်)", detailPh: "ပရိသတ်၊ ရည်ရွယ်ချက်၊ ထည့်ချင်တဲ့အချက်များ…",
    aiDetail: "AI ဖြင့် ဖန်တီးရန်", aiDetailing: "AI ရေးနေတယ်…", detailFilled: "AI က အသေးစိတ်ရေးပေးပြီးပါပြီ ✨",
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
    loadingTitle: "ဆလိုက်ထုတ်လုပ်နေတယ်",
    loadS1: "AI က အကြောင်းအရာရေးနေတယ်…",
    loadS2: "ဒီဇိုင်းနဲ့ အရောင်ရွေးနေတယ်…",
    loadS3: "layout ဆွဲနေတယ်…",
    loadS4: "PPTX ဖိုင်တည်ဆောက်နေတယ်…",
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
    aiDetail: "Generate with AI", aiDetailing: "AI is writing…", detailFilled: "AI wrote the details ✨",
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
    loadingTitle: "Generating your deck",
    loadS1: "AI is writing the content…",
    loadS2: "Choosing the design…",
    loadS3: "Laying out the slides…",
    loadS4: "Building the PPTX file…",
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
let selectedTheme = "midnight-glass";
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
// Keyword highlight: ==phrase== -> accent-colored bold span (HTML-escaped first).
function hl(t, accent) {
  return esc(t).replace(/==(.+?)==/g, `<span style="color:${accent};font-weight:700">$1</span>`);
}

function bText(b) {
  return b && typeof b === "object" ? String(b.text || "") : String(b || "");
}
function bIcon(b) {
  return b && typeof b === "object" ? String(b.icon || "") : "";
}
function bulletLis(bullets, c) {
  const ac = c ? `#${c.accent}` : "inherit";
  return (bullets || [])
    .map((b) => {
      const t = bText(b), ic = bIcon(b);
      return t ? `<li>${ic ? `<span style="margin-right:6px">${esc(ic)}</span>` : ""}${hl(t, ac)}</li>` : "";
    })
    .join("");
}
function takeawayHtml(tw, c) {
  if (!tw || !tw.trim()) return "";
  return `<div style="margin-top:14px;padding:10px 16px;border:1.5px solid #${c.accent};border-radius:999px;background:rgba(255,255,255,0.06);color:#${c.title};font-style:italic;font-size:clamp(11px,1.4vw,16px);line-height:1.4">✦&nbsp;&nbsp;${hl(tw.trim(), `#${c.accent}`)}</div>`;
}

function layoutOf(item) {
  if (item.layout === "stats" && (item.stats || []).length >= 2) return "stats";
  if ((item.layout === "stat" || item.layout === "stats") && ((item.stats || []).length === 1 || (item.stat && item.stat.value))) return "stat";
  if (item.layout === "timeline" && (item.points || []).length >= 2) return "timeline";
  if (item.layout === "split" && (item.bullets || []).length >= 2 && item.icon) return "split";
  if (item.layout === "hero" && (item.hero || item.heading)) return "hero";
  if (item.layout === "cards" && ((item.points || []).length >= 2 || (item.bullets || []).length >= 2)) return "cards";
  if (item.layout === "quote" && (item.quote || (item.bullets || []).length)) return "quote";
  if (item.layout === "two-col" && (item.bullets || []).length > 2) return "two-col";
  return "bullets";
}

function slideHtml(item, i, total, c, ui) {
  const layout = layoutOf(item);
  const head =
    `<div class="snum">${i + 1} / ${total}</div>` +
    `<div style="width:44px;height:4px;background:#${c.accent};border-radius:2px;margin-bottom:12px"></div>` +
    `<h3 style="color:#${c.title}">${hl(item.heading, `#${c.accent}`)}</h3>`;
  const iconBg = item.icon
    ? `<div style="position:absolute;right:4%;top:18%;font-size:110px;opacity:0.16;pointer-events:none">${esc(item.icon)}</div>`
    : "";

  if (layout === "stats") {
    const cards = (item.stats || []).slice(0, 4).map((st) => `
      <div class="pv-glass" style="flex:1;padding:14px 16px;min-width:0">
        <div style="width:22px;height:4px;background:#${c.accent};border-radius:2px;margin-bottom:10px"></div>
        <div style="font-size:clamp(24px,3vw,40px);font-weight:800;color:#${c.accent};line-height:1">${esc(st.value)}</div>
        <div style="margin-top:8px;font-size:clamp(11px,1.3vw,16px);color:#${c.text};line-height:1.6">${esc(st.label)}</div>
      </div>`).join("");
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${head}
      <div style="display:flex;gap:14px;margin-top:3%;flex-wrap:wrap">${cards}</div></div>`;
  }
  if (layout === "cards") {
    const pts = (item.points || []).length >= 2 ? item.points.slice(0, 3)
      : (item.bullets || []).slice(0, 3).map((b) => {
          const t = bText(b);
          const m = t.split(/[:—–-]\s(.+)/);
          return { icon: bIcon(b) || item.icon, title: m.length > 2 ? m[0].trim() : "", text: m.length > 2 ? m[1].trim() : t };
        });
    const cards = pts.map((p) => `
      <div class="pv-glass" style="flex:1;padding:16px;min-width:0;text-align:center">
        <div style="width:22px;height:4px;background:#${c.accent};border-radius:2px;margin:0 auto 12px"></div>
        ${p.icon ? `<div style="width:52px;height:52px;border-radius:50%;background:#${c.accent}29;display:flex;align-items:center;justify-content:center;font-size:30px;margin:0 auto 10px">${esc(p.icon)}</div>` : ""}
        ${p.title ? `<div style="font-weight:700;color:#${c.title};font-size:clamp(13px,1.8vw,24px);letter-spacing:-0.02em;line-height:1.1;margin-bottom:6px">${hl(p.title, `#${c.accent}`)}</div>` : ""}
        <div style="color:#${c.text};font-size:clamp(11px,1.3vw,16px);line-height:1.6">${hl(p.text, `#${c.accent}`)}</div>
      </div>`).join("");
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${head}
      <div style="display:flex;gap:14px;margin-top:3%">${cards}</div></div>`;
  }
  if (layout === "stat") {
    const st = (item.stats && item.stats[0]) || item.stat || {};
    const bullets = bulletLis(item.bullets, c);
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${iconBg}${head}
      <div style="display:flex;gap:24px;margin-top:4%;align-items:flex-start">
        <div class="pv-glass" style="padding:18px 22px;min-width:34%">
          <div style="font-size:clamp(34px,4.5vw,58px);font-weight:800;color:#${c.accent};line-height:1">${esc(st.value)}</div>
          <div style="margin-top:8px;font-size:clamp(11px,1.3vw,16px);color:#${c.text};line-height:1.6">${esc(st.label)}</div>
        </div>
        <ul style="color:#${c.text};margin:0;padding-left:20px;font-size:clamp(11px,1.4vw,16px);line-height:1.6">${bullets}</ul>
      </div></div>`;
  }
  if (layout === "quote") {
    const quote = item.quote || (item.bullets || []).map(bText).join(" ");
    return `<div class="slide-card" style="background:#${c.bgDeep};color:#${c.text};justify-content:center;align-items:center;text-align:center">
      <div class="snum">${i + 1} / ${total}</div>
      <div style="font-size:90px;color:#${c.accent};line-height:0.6;margin-bottom:16px">&ldquo;</div>
      <div style="font-size:clamp(15px,2vw,24px);font-style:italic;color:#${c.title};max-width:80%;line-height:1.3">${hl(quote, `#${c.accent}`)}</div>
      ${item.quoteBy ? `<div style="margin-top:14px;color:#9CA3AF;font-style:italic;font-size:12px">— ${esc(item.quoteBy)}</div>` : ""}
    </div>`;
  }
  if (layout === "two-col") {
    const mid = Math.ceil((item.bullets || []).length / 2);
    const l = bulletLis((item.bullets || []).slice(0, mid), c);
    const r = bulletLis((item.bullets || []).slice(mid), c);
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${head}
      <div style="display:flex;gap:28px;margin-top:3%">
        <ul style="flex:1;color:#${c.text};margin:0;padding-left:20px;font-size:clamp(11px,1.4vw,16px);line-height:1.6">${l}</ul>
        <div style="width:2px;background:#${c.bgDeep};border-radius:1px"></div>
        <ul style="flex:1;color:#${c.text};margin:0;padding-left:20px;font-size:clamp(11px,1.4vw,16px);line-height:1.6">${r}</ul>
      </div>${takeawayHtml(item.takeaway, c)}</div>`;
  }
  if (layout === "timeline") {
    const steps = (item.points || []).slice(0, 4);
    const nodes = steps.map((st, k) => `
      <div style="flex:1;min-width:0;text-align:center;position:relative">
        <div style="width:44px;height:44px;border-radius:50%;background:#${c.accent};color:#fff;font-weight:800;display:flex;align-items:center;justify-content:center;margin:0 auto 12px;font-size:16px;position:relative;z-index:1">${String(k + 1).padStart(2, "0")}</div>
        ${st.title ? `<div style="font-weight:700;color:#${c.title};font-size:clamp(12px,1.6vw,19px);letter-spacing:-0.02em;line-height:1.15;margin-bottom:6px">${hl(st.title, `#${c.accent}`)}</div>` : ""}
        <div style="color:#${c.text};font-size:clamp(10px,1.2vw,15px);line-height:1.5">${hl(st.text, `#${c.accent}`)}</div>
      </div>`).join("");
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${head}
      <div style="position:relative;margin-top:5%">
        <div style="position:absolute;top:22px;left:12%;right:12%;height:3px;background:#${c.accent}66;border-radius:2px"></div>
        <div style="display:flex;gap:18px;position:relative">${nodes}</div>
      </div></div>`;
  }
  if (layout === "split") {
    const bullets = bulletLis(item.bullets, c);
    return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${head}
      <div style="display:flex;gap:24px;margin-top:3%;align-items:stretch">
        <div class="pv-glass" style="flex:0 0 34%;display:flex;align-items:center;justify-content:center;min-height:220px;position:relative;overflow:hidden">
          <div style="position:absolute;width:170px;height:170px;border-radius:50%;border:3px solid #${c.accent}55"></div>
          <div style="position:absolute;width:120px;height:120px;border-radius:50%;background:#${c.accent}26"></div>
          <div style="font-size:96px;position:relative;z-index:1">${esc(item.icon || "")}</div>
        </div>
        <ul style="flex:1;color:#${c.text};margin:0;padding-left:20px;font-size:clamp(11px,1.4vw,17px);line-height:1.6;align-self:center">${bullets}</ul>
      </div>${takeawayHtml(item.takeaway, c)}</div>`;
  }
  if (layout === "hero") {
    const hero = item.hero || item.heading;
    return `<div class="slide-card" style="background:#${c.bgDeep};color:#${c.text};justify-content:center;align-items:center;text-align:center">
      <div class="snum">${i + 1} / ${total}</div>
      <div style="position:absolute;width:min(46vw,340px);height:min(46vw,340px);border-radius:50%;border:4px solid #${c.accent}40;pointer-events:none"></div>
      <div style="font-size:clamp(30px,4.6vw,60px);font-weight:800;color:#${c.title};letter-spacing:-0.02em;line-height:1.05;max-width:82%;position:relative">${hl(hero, `#${c.accent}`)}</div>
      ${item.sub ? `<div style="margin-top:18px;font-size:clamp(12px,1.6vw,18px);font-style:italic;color:#${c.text};max-width:70%;line-height:1.5">${hl(item.sub, `#${c.accent}`)}</div>` : ""}
    </div>`;
  }
  const bullets = bulletLis(item.bullets, c);
  return `<div class="slide-card" style="background:#${c.bg};color:#${c.text}">${iconBg}
    <div class="snum">${i + 1} / ${total}</div>
    <div style="width:44px;height:4px;background:#${c.accent};border-radius:2px;margin-bottom:12px"></div>
    <h3 style="color:#${c.title}">${hl(item.heading, `#${c.accent}`)}</h3>
    <ul style="color:#${c.text}">${bullets}</ul>
    ${takeawayHtml(item.takeaway, c)}
  </div>`;
}

function renderPreview(deck, themeKey, lang) {
  window._lastDeck = { deck, theme: themeKey, lang };
  const th = themeByKey(themeKey);
  const c = th.colors;
  const total = deck.slides.length + 2;
  let html =
    `<div class="slide-card title-slide" style="background:#${c.bgDeep};color:#${c.text}">
      ${deck.icon ? `<div style="font-size:64px;margin-bottom:8px">${esc(deck.icon)}</div>` : ""}
      <div class="kicker" style="color:#${c.accent}">${esc(t("kicker"))}</div>
      <h1 style="color:#${c.title}">${esc(deck.title)}</h1>
      ${deck.subtitle ? `<div class="subtitle" style="color:#${c.title}">${esc(deck.subtitle)}</div>` : ""}
    </div>`;
  deck.slides.forEach((s, i) => { html += slideHtml(s, i + 1, total, c, uiLang); });
  html +=
    `<div class="slide-card closing" style="background:#${c.bgDeep};color:#${c.text}">
      <h2 style="color:#${c.title};font-size:clamp(26px,4vw,44px);font-weight:800;letter-spacing:-0.02em;line-height:1.1;margin:0 0 12px">${lang === "en" ? "Thank You" : "ကျေးဇူးတင်ပါတယ်"}</h2>
      <div style="color:#${c.text};font-size:clamp(12px,1.5vw,18px);line-height:1.6">${esc(deck.title)}</div>
      <div style="margin-top:10px;color:#${c.accentSoft};font-style:italic;font-size:clamp(11px,1.4vw,16px)">${esc(t("closingQ"))}</div>
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

let loadTimer = null;

function showLoading() {
  const steps = [t("loadS1"), t("loadS2"), t("loadS3"), t("loadS4")];
  let i = 0;
  $("previewEmpty").classList.add("hidden");
  $("downloadBtn").classList.add("hidden");
  $("previewSlides").innerHTML =
    `<div class="loading-wrap">
       <div class="load-deck" aria-hidden="true">
         <div class="load-slide">
           <div class="ls-kicker"></div><div class="ls-title"></div>
           <div class="ls-line" style="width:88%"></div>
           <div class="ls-line" style="width:64%"></div>
           <div class="ls-line" style="width:76%"></div>
         </div>
         <div class="load-slide">
           <div class="ls-cards"><div></div><div></div><div></div></div>
           <div class="ls-line" style="width:52%"></div>
         </div>
         <div class="load-slide">
           <div class="ls-hero-ring"></div>
           <div class="ls-hero"></div>
           <div class="ls-line" style="width:44%"></div>
         </div>
       </div>
       <div class="loading-title">${esc(t("loadingTitle"))}<span class="loading-dots"></span></div>
       <div class="load-bar"><div class="load-bar-fill" id="loadBarFill"></div></div>
       <ol class="load-steps" id="loadSteps">
         ${steps.map((s) => `<li><span class="ls-dot"></span><span>${esc(s)}</span></li>`).join("")}
       </ol>
     </div>`;
  const paint = () => {
    document.querySelectorAll("#loadSteps li").forEach((li, k) => {
      li.classList.toggle("done", k < i);
      li.classList.toggle("active", k === i);
    });
    const f = $("loadBarFill");
    if (f) f.style.width = (((i + 1) / steps.length) * 100).toFixed(0) + "%";
  };
  paint();
  loadTimer = setInterval(() => {
    if (i < steps.length - 1) { i++; paint(); }
  }, 2800);
}

function hideLoading() {
  if (loadTimer) { clearInterval(loadTimer); loadTimer = null; }
}

async function generateDetails() {
  const topic = $("topic").value.trim();
  if (!topic) { setStatus(t("errTopic"), "error"); $("topic").focus(); return; }
  const btn = $("aiDetailBtn");
  const label = btn.querySelector("[data-i18n]");
  btn.disabled = true;
  label.textContent = t("aiDetailing");
  try {
    const r = await fetch("/api/details/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ topic, lang: $("deckLang").value, model: $("modelSel").value }),
    });
    const d = await r.json();
    if (!r.ok) {
      setStatus(d.error === "no-key" ? t("errKey") : t("errGen") + ": " + (d.error || ""), "error");
      return;
    }
    const ta = $("detail");
    ta.value = d.text;
    ta.classList.remove("flash");
    void ta.offsetWidth; /* restart the flash animation */
    ta.classList.add("flash");
    setStatus(t("detailFilled"), "ok");
  } catch (e) {
    setStatus(t("errGen") + ": " + String(e.message || e).slice(0, 200), "error");
  } finally {
    btn.disabled = false;
    label.textContent = t("aiDetail");
  }
}

async function generate() {
  const topic = $("topic").value.trim();
  if (!topic) { setStatus(t("errTopic"), "error"); return; }
  const btn = $("generateBtn");
  btn.disabled = true;
  setStatus(t("generating"), "");
  showLoading();
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
      $("previewSlides").innerHTML = "";
      $("previewEmpty").classList.remove("hidden");
      setStatus((d.error === "no-key" ? t("errKey") : t("errGen") + ": " + (d.error || "")) + (d.detail ? " — " + d.detail : ""), "error");
      return;
    }
    currentDeckId = d.id;
    renderPreview(d.deck, d.theme, d.lang);
    setStatus(t("genDone"), "ok");
    renderGallery();
  } catch (e) {
    $("previewSlides").innerHTML = "";
    $("previewEmpty").classList.remove("hidden");
    setStatus(t("errGen") + ": " + String(e.message || e).slice(0, 200), "error");
  } finally {
    hideLoading();
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
$("aiDetailBtn").onclick = generateDetails;
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
