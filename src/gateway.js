"use strict";
// Client for the user's OpenAI-compatible gateway + deck-outline generation.

const DEFAULT_BASE_URL = "https://claude-n-codex.com:8443/v1";

function deckSystemPrompt(topic, detail, count, lang, designBrief) {
  const t = String(topic || "").trim().slice(0, 300);
  const d = String(detail || "").trim().slice(0, 1200);
  const isMy = lang !== "en";
  const languageRules = isMy
    ? [
        "Write descriptive prose in natural Burmese (Myanmar language).",
        "IMPORTANT — do NOT translate everything: keep technical terms, product/brand names,",
        "acronyms (API, AI, CPU, UI, …), and widely-used English loanwords in their ORIGINAL English.",
        "When a term might be unfamiliar, keep the English term and add the Myanmar meaning in",
        'parentheses right after it, e.g. "Kubernetes (ကွန်တိန်နာ စီမံခန့်ခွဲရေးစနစ်)".',
        "This mixed style reads far better than translating every word.",
      ]
    : ["Write ALL slide text in English."];
  return [
    "You are a world-class presentation designer and copywriter.",
    ...languageRules,
    "You are given a presentation topic and must produce a complete slide outline.",
    `The deck must have EXACTLY ${count} content slides (not counting the title slide).`,
    "",
    "Rules for the outline:",
    "- Slide 1 of the outline is the TITLE SLIDE: a short punchy main title (max 10 words),",
    "  a one-line subtitle, and one emoji icon that represents the topic.",
    "- Content slides: short heading (max 8 words) and 3 to 5 punchy bullets.",
    "- EACH BULLET IS ONE PUNCHY DISPLAY LINE (8-12 words) — scannable at a glance,",
    "  with a concrete example, number, or comparison. Slides must NOT look like walls",
    "  of text. Put the deeper explanation (2-3 sentences) in \"notes\" (speaker notes),",
    "  never on the slide itself.",
    "- The last content slide is a strong closing / call-to-action slide.",
    "- Also write 2-3 sentences of speaker notes per slide (the detail lives here).",
    d ? `- Extra context from the user: ${d}` : "",
    "",
    "HIGHLIGHT KEY PHRASES (critical for visual rhythm — never skip):",
    "- Wrap the 1-2 MOST IMPORTANT phrases of EVERY slide in double equals: ==like this==.",
    "  They render in the accent color + bold, so the eye lands on them first.",
    "- Use it in headings, bullets, card titles/text, takeaways, quotes, hero text.",
    "- Never highlight more than 2 phrases per slide; never highlight a whole sentence.",
    '- Example bullet: {"icon":"⚡","text":"Docker image တစ်ခုတည်ဆောက်ပြီး ==နေရာတိုင်းမှာ== run နိုင်သည်"}.',
    "",
    "MEANINGFUL TITLES (critical — never skip this rule):",
    "- Every heading, card title, and stat label must be a COMPLETE, MEANINGFUL phrase",
    "  that makes sense on its own to someone who never saw the source text.",
    "- NEVER copy a raw English fragment from the source/topic text as a title.",
    '- FORBIDDEN titles (meaningless fragments): "1 Image", "N Containers", "0 Code Change", "1 Command".',
    '- GOOD titles instead: "တစ်ကြိမ်တည်ဆောက်၊ အကြိမ်ကြိမ်သုံး", "Environment ပြောင်းလဲစရာမလို",',
    '  "Command တစ်ခုတည်းနဲ့ Run". Card titles: max 6 words, must read as a real heading.',
    "- Stat labels must also be full meaningful phrases (max 10 words), never fragments.",
    "",
    "VISUAL VARIETY (important — avoid text-only monotony):",
    '- Give every slide an "icon": ONE single emoji that visually represents its heading (e.g. 🚀 📊 💡 🌱).',
    '- Give EVERY bullet its own emoji icon too, illustrating that specific bullet —',
    '  "bullets": [{"icon":"⚡","text":"Full informative sentence …"}, …].',
    '- Give every content slide a "takeaway": ONE punchy sentence (max 20 words) — the single',
    "  key message of the slide, specific and memorable. It renders as a highlighted strip.",
    "VISUAL VOCABULARY — layouts are tools, not templates. You are the ART DIRECTOR:",
    "- Shape the layout rhythm to the topic's narrative arc. A \"what is X\" opener earns",
    "  a split or hero; a process earns a timeline; proof points earn stats; features earn cards;",
    "  a comparison earns two-col; a voice worth hearing earns a quote.",
    "- NEVER a flat, uniform sequence — but never force a layout the content doesn't earn.",
    '- Use at least 3 DIFFERENT layouts per deck; "bullets" for at most HALF the slides.',
    '  \u2022 "bullets" — heading + bullet list.',
    '  \u2022 "timeline" — a horizontal 3-4 step process: how it works, workflow, history, evolution.',
    '    Fill "points": [{"title":"Step name (max 6 words)","text":"one punchy line, max 12 words"}, \u2026] (3 or 4).',
    '  \u2022 "split" — illustration art panel on one side, bullets on the other.',
    "    Best for \"what is X\" / definition slides. The deck's bespoke artwork fills the panel —",
    '    give this slide "art" with "at": "left" or "right".',
    '  \u2022 "hero" — ONE massive statement or number as the entire slide message.',
    '    Fill "hero": "==90%== of failures are config errors" (max 10 words) and',
    '    "sub": "one supporting line (max 15 words)". Use at most once per deck,',
    "    for the single most striking insight.",
    '  \u2022 "cards" — 3 feature cards in a row, like a modern SaaS pitch deck.',
    '    Fill "points": [{"icon":"\U0001F517","title":"Complete meaningful heading","text":"1-2 sentences, up to 30 words, with a concrete detail or example"}, \u2026] (exactly 3).',
    "    Titles must follow the MEANINGFUL TITLES rule — never a fragment.",
    '  \u2022 "stats" — a 2x2 grid of big striking numbers on glass cards.',
    '    Fill "stats": [{"value":"30%","label":"full meaningful phrase saying what this number means"}, \u2026] (2 to 4 items).',
    '    Use ONLY for real, striking figures from the topic — percentages, market size, adoption',
    '    numbers, survey results. If the topic has no meaningful numbers, use a DIFFERENT layout;',
    '    NEVER invent trivial counts like 1/2/3 just to fill a stats grid.',
    '  \u2022 "stat" — legacy single big number; prefer "stats" instead.',
    '  \u2022 "two-col" — heading + bullets split into two balanced columns.',
    '  \u2022 "quote" — one memorable quote/statement as the hero; put it in "quote": "..."',
    '    plus optional "quoteBy": "who said it". Use at most once or twice.',
    "",
    "ART DIRECTION — 3 bespoke illustrations were painted for THIS deck and no other (see ARTWORK below).",
    "Cast them across the slides like a magazine art director — imagery is what makes a deck unforgettable:",
    '- Every content slide gets "art": {"n": 0|1|2, "at": "right|left|bg|none"}.',
    '  "right"/"left" = the illustration fills a tall side panel (perfect for bullets, cards, stats).',
    '  "bg" = full-bleed faded backdrop (quote, hero, or very sparse slides ONLY — never behind dense text).',
    '  "none" = let typography carry the slide.',
    "- Give art to at least 60% of slides; never use the same placement twice in a row;",
    "  match each artwork's motif to the slide's subject (a slide about networks gets the network artwork).",
    '- The deck also gets "coverArt": {"n": 0, "at": "bg"} for the title slide (or "none").',
    "",
    "Output STRICT JSON only — no explanations, no markdown fences. The JSON shape:",
    '{ "title": "...", "subtitle": "...", "icon": "🚀", "coverArt": {"n": 0, "at": "bg"}, "slides": [',
    '  { "heading": "...", "bullets": [{"icon":"⚡","text":"..."}], "takeaway": "...", "notes": "...",',
    '    "icon": "📊", "layout": "bullets", "art": {"n": 1, "at": "right"},',
    '    "stats": [{"value": "", "label": ""}],',
    '    "points": [{"icon": "", "title": "", "text": ""}],',
    '    "hero": "", "sub": "",',
    '    "stat": {"value": "", "label": ""}, "quote": "", "quoteBy": "" }',
    "] }",
    `Topic: ${t}`,
    designBrief ? String(designBrief).trim().slice(0, 600) : "",
  ]
    .filter(Boolean)
    .join("\n");
}

async function chatCompletion(baseUrl, apiKey, model, system, user, opts = {}) {
  const url = baseUrl.replace(/\/+$/, "") + "/chat/completions";
  const timeoutMs = opts.timeoutMs || 300000;
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + apiKey,
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: system },
          { role: "user", content: user },
        ],
        temperature: opts.temperature ?? 0.7,
        max_tokens: opts.maxTokens ?? 6000,
      }),
      signal: ctrl.signal,
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`gateway HTTP ${res.status}: ${text.slice(0, 300)}`);
    const data = JSON.parse(text);
    const content = data?.choices?.[0]?.message?.content;
    if (!content) throw new Error("gateway returned no content");
    return content;
  } finally {
    clearTimeout(t);
  }
}

async function listModels(baseUrl, apiKey, timeoutMs = 30000) {
  const url = baseUrl.replace(/\/+$/, "") + "/models";
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + apiKey },
      signal: ctrl.signal,
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`gateway HTTP ${res.status}: ${text.slice(0, 300)}`);
    const data = JSON.parse(text);
    return (data.data || []).map((m) => m.id).filter(Boolean).sort();
  } finally {
    clearTimeout(t);
  }
}

// Pull a JSON object out of a model reply (handles fenced blocks and stray prose).
function extractDeckJson(text) {
  const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const cand = (fence ? fence[1] : text).trim();
  const start = cand.indexOf("{");
  const end = cand.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) return null;
  const slice = cand.slice(start, end + 1);
  try {
    return JSON.parse(slice);
  } catch {
    // try a light repair: trailing commas before } or ]
    const repaired = slice.replace(/,\s*([}\]])/g, "$1");
    try {
      return JSON.parse(repaired);
    } catch {
      return null;
    }
  }
}

const LAYOUTS = ["bullets", "cards", "stats", "stat", "two-col", "quote", "timeline", "split", "hero"];
const EMOJI_RE = /(\u00a9|\u00ae|[\u2000-\u3300]|\ud83c[\ud000-\udfff]|\ud83d[\ud000-\udfff]|\ud83e[\ud000-\udfff])/;

function cleanIcon(v) {
  const m = String(v || "").match(EMOJI_RE);
  return m ? m[0] : "";
}

function cleanBullets(v) {
  return (Array.isArray(v) ? v : [])
    .map((b) => {
      if (b && typeof b === "object")
        return { icon: cleanIcon(b.icon), text: String(b.text || "").trim().slice(0, 220) };
      return { icon: "", text: String(b || "").trim().slice(0, 220) };
    })
    .filter((b) => b.text)
    .slice(0, 6);
}

function bulletTexts(bullets) {
  return (bullets || []).map((b) => (b && typeof b === "object" ? b.text : String(b || "")));
}

function cleanPoints(v) {
  return (Array.isArray(v) ? v : [])
    .map((p) => ({
      icon: cleanIcon(p?.icon),
      title: String(p?.title || "").trim().slice(0, 80),
      text: String(p?.text || "").trim().slice(0, 200),
    }))
    .filter((p) => p.title || p.text)
    .slice(0, 3);
}

function cleanStats(v) {
  return (Array.isArray(v) ? v : [])
    .map((s) => ({
      value: String(s?.value || "").trim().slice(0, 24),
      label: String(s?.label || "").trim().slice(0, 120),
    }))
    .filter((s) => s.value)
    .slice(0, 4);
}

function normalizeDeck(raw, count) {
  if (!raw || typeof raw !== "object") return null;
  const slides = Array.isArray(raw.slides) ? raw.slides : [];
  const norm = slides
    .map((s) => {
      const layout = LAYOUTS.includes(s?.layout) ? s.layout : "bullets";
      const stat = s?.stat && typeof s.stat === "object" ? s.stat : {};
      const stats = cleanStats(s?.stats);
      // legacy single stat -> stats array
      const legacy = String(stat.value || "").trim();
      if (!stats.length && legacy) {
        stats.push({
          value: legacy.slice(0, 24),
          label: String(stat.label || "").trim().slice(0, 120),
        });
      }
      return {
        heading: String(s?.heading || "").trim().slice(0, 120),
        bullets: cleanBullets(s?.bullets),
        takeaway: String(s?.takeaway || "").trim().slice(0, 160),
        notes: String(s?.notes || "").trim().slice(0, 600),
        icon: cleanIcon(s?.icon),
        layout,
        stats,
        points: cleanPoints(s?.points),
        stat: {
          value: legacy.slice(0, 24),
          label: String(stat.label || "").trim().slice(0, 120),
        },
        quote: String(s?.quote || "").trim().slice(0, 300),
        quoteBy: String(s?.quoteBy || "").trim().slice(0, 120),
        hero: String(s?.hero || "").trim().slice(0, 140),
        sub: String(s?.sub || "").trim().slice(0, 200),
        art: s?.art,
      };
    })
    .filter((s) => s.heading || s.bullets.length > 0 || s.quote || s.points.length || s.stats.length || s.hero);
  // enforce exact slide count: trim extras, pad with empty slots if short
  const fixed = norm.slice(0, count);
  while (fixed.length < count)
    fixed.push({ heading: "", bullets: [], takeaway: "", notes: "", icon: "", layout: "bullets", stats: [], points: [], hero: "", sub: "", stat: { value: "", label: "" }, quote: "", quoteBy: "", art: { n: 0, at: "none" } });
  return {
    title: String(raw.title || "").trim().slice(0, 140),
    subtitle: String(raw.subtitle || "").trim().slice(0, 200),
    icon: cleanIcon(raw.icon),
    slides: fixed,
  };
}

// Prompt for the "AI illustrator": paints 3 bespoke SVG artworks for the deck.
// The gateway has no image models, so the LLM paints vector art directly.
function artSystemPrompt(topic, design, lang, avoid) {
  const d = design || {};
  const pal = ["bg", "bgDeep", "band", "accent", "accentSoft"]
    .map((k) => `${k}=#${d[k] || "000000"}`)
    .join(" ");
  const avoidLine =
    Array.isArray(avoid) && avoid.length
      ? "Do NOT repeat these recent motifs, invent clearly different imagery:\n" +
        avoid.map((m) => `- ${m}`).join("\n") + "\n"
      : "";
  return (
    "You are a celebrated vector illustrator. Paint 3 ORIGINAL, premium-quality SVG illustrations " +
    `for a slide deck about "${topic}". These will appear inside slides next to text, so they must be ` +
    "beautiful at a glance and never generic stock-looking.\n" +
    avoidLine +
    `DESIGN IDENTITY to honor — palette (use ONLY these + white/black with transparency): ${pal}. ` +
    `Mood: ${d.mood || "elegant"}. Motif language: ${d.decor || "dots"}.\n` +
    "The 3 artworks must be CLEARLY different from each other:\n" +
    "1. FLOWING ABSTRACT — sweeping shapes, gradients and light evoking the topic's mood.\n" +
    "2. THEMATIC SCENE — simple iconic objects drawn from basic shapes (e.g. containers, waves, " +
    "circuits, leaves) that a viewer instantly connects to the topic.\n" +
    "3. MACRO DETAIL — a close-up geometric pattern or texture, bolder and more graphic.\n" +
    "STRICT technical rules:\n" +
    '- Each artwork is ONE self-contained <svg> with viewBox="0 0 1600 900" and xmlns set.\n' +
    "- FLAT VECTOR style: <rect>, <circle>, <ellipse>, <path>, <polygon> with solid fills and " +
    "<linearGradient>/<radialGradient>. NO <text> elements (fonts cannot be trusted), NO <script>, " +
    "NO event handlers, NO external images or links, NO filters.\n" +
    "- Compose for a 16:9 frame; keep the edges clean; leave some breathing room (not every pixel filled).\n" +
    "- Each svg must be UNDER 12000 characters.\n" +
    (lang === "en"
      ? 'Write "motif" in English (3-6 words describing the imagery).\n'
      : '"motif" ကို မြန်မာလို ၃-၆ လုံးနဲ့ ရေးပါ (ပုံရဲ့ အကြောင်းအရာ).\n') +
    "Return EXACTLY 3 artworks in the array — a response with fewer than 3 is a failure, " +
    "always finish all three. Keep each svg COMPACT (under 6000 characters) so all three fit. " +
    "Reply with ONLY this JSON and nothing else:\n" +
    '{"artworks": [{"motif": "...", "svg": "<svg ...>...</svg>"}, {"motif": "...", "svg": "<svg ...>...</svg>"}, {"motif": "...", "svg": "<svg ...>...</svg>"}]}'
  );
}

// ---- AI artwork placement ---------------------------------------------------
// Each slide may carry art: { n: 0|1|2, at: "left"|"right"|"bg"|"none" }.
const ART_PLACEMENTS = ["left", "right", "bg", "none"];

function cleanArt(v, maxN) {
  const n = Number(v && v.n);
  const at = String((v && v.at) || "none");
  if (!ART_PLACEMENTS.includes(at)) return { n: 0, at: "none" };
  if (!Number.isInteger(n) || n < 0 || n >= maxN) return { n: 0, at: "none" };
  return { n, at };
}

// Layouts that may carry a full-bleed faded art background.
const BG_OK = new Set(["quote", "hero", "split", "bullets", "stat"]);

// Validate + coerce every slide's art choice now that the artwork count is known.
function finalizeArt(deck) {
  if (!deck || !Array.isArray(deck.slides)) return deck;
  const n = (deck.artworks || []).length;
  deck.coverArt = cleanArt(deck.coverArt, n);
  if (!["bg", "none"].includes(deck.coverArt.at)) deck.coverArt.at = "none";
  let prevAt = "none";
  deck.slides.forEach((s) => {
    s.art = cleanArt(s.art, n);
    const L = s.layout;
    if (s.art.at === "bg" && !BG_OK.has(L)) s.art.at = "right"; // dense grids: side panel, not backdrop
    if ((L === "quote" || L === "hero") && (s.art.at === "left" || s.art.at === "right")) s.art.at = "bg";
    if (s.art.at !== "none" && s.art.at === prevAt) {
      // never the same placement twice in a row
      s.art.at = s.art.at === "right" ? "left" : s.art.at === "left" ? "right" : "none";
    }
    prevAt = s.art.at;
  });
  return deck;
}
function detailsSystemPrompt(topic, lang) {
  const langLine = lang === "en"
    ? "Write the brief in English."
    : "အသေးစိတ်ကို မြန်မာလိုရေး — နည်းပညာအခေါ်အဝေါ်၊ product နာမည်၊ acronym တွေကို အင်္ဂလိပ်လိုထားပြီး မြန်မာအနက်ကို ကွင်းစကွင်းပိတ်ထဲထည့် (ဥပမာ — Docker (ကွန်တိန်နာ နည်းပညာ))။";
  return (
    "You help a user expand a presentation topic into a rich detail brief. " +
    "The brief will guide an AI that writes the actual slide deck, so make it concrete and inspiring.\n" +
    "Cover in 4-6 short sentences: what the topic is really about, 2-3 key aspects worth including, " +
    "a suggested audience angle (who this is for), and the tone/purpose (e.g. convince, teach, inspire).\n" +
    langLine + "\n" +
    "Rules: plain sentences only, no JSON, no markdown, no bullet points, no quotes around the text, under 700 characters."
  );
}

// Prompt for the "AI design director": invents a COMPLETELY NEW visual identity
// for every deck — never picks from a list, never repeats a generic template.
function designSystemPrompt(topic, detail, lang, opts) {
  const langLine = lang === "en"
    ? 'Write "name" and "reason" in English.'
    : '"name" နဲ့ "reason" ကို မြန်မာလို ရေးပါ (name က ၂-၄ လုံးပါတဲ့ ဆန်းသစ်တဲ့ နာမည်).';
  const o = opts || {};
  const banned = Array.isArray(o.banned) ? o.banned.filter((b) => b && b.bg) : [];
  let variety = "";
  if (banned.length) {
    variety +=
      "CRITICAL — these visual identities were used very recently. DO NOT repeat them or stay close to them:\n" +
      banned.map((b) => `- "${b.name || "?"}": background #${b.bg}, accent #${b.accent || "?"}`).join("\n") +
      "\nYour background hue family must be OBVIOUSLY different from every one above. " +
      "If they are dark, go somewhere totally different.\n";
  }
  if (o.light) {
    variety +=
      "This time design a LIGHT theme: warm paper/cream background (e.g. #FAF6EF), dark ink text (#1A2333), " +
      "one vivid accent. Make it feel editorial and premium — a striking change of pace.\n";
  } else {
    variety +=
      "Background variety is the #1 priority: rotate between deep teal, charcoal black, warm cream, midnight blue, " +
      "forest green, wine red, pure ink, sunset amber... NEVER default to purple/violet backgrounds.\n";
  }
  return (
    "You are an award-winning presentation art director. Invent a COMPLETELY NEW, original visual identity for a slide deck about the topic below. " +
    "This must look like nothing the audience has seen before — a fresh design every single time.\n" +
    variety +
    `Topic: ${topic}\n` +
    (detail ? `Extra context: ${String(detail).slice(0, 400)}\n` : "") +
    "Rules for originality:\n" +
    "- Do NOT fall back to generic corporate blue, plain dark-navy+gold, or boring minimal white. Be daring but tasteful.\n" +
    "- Choose an unexpected but harmonious palette: deep plum + lime, ink black + sakura pink, forest green + amber, " +
    "ocean teal + coral, charcoal + electric violet, midnight + champagne, etc. Dark backgrounds usually look more premium.\n" +
    "- bg = slide background, bgDeep = darker shade for title/closing slides, band = muted surface color for shapes/footers, " +
    "accent = the one signature color (used for highlights, icons, rules), accentSoft = lighter tint of accent, " +
    "title/text = must contrast strongly with bg, muted = secondary text, footer = faint text.\n" +
    "- titleStyle: 'monument' = gigantic typography statement; 'band' = bold vertical accent band layout; 'halo' = centered with a large accent ring.\n" +
    "- corners: 'sharp' = edgy rectangles, 'soft' = gentle rounding, 'round' = very rounded friendly cards.\n" +
    "- decor: 'dots' = dot-grid motif, 'streaks' = diagonal light streaks, 'rings' = large thin circles, 'none' = pure minimal.\n" +
    "- mood: energetic (fast, punchy), elegant (smooth, calm), bold (dramatic), calm (gentle) — match the topic's energy.\n" +
    "- headerStyle: 'kicker' = number pill + kicker rule + title; 'numeral' = giant ghost slide number behind the title; 'tab' = accent side-tab marker + title.\n" +
    "- bulletStyle: 'chips' = bullets on a glass card with icon chips; 'numerals' = no card, big accent numerals 01 02 03 lead each bullet; 'rules' = no card, minimal hairline separators.\n" +
    "- cardStyle: 'glass' = translucent glass cards; 'solid' = solid filled cards with thick top accent bar; 'outline' = transparent cards with accent outline only.\n" +
    "- statStyle: 'cards' = stats on glass cards; 'giant' = no cards, enormous numbers with thin dividers; 'bands' = full-width rows with accent side bands.\n" +
    "- Keep the four style picks COHERENT with each other and with corners/mood (e.g. sharp corners pair with numerals/rules; round corners pair with chips/glass).\n" +
    langLine + "\n" +
    "Reply with ONLY this JSON and nothing else:\n" +
    '{"name":"<original 2-4 word design name>","palette":{"bg":"RRGGBB","bgDeep":"RRGGBB","band":"RRGGBB","accent":"RRGGBB","accentSoft":"RRGGBB","title":"RRGGBB","text":"RRGGBB","muted":"RRGGBB","footer":"RRGGBB"},' +
    '"titleStyle":"monument|band|halo","corners":"sharp|soft|round","decor":"dots|streaks|rings|none","mood":"energetic|elegant|bold|calm",' +
    '"headerStyle":"kicker|numeral|tab","bulletStyle":"chips|numerals|rules","cardStyle":"glass|solid|outline","statStyle":"cards|giant|bands",' +
    '"reason":"<one sentence: why this design fits the topic>"}'
  );
}

module.exports = {
  DEFAULT_BASE_URL,
  deckSystemPrompt,
  designSystemPrompt,
  detailsSystemPrompt,
  artSystemPrompt,
  cleanArt,
  finalizeArt,
  chatCompletion,
  listModels,
  extractDeckJson,
  normalizeDeck,
};
