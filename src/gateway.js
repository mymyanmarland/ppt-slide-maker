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
    '- Give every slide a "layout" — visual variety is MANDATORY, never a text-only deck:',
    '  • "bullets" — heading + bullet list. Use for at most HALF the slides.',
    '  • "timeline" — a horizontal 3-4 step process: how it works, workflow, history, evolution.',
    '    Fill "points": [{"title":"Step name (max 6 words)","text":"one punchy line, max 12 words"}, …] (3 or 4).',
    '  • "split" — giant icon art panel on the left, bullets on the right.',
    '    Best for "what is X" / definition slides with a strong single icon.',
    '  • "hero" — ONE massive statement or number as the entire slide message.',
    '    Fill "hero": "==90%== of failures are config errors" (max 10 words) and',
    '    "sub": "one supporting line (max 15 words)". Use EXACTLY once per deck,',
    "    for the single most striking insight.",
    '  • "cards" — 3 feature cards in a row, like a modern SaaS pitch deck.',
    '    Fill "points": [{"icon":"🔗","title":"Complete meaningful heading","text":"1-2 sentences, up to 30 words, with a concrete detail or example"}, …] (exactly 3).',
    "    Titles must follow the MEANINGFUL TITLES rule — never a fragment.",
    '  • "stats" — a 2x2 grid of big striking numbers on glass cards.',
    '    Fill "stats": [{"value":"30%","label":"full meaningful phrase saying what this number means"}, …] (2 to 4 items).',
    '    Use ONLY for real, striking figures from the topic — percentages, market size, adoption',
    '    numbers, survey results. If the topic has no meaningful numbers, use a DIFFERENT layout;',
    '    NEVER invent trivial counts like 1/2/3 just to fill a stats grid.',
    '  • "stat" — legacy single big number; prefer "stats" instead.',
    '  • "two-col" — heading + bullets split into two balanced columns.',
    '  • "quote" — one memorable quote/statement as the hero; put it in "quote": "..."',
    '    plus optional "quoteBy": "who said it". Use at most once or twice.',
    "- COMPOSITION RULE: every deck must use at least 3 DIFFERENT layouts.",
    '  Include at least 1 "timeline" or "split", exactly 1 "hero",',
    '  at least 1 "stats" grid and at least 1 "cards" layout.',
    "",
    "Output STRICT JSON only — no explanations, no markdown fences. The JSON shape:",
    '{ "title": "...", "subtitle": "...", "icon": "🚀", "slides": [',
    '  { "heading": "...", "bullets": [{"icon":"⚡","text":"..."}], "takeaway": "...", "notes": "...",',
    '    "icon": "📊", "layout": "bullets",',
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
      };
    })
    .filter((s) => s.heading || s.bullets.length > 0 || s.quote || s.points.length || s.stats.length || s.hero);
  // enforce exact slide count: trim extras, pad with empty slots if short
  const fixed = norm.slice(0, count);
  while (fixed.length < count)
    fixed.push({ heading: "", bullets: [], takeaway: "", notes: "", icon: "", layout: "bullets", stats: [], points: [], hero: "", sub: "", stat: { value: "", label: "" }, quote: "", quoteBy: "" });
  return {
    title: String(raw.title || "").trim().slice(0, 140),
    subtitle: String(raw.subtitle || "").trim().slice(0, 200),
    icon: cleanIcon(raw.icon),
    slides: fixed,
  };
}

// Prompt that expands a bare topic into a rich "details" brief for deck generation.
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

// Prompt for the "AI design director": analyzes the topic and picks the visual design.
function designSystemPrompt(topic, detail, lang, themeKeys) {
  const langLine = lang === "en"
    ? 'Write "reason" in English, one short sentence.'
    : '"reason" ကို မြန်မာလို တစ်ကြောင်းတည်း တိုတိုရေးပါ။';
  return (
    "You are an art director for slide presentations. Analyze the topic below and choose the visual design that fits it best.\n" +
    `Topic: ${topic}\n` +
    (detail ? `Extra context: ${String(detail).slice(0, 400)}\n` : "") +
    `Available theme keys: ${themeKeys.join(", ")}.\n` +
    "Theme guide: midnight-glass = sleek modern tech; violet = creative/futuristic tech; navy-gold = premium/luxury; " +
    "corporate = business/finance; minimal = clean formal/education; impact = bold, loud, youth/entertainment; " +
    "myanmar = Myanmar culture, tradition, food, festivals.\n" +
    "Also choose a transition mood for the slide animations: energetic (fast, punchy), elegant (smooth, calm), bold (dramatic), or calm (gentle) — match it to the topic's energy.\n" +
    langLine + "\n" +
    'Reply with ONLY this JSON and nothing else: {"theme":"<one key from the list>","mood":"energetic|elegant|bold|calm","reason":"<why this theme fits the topic>"}.'
  );
}

module.exports = {
  DEFAULT_BASE_URL,
  deckSystemPrompt,
  designSystemPrompt,
  detailsSystemPrompt,
  chatCompletion,
  listModels,
  extractDeckJson,
  normalizeDeck,
};
