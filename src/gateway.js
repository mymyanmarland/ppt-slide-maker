"use strict";
// Client for the user's OpenAI-compatible gateway + deck-outline generation.

const DEFAULT_BASE_URL = "https://claude-n-codex.com:8443/v1";

function deckSystemPrompt(topic, detail, count, lang) {
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
    "- Content slides: short heading (max 8 words) and 3 to 4 crisp bullets,",
    "  EACH BULLET MAX 12 WORDS. Short beats long — no walls of text.",
    "- The last content slide is a strong closing / call-to-action slide.",
    "- Also write 1-2 sentences of speaker notes per slide.",
    d ? `- Extra context from the user: ${d}` : "",
    "",
    "VISUAL VARIETY (important — avoid text-only monotony):",
    '- Give every slide an "icon": ONE single emoji that visually represents its heading (e.g. 🚀 📊 💡 🌱).',
    '- Give every slide a "layout", varying across the deck:',
    '  • "bullets" — heading + bullet list (default, use for ~half the slides)',
    '  • "stat" — one big striking number/fact + its label + 2 short bullets.',
    '    Also fill "stat": {"value": "70%", "label": "what this number means"}. Use for key figures.',
    '  • "two-col" — heading + bullets split into two balanced columns.',
    '  • "quote" — one memorable quote/statement as the hero; put it in "quote": "..."',
    '    plus optional "quoteBy": "who said it". Use at most once or twice.',
    "- Use at least 2 \"stat\" slides and at least 1 non-bullets layout in every deck.",
    "",
    "Output STRICT JSON only — no explanations, no markdown fences. The JSON shape:",
    '{ "title": "...", "subtitle": "...", "icon": "🚀", "slides": [',
    '  { "heading": "...", "bullets": ["...", "..."], "notes": "...",',
    '    "icon": "📊", "layout": "bullets",',
    '    "stat": {"value": "", "label": ""}, "quote": "", "quoteBy": "" }',
    "] }",
    `Topic: ${t}`,
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

const LAYOUTS = ["bullets", "stat", "two-col", "quote"];
const EMOJI_RE = /(\u00a9|\u00ae|[\u2000-\u3300]|\ud83c[\ud000-\udfff]|\ud83d[\ud000-\udfff]|\ud83e[\ud000-\udfff])/;

function cleanIcon(v) {
  const m = String(v || "").match(EMOJI_RE);
  return m ? m[0] : "";
}

function normalizeDeck(raw, count) {
  if (!raw || typeof raw !== "object") return null;
  const slides = Array.isArray(raw.slides) ? raw.slides : [];
  const norm = slides
    .map((s) => {
      const layout = LAYOUTS.includes(s?.layout) ? s.layout : "bullets";
      const stat = s?.stat && typeof s.stat === "object" ? s.stat : {};
      return {
        heading: String(s?.heading || "").trim().slice(0, 120),
        bullets: (Array.isArray(s?.bullets) ? s.bullets : [])
          .map((b) => String(b || "").trim().slice(0, 200))
          .filter(Boolean)
          .slice(0, 6),
        notes: String(s?.notes || "").trim().slice(0, 600),
        icon: cleanIcon(s?.icon),
        layout,
        stat: {
          value: String(stat.value || "").trim().slice(0, 24),
          label: String(stat.label || "").trim().slice(0, 120),
        },
        quote: String(s?.quote || "").trim().slice(0, 300),
        quoteBy: String(s?.quoteBy || "").trim().slice(0, 120),
      };
    })
    .filter((s) => s.heading || s.bullets.length > 0 || s.quote);
  // enforce exact slide count: trim extras, pad with empty slots if short
  const fixed = norm.slice(0, count);
  while (fixed.length < count)
    fixed.push({ heading: "", bullets: [], notes: "", icon: "", layout: "bullets", stat: { value: "", label: "" }, quote: "", quoteBy: "" });
  return {
    title: String(raw.title || "").trim().slice(0, 140),
    subtitle: String(raw.subtitle || "").trim().slice(0, 200),
    icon: cleanIcon(raw.icon),
    slides: fixed,
  };
}

module.exports = {
  DEFAULT_BASE_URL,
  deckSystemPrompt,
  chatCompletion,
  listModels,
  extractDeckJson,
  normalizeDeck,
};
