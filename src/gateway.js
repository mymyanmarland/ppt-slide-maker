"use strict";
// Client for the user's OpenAI-compatible gateway + deck-outline generation.

const DEFAULT_BASE_URL = "https://claude-n-codex.com:8443/v1";

function deckSystemPrompt(topic, detail, count, lang) {
  const t = String(topic || "").trim().slice(0, 300);
  const d = String(detail || "").trim().slice(0, 1200);
  const languageName = lang === "en" ? "English" : "Burmese (Myanmar language)";
  return [
    "You are a world-class presentation designer and copywriter.",
    `Write ALL slide text in ${languageName}.`,
    "You are given a presentation topic and must produce a complete slide outline.",
    `The deck must have EXACTLY ${count} content slides (not counting the title slide).`,
    "",
    "Rules for the outline:",
    "- Slide 1 of the outline is the TITLE SLIDE: give a short punchy main title (max 10 words) and a one-line subtitle.",
    "- Content slides: each has a short heading (max 8 words) and 3 to 5 crisp bullet points (each bullet max 18 words).",
    "- Bullets must be specific, insightful, and non-repetitive — no filler like \"In conclusion\".",
    "- The last content slide should be a strong closing / call-to-action slide.",
    "- Also write 1-2 sentences of speaker notes per slide.",
    d ? `- Extra context from the user: ${d}` : "",
    "",
    "Output STRICT JSON only — no explanations, no markdown fences. The JSON shape:",
    '{ "title": "...", "subtitle": "...", "slides": [',
    '  { "heading": "...", "bullets": ["...", "..."], "notes": "..." }',
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

function normalizeDeck(raw, count) {
  if (!raw || typeof raw !== "object") return null;
  const slides = Array.isArray(raw.slides) ? raw.slides : [];
  const norm = slides
    .map((s) => ({
      heading: String(s?.heading || "").trim().slice(0, 120),
      bullets: (Array.isArray(s?.bullets) ? s.bullets : [])
        .map((b) => String(b || "").trim().slice(0, 300))
        .filter(Boolean)
        .slice(0, 6),
      notes: String(s?.notes || "").trim().slice(0, 600),
    }))
    .filter((s) => s.heading || s.bullets.length > 0);
  // enforce exact slide count: trim extras, pad with empty slots if short
  const fixed = norm.slice(0, count);
  while (fixed.length < count) fixed.push({ heading: "", bullets: [], notes: "" });
  return {
    title: String(raw.title || "").trim().slice(0, 140),
    subtitle: String(raw.subtitle || "").trim().slice(0, 200),
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
