"use strict";
// AI illustrator pipeline: the gateway paints 3 bespoke SVG artworks per deck,
// we sanitize them and rasterize to PNG (data URLs) for embedding in PPTX + preview.
const gw = require("./gateway");

// Strip anything a rasterizer or the browser shouldn't see: scripts, event
// handlers, external refs, and <text> (fonts can't be trusted server-side).
function sanitizeSvg(svg) {
  let s = String(svg || "");
  const start = s.indexOf("<svg");
  const end = s.lastIndexOf("</svg>");
  if (start === -1 || end === -1 || end <= start) return null;
  s = s.slice(start, end + 6);
  if (s.length > 25000) return null;
  s = s
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, "")
    .replace(/<text[\s\S]*?<\/text>/gi, "")
    .replace(/<text\b[^>]*\/>/gi, "")
    .replace(/\s(?:xlink:)?href\s*=\s*"(?!#)[^"]*"/gi, "")
    .replace(/url\(\s*https?:[^)]*\)/gi, "url(#x)");
  if (!s.includes("<svg")) return null;
  if (!/xmlns=/.test(s)) s = s.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  if (!/viewBox=/.test(s)) s = s.replace("<svg", '<svg viewBox="0 0 1600 900"');
  return s;
}

// SVG -> PNG data URL via sharp (lazy require: graceful if unavailable).
async function rasterize(svg, w = 1280, h = 720) {
  const sharp = require("sharp");
  const buf = await sharp(Buffer.from(svg), { density: 110 })
    .resize(w, h, { fit: "fill" })
    .png({ compressionLevel: 8 })
    .toBuffer();
  return "data:image/png;base64," + buf.toString("base64");
}

async function generateArtworks({ topic, detail, design, lang, baseUrl, apiKey, model, avoidMotifs }) {
  const raw = await gw.chatCompletion(
    baseUrl,
    apiKey,
    model,
    gw.artSystemPrompt(topic, design, lang, avoidMotifs),
    `Topic: ${topic}${detail ? `\nContext: ${String(detail).slice(0, 300)}` : ""}`,
    { maxTokens: 12000, temperature: 1.0 }
  );
  const parsed = gw.extractDeckJson(raw);
  const list = parsed && Array.isArray(parsed.artworks) ? parsed.artworks : [];
  const out = [];
  for (const a of list.slice(0, 3)) {
    const svg = sanitizeSvg(a && a.svg);
    if (!svg) continue;
    let png = null;
    try {
      png = await rasterize(svg);
    } catch (e) {
      console.warn("[art] rasterize failed:", String((e && e.message) || e).slice(0, 120));
    }
    if (!png) continue;
    out.push({ motif: String((a && a.motif) || "abstract art").trim().slice(0, 80) || "abstract art", png });
  }
  return out;
}

module.exports = { sanitizeSvg, rasterize, generateArtworks };
