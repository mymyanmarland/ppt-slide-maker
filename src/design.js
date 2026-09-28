"use strict";
// AI-generated visual identities: validation + normalization + internal fallback.
// A "theme" here is a flat color object shaped like the old THEMES entries
// (bg, bgDeep, band, accent, accentSoft, title, text, muted, footer, glass, glassBorder)
// plus design knobs: titleStyle, corners, decor, radius.

const HEX6 = /^[0-9A-Fa-f]{6}$/;

const TITLE_STYLES = ["monument", "band", "halo"];
const CORNERS = ["sharp", "soft", "round"];
const DECORS = ["dots", "streaks", "rings", "none"];
const MOODS = ["energetic", "elegant", "bold", "calm"];

// Internal fallback only — never shown in the UI. Used when the AI returns garbage.
const FALLBACK_COLORS = {
  bg: "0B1220", bgDeep: "060B16", band: "16233D",
  accent: "5EB1FF", accentSoft: "A8D4FF",
  title: "FFFFFF", text: "E5E7EB", muted: "8CA3C4", footer: "54687F",
  glass: "FFFFFF", glassBorder: "2E435F",
};

function cleanHex(v, fb) {
  const s = String(v || "").trim().replace(/^#/, "");
  return HEX6.test(s) ? s.toUpperCase() : fb;
}

function cleanEnum(v, list, fb) {
  return list.includes(v) ? v : fb;
}

// raw = parsed JSON from the design director. Returns a full theme object.
function normalizeDesign(raw, lang) {
  const r = raw && typeof raw === "object" ? raw : {};
  const p = r.palette && typeof r.palette === "object" ? r.palette : {};
  const colors = {
    bg: cleanHex(p.bg, FALLBACK_COLORS.bg),
    bgDeep: cleanHex(p.bgDeep, FALLBACK_COLORS.bgDeep),
    band: cleanHex(p.band, FALLBACK_COLORS.band),
    accent: cleanHex(p.accent, FALLBACK_COLORS.accent),
    accentSoft: cleanHex(p.accentSoft, FALLBACK_COLORS.accentSoft),
    title: cleanHex(p.title, FALLBACK_COLORS.title),
    text: cleanHex(p.text, FALLBACK_COLORS.text),
    muted: cleanHex(p.muted, FALLBACK_COLORS.muted),
    footer: cleanHex(p.footer, FALLBACK_COLORS.footer),
    glass: "FFFFFF",
    glassBorder: cleanHex(p.band, FALLBACK_COLORS.band),
  };
  const corners = cleanEnum(r.corners, CORNERS, "soft");
  const theme = {
    ...colors,
    name: String(r.name || (lang === "en" ? "Fresh Design" : "ဒီဇိုင်းအသစ်")).trim().slice(0, 40) || "Fresh Design",
    titleStyle: cleanEnum(r.titleStyle, TITLE_STYLES, "monument"),
    corners,
    decor: cleanEnum(r.decor, DECORS, "none"),
    mood: cleanEnum(r.mood, MOODS, "elegant"),
    radius: corners === "sharp" ? 0.02 : corners === "round" ? 0.24 : 0.12,
    reason: String(r.reason || "").trim().slice(0, 200),
  };
  return theme;
}

// Accept a theme object (new decks) or a legacy theme key (old decks).
function resolveTheme(t, legacyThemes) {
  if (t && typeof t === "object" && t.bg && t.accent) return t;
  if (typeof t === "string" && legacyThemes && legacyThemes[t]) return legacyThemes[t];
  return { ...FALLBACK_COLORS, name: "Fresh Design", titleStyle: "monument", corners: "soft", decor: "none", mood: "elegant", radius: 0.12, reason: "" };
}

module.exports = { normalizeDesign, resolveTheme, FALLBACK_COLORS, MOODS };
