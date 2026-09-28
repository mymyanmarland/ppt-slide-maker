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
const HEADER_STYLES = ["kicker", "numeral", "tab"];
const BULLET_STYLES = ["chips", "numerals", "rules"];
const CARD_STYLES = ["glass", "solid", "outline"];
const STAT_STYLES = ["cards", "giant", "bands"];

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
    headerStyle: cleanEnum(r.headerStyle, HEADER_STYLES, "kicker"),
    bulletStyle: cleanEnum(r.bulletStyle, BULLET_STYLES, "chips"),
    cardStyle: cleanEnum(r.cardStyle, CARD_STYLES, "glass"),
    statStyle: cleanEnum(r.statStyle, STAT_STYLES, "cards"),
    radius: corners === "sharp" ? 0.02 : corners === "round" ? 0.24 : 0.12,
    reason: String(r.reason || "").trim().slice(0, 200),
  };
  return theme;
}

// Accept a theme object (new decks) or a legacy theme key (old decks).
function resolveTheme(t, legacyThemes) {
  let th;
  if (t && typeof t === "object" && t.bg && t.accent) th = t;
  else if (typeof t === "string" && legacyThemes && legacyThemes[t]) th = legacyThemes[t];
  else th = { ...FALLBACK_COLORS, name: "Fresh Design", titleStyle: "monument", corners: "soft", decor: "none", mood: "elegant", headerStyle: "kicker", bulletStyle: "chips", cardStyle: "glass", statStyle: "cards", radius: 0.12, reason: "" };
  // Fill style-DNA defaults for legacy themes that predate them.
  return { headerStyle: "kicker", bulletStyle: "chips", cardStyle: "glass", statStyle: "cards", ...th };
}

function hexRgb(h) {
  const s = String(h || "").replace("#", "");
  return [parseInt(s.slice(0, 2), 16) || 0, parseInt(s.slice(2, 4), 16) || 0, parseInt(s.slice(4, 6), 16) || 0];
}

function hexDist(a, b) {
  const [r1, g1, b1] = hexRgb(a), [r2, g2, b2] = hexRgb(b);
  return Math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2);
}

// Hue-aware background similarity: two backgrounds are "too close" when they share
// a similar hue AND similar lightness (or both are near-grayscale with similar lightness).
function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  const l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn;
  const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  let h = 0;
  if (mx === r) h = ((g - b) / d + (g < b ? 6 : 0));
  else if (mx === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h * 60, s, l];
}

function bgTooClose(bg, recent) {
  const [h1, s1, l1] = rgbToHsl(...hexRgb(bg));
  return (recent || []).some((r) => {
    if (!r || !r.bg) return false;
    const [h2, s2, l2] = rgbToHsl(...hexRgb(r.bg));
    const hueDiff = Math.abs(h1 - h2) > 180 ? 360 - Math.abs(h1 - h2) : Math.abs(h1 - h2);
    const lightDiff = Math.abs(l1 - l2);
    const dark1 = l1 < 0.16, dark2 = l2 < 0.16;
    if (dark1 && dark2) {
      // both very dark: need a clearly different hue to count as different
      if (s1 < 0.12 || s2 < 0.12) return lightDiff < 0.1;
      return hueDiff < 60 && lightDiff < 0.12;
    }
    if (s1 < 0.12 && s2 < 0.12) return lightDiff < 0.16; // both grayscale: compare lightness only
    if (s1 < 0.12 || s2 < 0.12) return false; // one colorful, one gray -> different enough
    return hueDiff < 32 && lightDiff < 0.2;
  });
}

module.exports = { normalizeDesign, resolveTheme, FALLBACK_COLORS, MOODS, hexDist, bgTooClose };
