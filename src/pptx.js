"use strict";
// Deck JSON -> real .pptx bytes via pptxgenjs, with rich layouts + slide transitions.
const PptxGenJS = require("pptxgenjs");
const AdmZip = require("adm-zip");
const { THEMES, THEME_KEYS } = require("./themes");

const MY_RE = /[\u1000-\u109F]/;
// Pyidaungsu is the standard Myanmar Unicode font on the user's machines;
// PowerPoint falls back gracefully if it is not installed.
const fontFor = (s) => (MY_RE.test(String(s || "")) ? "Pyidaungsu" : "Calibri");
const isLight = (theme) => theme.bgDeep === "F4F5F7";

// ---- typography system ------------------------------------------------------
// Headings: Bold (700), tight letter-spacing (-0.02em), line-height 1.1
//   H1 44pt (title slide) / H2 32pt (content headings) / H3 24pt (subtitle, card titles)
// Body: Regular (400), 16-18pt, line-height 1.6, letter-spacing 0, bullet "•"
//   (Medium 500 is not expressible in pptx — emphasis is done via italic/color)
// Caption: Regular, 12pt, #9CA3AF gray, italic (footers, attributions, dates)
const CAPTION_COLOR = "9CA3AF";
const tight = (pt) => Math.round(pt * -0.02 * 10) / 10; // -0.02em expressed in points
const T = {
  h1: (color) => ({ fontSize: 44, bold: true, color, charSpacing: tight(44), lineSpacingMultiple: 1.1 }),
  h2: (color) => ({ fontSize: 32, bold: true, color, charSpacing: tight(32), lineSpacingMultiple: 1.1 }),
  h3: (color) => ({ fontSize: 24, bold: true, color, charSpacing: tight(24), lineSpacingMultiple: 1.1 }),
  quote: (color) => ({ fontSize: 24, italic: true, color, lineSpacingMultiple: 1.3 }),
  caption: () => ({ fontSize: 12, italic: true, color: CAPTION_COLOR }),
};

// ---- shared bits -----------------------------------------------------------

// Soft ambient light: one accent-tinted glow + one neutral wash. Add FIRST.
function ambient(slide, pptx, theme) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 8.2, y: -3.4, w: 8.0, h: 8.0,
    fill: { color: theme.accent, transparency: 90 }, line: { color: theme.accent, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: -3.2, y: 4.4, w: 6.4, h: 6.4,
    fill: { color: theme.band, transparency: 82 }, line: { color: theme.band, transparency: 100 },
  });
}

// Frosted glass card: translucent fill + hairline border + top sheen highlight.
function glassCard(slide, pptx, theme, x, y, w, h, r) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: r == null ? 0.14 : r,
    fill: { color: theme.glass || "FFFFFF", transparency: 90 },
    line: { color: theme.glassBorder || theme.band, width: 1 },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.12, y: y + 0.07, w: Math.max(0.2, w - 0.24), h: 0.05, rectRadius: 0.025,
    fill: { color: "FFFFFF", transparency: 72 }, line: { color: "FFFFFF", transparency: 100 },
  });
}

function topBar(slide, pptx, theme) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.09, fill: { color: theme.accent }, line: { color: theme.accent },
  });
}

function numberPill(slide, pptx, theme, n) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55, y: 0.42, w: 0.72, h: 0.42, rectRadius: 0.21,
    fill: { color: theme.accent }, line: { color: theme.accent },
  });
  slide.addText(String(n).padStart(2, "0"), {
    x: 0.55, y: 0.42, w: 0.72, h: 0.42, align: "center",
    fontSize: 14, bold: true, color: isLight(theme) ? "FFFFFF" : theme.bg, fontFace: "Calibri",
  });
}

function headingBlock(slide, pptx, theme, heading) {
  const base = { ...T.h2(theme.title), fontFace: fontFor(heading) };
  const runs = richText(heading || "", theme, base);
  slide.addText(runs.length ? runs : "", {
    x: 1.5, y: 0.32, w: 11.3, h: 1.1,
    ...T.h2(theme.title), fontFace: fontFor(heading), valign: "middle",
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.52, y: 1.32, w: 1.1, h: 0.05, fill: { color: theme.accent }, line: { color: theme.accent },
  });
}

// Big translucent-feel icon watermark at the right edge (kept clear of text).
function watermark(slide, theme, icon) {
  if (!icon) return;
  slide.addText(icon, {
    x: 10.7, y: 2.1, w: 2.3, h: 2.3, align: "center",
    fontSize: 96, color: theme.band, fontFace: "Segoe UI Emoji",
  });
}

function bulletRuns(bullets, theme, fontSize = 18) {
  return (bullets || [])
    .map((b) => {
      const icon = b && typeof b === "object" ? String(b.icon || "") : "";
      const text = b && typeof b === "object" ? String(b.text || "") : String(b || "");
      return { icon, text };
    })
    .filter((b) => b.text)
    .flatMap((b) =>
      richText(
        (b.icon ? b.icon + "  " : "") + b.text,
        theme,
        { fontSize, color: theme.text, fontFace: fontFor(b.text) },
        {
          bullet: { code: "2022", color: theme.accent, indent: 20 },
          paraSpaceAfter: 8,
          lineSpacingMultiple: 1.6,
        }
      )
    );
}

// Text of a bullet whether it is a string or {icon, text}.
function btext(b) {
  return b && typeof b === "object" ? String(b.text || "") : String(b || "");
}
function bicon(b) {
  return b && typeof b === "object" ? String(b.icon || "") : "";
}

// Keyword highlighting: "==phrase==" markers become accent-colored bold runs.
// Returns an array of {text, options} for addText; runs merge into ONE paragraph
// via breakLine:false on all but the last entry (paragraph opts live on the first).
function richText(text, theme, runBase, paraOpts = {}) {
  const parts = String(text || "").split(/(==.+?==)/g).filter((p) => p !== "");
  return parts.map((p, i) => {
    const hl = p.length > 4 && p.startsWith("==") && p.endsWith("==");
    return {
      text: hl ? p.slice(2, -2) : p,
      options: {
        ...runBase,
        ...(hl ? { color: theme.accent, bold: true } : {}),
        ...(i === 0 ? paraOpts : {}),
        ...(i < parts.length - 1 ? { breakLine: false } : {}),
      },
    };
  });
}

// Highlighted takeaway strip: one punchy key message at the bottom of a slide.
function takeawayStrip(slide, pptx, theme, text, x, y, w) {
  if (!text) return;
  const h = 0.62;
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.31,
    fill: { color: theme.glass || "FFFFFF", transparency: 88 },
    line: { color: theme.accent, width: 1.25 },
  });
  const base = {
    fontSize: 16, italic: true, color: theme.title,
    fontFace: fontFor(text), lineSpacingMultiple: 1.1,
  };
  const runs = richText(text, theme, base);
  if (runs.length) runs[0].text = "✦  " + runs[0].text;
  slide.addText(runs.length ? runs : "", {
    x: x + 0.4, y: y + 0.02, w: w - 0.8, h: h - 0.04, valign: "middle",
    ...base,
  });
}

function addFooter(slide, pptx, theme, left, right) {
  // footers never show highlight markers
  const clean = String(left || "").replace(/==/g, "");
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: 6.94, w: 12.33, h: 0.025, fill: { color: theme.band }, line: { color: theme.band },
  });
  slide.addText(clean, {
    x: 0.5, y: 7.02, w: 8, h: 0.3,
    ...T.caption(), fontFace: fontFor(clean),
  });
  slide.addText(String(right || ""), {
    x: 11.5, y: 7.02, w: 1.33, h: 0.3, align: "right",
    ...T.caption(), fontFace: "Calibri",
  });
}

// ---- slide layouts ---------------------------------------------------------

function addTitleSlide(pptx, theme, deck, lang) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  s.addShape(pptx.ShapeType.ellipse, {
    x: 9.4, y: -2.6, w: 7.2, h: 7.2,
    fill: { color: theme.band, transparency: 35 }, line: { color: theme.band, transparency: 100 },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: -2.4, y: 5.4, w: 5.2, h: 5.2,
    fill: { color: theme.band, transparency: 55 }, line: { color: theme.band, transparency: 100 },
  });
  // thin accent ring
  s.addShape(pptx.ShapeType.ellipse, {
    x: 10.6, y: 4.4, w: 2.6, h: 2.6,
    fill: { color: theme.bgDeep, transparency: 100 }, line: { color: theme.accent, width: 2 },
  });
  if (deck.icon) {
    s.addText(deck.icon, {
      x: 10.55, y: 4.35, w: 2.7, h: 2.7, align: "center",
      fontSize: 96, fontFace: "Segoe UI Emoji",
    });
  }
  const kicker = lang === "en" ? "AI SLIDE DECK" : "AI ဆလိုက်ဒ်";
  s.addText(kicker, {
    x: 0.9, y: 1.7, w: 6, h: 0.4,
    fontSize: 13, bold: true, color: theme.accent, charSpacing: 6, fontFace: fontFor(kicker),
  });
  // double rule: thick accent block + thin hairline
  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 2.18, w: 1.4, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 2.42, y: 2.2, w: 3.2, h: 0.025, fill: { color: theme.band }, line: { color: theme.band },
  });
  s.addText(richText(deck.title || "", theme, { ...T.h1(theme.title), fontFace: fontFor(deck.title) }), {
    x: 0.9, y: 2.42, w: 8.6, h: 2.35,
    ...T.h1(theme.title), fontFace: fontFor(deck.title),
  });
  if (deck.subtitle) {
    s.addText(richText(deck.subtitle, theme, { ...T.h3(theme.title), fontFace: fontFor(deck.subtitle) }), {
      x: 0.9, y: 4.9, w: 8.6, h: 1.2,
      ...T.h3(theme.title), fontFace: fontFor(deck.subtitle),
    });
  }
  const dateStr = new Date().toLocaleDateString(lang === "en" ? "en-US" : "my-MM", {
    year: "numeric", month: "long", day: "numeric",
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 6.42, w: 5.6, h: 0.025, fill: { color: theme.band }, line: { color: theme.band },
  });
  s.addText(dateStr, {
    x: 0.9, y: 6.55, w: 6, h: 0.4,
    ...T.caption(), fontFace: "Calibri",
  });
}

function addBulletsSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  glassCard(s, pptx, theme, 1.3, 1.65, 11.1, 5.1, 0.16);
  watermark(s, theme, item.icon);
  const hasTakeaway = !!(item.takeaway && item.takeaway.trim());
  const nBullets = (item.bullets || []).length;
  // Type scale is 16-18pt: step down as density rises so 1.6 line-height always fits.
  const fs = hasTakeaway ? (nBullets > 4 ? 16 : 17) : (nBullets > 4 ? 17 : 18);
  const runs = bulletRuns(item.bullets, theme, fs);
  if (runs.length)
    s.addText(runs, { x: 1.75, y: 1.9, w: item.icon ? 8.35 : 9.85, h: hasTakeaway ? 4.0 : 4.6, valign: "top" });
  if (hasTakeaway) takeawayStrip(s, pptx, theme, item.takeaway.trim(), 1.6, 5.98, 10.5);
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addStatSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const hasStat = item.stat && item.stat.value;
  if (hasStat) {
    // stat card
    s.addShape(pptx.ShapeType.roundRect, {
      x: 1.5, y: 1.9, w: 4.6, h: 2.5, rectRadius: 0.18,
      fill: { color: theme.band }, line: { color: theme.band },
    });
    s.addShape(pptx.ShapeType.rect, {
      x: 1.5, y: 1.9, w: 0.12, h: 2.5, fill: { color: theme.accent }, line: { color: theme.accent },
    });
    s.addText(item.stat.value, {
      x: 1.85, y: 2.0, w: 4.0, h: 1.3,
      fontSize: 64, bold: true, color: theme.accent, fontFace: fontFor(item.stat.value),
    });
    s.addText(item.stat.label || "", {
      x: 1.85, y: 3.25, w: 4.0, h: 1.0,
      fontSize: 16, color: theme.text, fontFace: fontFor(item.stat.label), lineSpacingMultiple: 1.6,
    });
    const runs = bulletRuns(item.bullets, theme, 16);
    if (runs.length) s.addText(runs, { x: 6.7, y: 1.9, w: 5.1, h: 4.4, valign: "top" });
  } else {
    const runs = bulletRuns(item.bullets, theme, 18);
    if (runs.length) s.addText(runs, { x: 1.5, y: 1.85, w: 9.6, h: 4.6, valign: "top" });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

// ---- NEW: feature cards (3 glass cards in a row, like a modern SaaS pitch) ----
function addCardsSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const pts = item.points.length >= 2 ? item.points : (item.bullets || []).slice(0, 3).map((b) => {
    const t = btext(b);
    const m = t.split(/[:—–-]\s(.+)/);
    return { icon: bicon(b) || item.icon, title: m.length > 2 ? m[0].trim() : "", text: m.length > 2 ? m[1].trim() : t };
  });
  const n = Math.min(3, pts.length);
  const gap = 0.4, avail = 11.43, w = (avail - (n - 1) * gap) / n;
  pts.slice(0, n).forEach((p, k) => {
    const x = 0.95 + k * (w + gap), y = 2.0, h = 4.35;
    glassCard(s, pptx, theme, x, y, w, h, 0.16);
    // top accent bar
    s.addShape(pptx.ShapeType.rect, {
      x: x + 0.32, y: y + 0.26, w: 0.55, h: 0.06,
      fill: { color: theme.accent }, line: { color: theme.accent },
    });
    let ty = y + 0.5;
    if (p.icon) {
      // icon in a soft accent chip
      const cs = 0.8, cxp = x + w / 2;
      s.addShape(pptx.ShapeType.ellipse, {
        x: cxp - cs / 2, y: ty, w: cs, h: cs,
        fill: { color: theme.accent, transparency: 78 }, line: { color: theme.accent, transparency: 100 },
      });
      s.addText(p.icon, {
        x: cxp - cs / 2, y: ty + 0.03, w: cs, h: cs, align: "center",
        fontSize: 34, fontFace: "Segoe UI Emoji",
      });
      ty += cs + 0.2;
    }
    if (p.title) {
      const tBase = { ...T.h3(theme.title), fontFace: fontFor(p.title) };
      s.addText(richText(p.title, theme, tBase), {
        x: x + 0.3, y: ty, w: w - 0.6, h: 1.1, align: "center",
        ...tBase, valign: "top",
      });
      ty += 1.15;
    }
    if (p.text) {
      const bBase = { fontSize: 16, color: theme.text, fontFace: fontFor(p.text), lineSpacingMultiple: 1.6 };
      s.addText(richText(p.text, theme, bBase), {
        x: x + 0.3, y: ty, w: w - 0.6, h: h - (ty - y) - 0.3, align: "center",
        ...bBase, valign: "top",
      });
    }
  });
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

// ---- NEW: stats grid (2-4 big numbers on glass cards) ----
function addStatsSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const stats = item.stats.slice(0, 4);
  const n = stats.length;
  const cols = n <= 3 ? n : 2;
  const rows = Math.ceil(n / cols);
  const gapX = 0.4, gapY = 0.4, availW = 11.43, availH = 4.5;
  const cw = (availW - (cols - 1) * gapX) / cols;
  const ch = (availH - (rows - 1) * gapY) / rows;
  stats.forEach((st, k) => {
    const col = k % cols, row = Math.floor(k / cols);
    const x = 0.95 + col * (cw + gapX), y = 1.95 + row * (ch + gapY);
    glassCard(s, pptx, theme, x, y, cw, ch, 0.16);
    s.addShape(pptx.ShapeType.rect, {
      x: x + 0.3, y: y + 0.32, w: 0.5, h: 0.06,
      fill: { color: theme.accent }, line: { color: theme.accent },
    });
    s.addText(st.value, {
      x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.7,
      fontSize: Math.min(40, ch * 20), bold: true, color: theme.accent, fontFace: fontFor(st.value),
    });
    s.addText(st.label || "", {
      x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: Math.max(0.85, ch - 1.2), valign: "top",
      fontSize: 16, color: theme.text, fontFace: fontFor(st.label), lineSpacingMultiple: 1.6,
    });
  });
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addTwoColSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const bullets = item.bullets || [];
  const mid = Math.ceil(bullets.length / 2);
  const left = bullets.slice(0, mid);
  const right = bullets.slice(mid);
  const hasTakeaway = !!(item.takeaway && item.takeaway.trim());
  const panelH = hasTakeaway ? 3.9 : 4.6;
  glassCard(s, pptx, theme, 1.3, 1.8, 5.25, panelH, 0.16);
  glassCard(s, pptx, theme, 6.78, 1.8, 5.25, panelH, 0.16);
  const lr = bulletRuns(left, theme, 16);
  const rr = bulletRuns(right, theme, 16);
  const textH = hasTakeaway ? 3.35 : 4.1;
  if (lr.length) s.addText(lr, { x: 1.65, y: 2.05, w: 4.55, h: textH, valign: "top" });
  if (rr.length) s.addText(rr, { x: 7.13, y: 2.05, w: 4.55, h: textH, valign: "top" });
  if (hasTakeaway) takeawayStrip(s, pptx, theme, item.takeaway.trim(), 1.3, 5.85, 11.1);
  if (item.icon) {
    s.addText(item.icon, {
      x: 11.9, y: 6.2, w: 0.9, h: 0.9, align: "center", fontSize: 36, fontFace: "Segoe UI Emoji",
    });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addQuoteSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  s.addShape(pptx.ShapeType.ellipse, {
    x: -1.8, y: -1.8, w: 5.4, h: 5.4,
    fill: { color: theme.band, transparency: 55 }, line: { color: theme.band, transparency: 100 },
  });
  numberPill(s, pptx, theme, idx + 1);
  const quote = item.quote || (item.bullets || []).map(btext).join(" ") || item.heading;
  glassCard(s, pptx, theme, 1.6, 2.15, 10.13, 3.35, 0.2);
  s.addText("\u201C", {
    x: 1.5, y: 1.15, w: 10.3, h: 1.2, align: "center",
    fontSize: 110, bold: true, color: theme.accent, fontFace: "Calibri",
  });
  s.addText(richText(quote, theme, { ...T.quote(theme.title), fontFace: fontFor(quote) }), {
    x: 2.2, y: 2.75, w: 8.9, h: 2.1, align: "center",
    ...T.quote(theme.title), fontFace: fontFor(quote),
  });
  if (item.quoteBy) {
    s.addShape(pptx.ShapeType.rect, {
      x: 6.16, y: 5.5, w: 1.0, h: 0.06, fill: { color: theme.accent }, line: { color: theme.accent },
    });
    s.addText("— " + item.quoteBy, {
      x: 2.2, y: 5.7, w: 8.9, h: 0.6, align: "center",
      ...T.caption(), fontFace: fontFor(item.quoteBy),
    });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, item.heading, `${idx + 1} / ${deck.slides.length}`);
}

// ---- NEW: timeline — horizontal process with numbered nodes + connector ----
function addTimelineSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const steps = (item.points || []).slice(0, 4);
  const n = steps.length;
  const gap = 0.5, x0 = 0.95, avail = 11.43;
  const w = (avail - (n - 1) * gap) / n;
  const nodeY = 2.55, nodeD = 0.66;
  // connector line behind the nodes
  s.addShape(pptx.ShapeType.rect, {
    x: x0 + w / 2, y: nodeY + nodeD / 2 - 0.02,
    w: (x0 + avail - w / 2) - (x0 + w / 2), h: 0.04,
    fill: { color: theme.accent, transparency: 55 }, line: { color: theme.accent, transparency: 100 },
  });
  steps.forEach((st, k) => {
    const x = x0 + k * (w + gap), cx = x + w / 2;
    s.addShape(pptx.ShapeType.ellipse, {
      x: cx - nodeD / 2, y: nodeY, w: nodeD, h: nodeD,
      fill: { color: theme.accent }, line: { color: theme.accent },
    });
    s.addText(String(k + 1).padStart(2, "0"), {
      x: cx - nodeD / 2, y: nodeY, w: nodeD, h: nodeD, align: "center",
      fontSize: 19, bold: true, color: isLight(theme) ? "FFFFFF" : theme.bg, fontFace: "Calibri",
    });
    if (st.title) {
      const tBase = { fontSize: 19, bold: true, color: theme.title, charSpacing: tight(19), lineSpacingMultiple: 1.15, fontFace: fontFor(st.title) };
      s.addText(richText(st.title, theme, tBase), {
        x, y: nodeY + nodeD + 0.28, w, h: 0.85, align: "center",
        ...tBase, valign: "top",
      });
    }
    if (st.text) {
      const bBase = { fontSize: 15, color: theme.text, lineSpacingMultiple: 1.5, fontFace: fontFor(st.text) };
      s.addText(richText(st.text, theme, bBase), {
        x, y: nodeY + nodeD + 1.2, w, h: 1.65, align: "center",
        ...bBase, valign: "top",
      });
    }
  });
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

// ---- NEW: split — giant icon art panel left, bullets right ----
function addSplitSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const px = 1.3, py = 1.95, pw = 4.2, ph = 3.9;
  glassCard(s, pptx, theme, px, py, pw, ph, 0.18);
  // spotlight rings behind the icon
  s.addShape(pptx.ShapeType.ellipse, {
    x: px + pw / 2 - 1.2, y: py + 0.5, w: 2.4, h: 2.4,
    fill: { color: theme.bg, transparency: 100 }, line: { color: theme.accent, width: 2.5, transparency: 35 },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: px + pw / 2 - 0.88, y: py + 0.82, w: 1.76, h: 1.76,
    fill: { color: theme.accent, transparency: 80 }, line: { color: theme.accent, transparency: 100 },
  });
  if (item.icon) {
    s.addText(item.icon, {
      x: px, y: py + 0.62, w: pw, h: 2.16, align: "center",
      fontSize: 110, fontFace: "Segoe UI Emoji",
    });
  }
  const runs = bulletRuns(item.bullets, theme, 17);
  if (runs.length) s.addText(runs, { x: 6.1, y: 1.95, w: 6.0, h: 3.7, valign: "middle" });
  const tw = item.takeaway && item.takeaway.trim();
  if (tw) takeawayStrip(s, pptx, theme, tw, 1.3, 5.98, 10.73);
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

// ---- NEW: hero — one massive statement/number as the whole message ----
function addHeroSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  ambient(s, pptx, theme);
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  // giant accent ring as the backdrop
  s.addShape(pptx.ShapeType.ellipse, {
    x: 5.16, y: 1.3, w: 3.0, h: 3.0,
    fill: { color: theme.bgDeep, transparency: 100 }, line: { color: theme.accent, width: 3, transparency: 25 },
  });
  const hero = item.hero || item.heading;
  const base = {
    fontSize: 60, bold: true, color: theme.title,
    charSpacing: tight(60), lineSpacingMultiple: 1.05, fontFace: fontFor(hero),
  };
  s.addText(richText(hero, theme, base), {
    x: 1.6, y: 2.0, w: 10.1, h: 2.6, align: "center",
    ...base, valign: "middle",
  });
  if (item.sub) {
    const subBase = { fontSize: 18, italic: true, color: theme.text, lineSpacingMultiple: 1.5, fontFace: fontFor(item.sub) };
    s.addText(richText(item.sub, theme, subBase), {
      x: 2.6, y: 4.85, w: 8.1, h: 1.2, align: "center",
      ...subBase, valign: "top",
    });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, pptx, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addClosingSlide(pptx, theme, deck, lang) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  s.addShape(pptx.ShapeType.ellipse, {
    x: 3.06, y: 1.2, w: 7.2, h: 7.2,
    fill: { color: theme.band, transparency: 40 }, line: { color: theme.band, transparency: 100 },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: 5.46, y: 2.1, w: 2.4, h: 2.4,
    fill: { color: theme.bgDeep, transparency: 100 }, line: { color: theme.accent, width: 2.5 },
  });
  const thanks = lang === "en" ? "Thank You" : "ကျေးဇူးတင်ပါတယ်";
  s.addText(thanks, {
    x: 1.5, y: 3.0, w: 10.33, h: 1.3, align: "center",
    ...T.h1(theme.title), fontFace: fontFor(thanks),
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 6.16, y: 4.45, w: 1.0, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  const ctBase = { fontSize: 18, color: theme.text, fontFace: fontFor(deck.title), lineSpacingMultiple: 1.6 };
  s.addText(richText(deck.title || "", theme, ctBase), {
    x: 1.5, y: 4.7, w: 10.33, h: 1.0, align: "center",
    ...ctBase,
  });
  const q = lang === "en" ? "Questions & Discussion" : "မေးခွန်းများ နှင့် ဆွေးနွေးခန်း";
  s.addText(q, {
    x: 1.5, y: 5.8, w: 10.33, h: 0.6, align: "center",
    fontSize: 16, italic: true, color: theme.accentSoft, fontFace: fontFor(q),
  });
}

// ---- slide transitions (post-process the pptx zip) --------------------------

const TRANSITION_XML = {
  fade: '<p:transition spd="med" advOnClk="1"><p:fade thruBlk="1"/></p:transition>',
  push: '<p:transition spd="med" advOnClk="1"><p:push thruBlk="1" dir="l"/></p:transition>',
  wipe: '<p:transition spd="med" advOnClk="1"><p:wipe thruBlk="1" dir="l"/></p:transition>',
  cover: '<p:transition spd="med" advOnClk="1"><p:cover thruBlk="1" dir="d"/></p:transition>',
};

// Transition sequences per topic mood, chosen by the AI design director.
const MOOD_TRANSITIONS = {
  energetic: ["push", "wipe", "push", "cover"],
  elegant: ["fade", "fade", "cover", "fade"],
  bold: ["cover", "wipe", "push", "cover"],
  calm: ["fade", "wipe", "fade", "fade"],
};
const DEFAULT_TRANSITIONS = ["fade", "push", "wipe", "cover"];

function applyTransitions(pptxBuffer, mood) {
  const seq = MOOD_TRANSITIONS[mood] || DEFAULT_TRANSITIONS;
  const zip = new AdmZip(Buffer.from(pptxBuffer));
  const entries = zip.getEntries().filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName));
  entries.forEach((entry, i) => {
    let xml = entry.getData().toString("utf8");
    if (xml.includes("<p:transition")) return;
    const trans = TRANSITION_XML[seq[i % seq.length]];
    if (!xml.includes("</p:cSld>")) return;
    xml = xml.replace("</p:cSld>", "</p:cSld>" + trans);
    zip.updateFile(entry.entryName, Buffer.from(xml, "utf8"));
  });
  return zip.toBuffer();
}

// ---- main -------------------------------------------------------------------

// Pick the effective layout, falling back gracefully when data is missing.
function resolveLayout(item) {
  const L = item.layout;
  if ((L === "stats" || L === "stat") && item.stats.length >= 2) return "stats";
  if ((L === "stat" || L === "stats") && (item.stats.length === 1 || (item.stat && item.stat.value))) return "stat";
  if (L === "timeline" && item.points.length >= 2) return "timeline";
  if (L === "split" && item.bullets.length >= 2 && item.icon) return "split";
  if (L === "hero" && (item.hero || item.heading)) return "hero";
  if (L === "cards" && (item.points.length >= 2 || item.bullets.length >= 2)) return "cards";
  if (L === "quote" && (item.quote || item.bullets.length)) return "quote";
  if (L === "two-col" && item.bullets.length > 2) return "two-col";
  return "bullets";
}

async function buildPptx(deck, themeKey, lang = "my", mood) {
  const theme = THEMES[themeKey] ? THEMES[themeKey] : THEMES["navy-gold"];
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE169", width: 13.33, height: 7.5 });
  pptx.layout = "WIDE169";
  pptx.author = "PPT Slide Maker";
  pptx.title = deck.title || "Slide Deck";
  pptx.subject = deck.subtitle || "";

  addTitleSlide(pptx, theme, deck, lang);
  deck.slides.forEach((item, i) => {
    const layout = resolveLayout(item);
    if (layout === "stats") addStatsSlide(pptx, theme, deck, i, item);
    else if (layout === "stat") addStatSlide(pptx, theme, deck, i, item);
    else if (layout === "timeline") addTimelineSlide(pptx, theme, deck, i, item);
    else if (layout === "split") addSplitSlide(pptx, theme, deck, i, item);
    else if (layout === "hero") addHeroSlide(pptx, theme, deck, i, item);
    else if (layout === "cards") addCardsSlide(pptx, theme, deck, i, item);
    else if (layout === "quote") addQuoteSlide(pptx, theme, deck, i, item);
    else if (layout === "two-col") addTwoColSlide(pptx, theme, deck, i, item);
    else addBulletsSlide(pptx, theme, deck, i, item);
  });
  addClosingSlide(pptx, theme, deck, lang);

  const raw = await pptx.write({ outputType: "nodebuffer" });
  return applyTransitions(raw, mood);
}

module.exports = { buildPptx, THEME_KEYS };
