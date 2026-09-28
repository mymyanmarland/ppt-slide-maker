"use strict";
// Deck JSON -> real .pptx bytes via pptxgenjs, with rich layouts + slide transitions.
const PptxGenJS = require("pptxgenjs");
const AdmZip = require("adm-zip");
const { THEMES, THEME_KEYS } = require("./themes");

const MY_RE = /[\u1000-\u109F]/;
const fontFor = (s) => (MY_RE.test(String(s || "")) ? "Noto Sans Myanmar" : "Calibri");
const isLight = (theme) => theme.bgDeep === "F4F5F7";

// ---- shared bits -----------------------------------------------------------

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
  slide.addText(heading || "", {
    x: 1.5, y: 0.32, w: 11.3, h: 1.1,
    fontSize: 30, bold: true, color: theme.title, fontFace: fontFor(heading), valign: "middle",
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
  return (bullets || []).map((b, i, arr) => ({
    text: b,
    options: {
      fontSize,
      color: theme.text,
      fontFace: fontFor(b),
      bullet: { code: "2022", color: theme.accent, indent: 20 },
      paraSpaceAfter: 14,
      lineSpacingMultiple: 1.25,
      breakLine: i < arr.length - 1,
    },
  }));
}

function addFooter(slide, theme, left, right) {
  slide.addText(String(left || ""), {
    x: 0.5, y: 7.02, w: 8, h: 0.3,
    fontSize: 9, color: theme.footer, fontFace: fontFor(left),
  });
  slide.addText(String(right || ""), {
    x: 11.5, y: 7.02, w: 1.33, h: 0.3, align: "right",
    fontSize: 9, color: theme.footer, fontFace: "Calibri",
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
  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 2.2, w: 1.4, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  s.addText(deck.title || "", {
    x: 0.9, y: 2.45, w: 8.6, h: 2.3,
    fontSize: 44, bold: true, color: theme.title, fontFace: fontFor(deck.title), lineSpacingMultiple: 1.05,
  });
  if (deck.subtitle) {
    s.addText(deck.subtitle, {
      x: 0.9, y: 4.9, w: 8.6, h: 1.2,
      fontSize: 18, color: theme.muted, fontFace: fontFor(deck.subtitle), lineSpacingMultiple: 1.15,
    });
  }
  const dateStr = new Date().toLocaleDateString(lang === "en" ? "en-US" : "my-MM", {
    year: "numeric", month: "long", day: "numeric",
  });
  s.addText(dateStr, {
    x: 0.9, y: 6.55, w: 6, h: 0.4,
    fontSize: 11, color: theme.footer, fontFace: "Calibri",
  });
}

function addBulletsSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  watermark(s, theme, item.icon);
  const runs = bulletRuns(item.bullets, theme, 18);
  if (runs.length) s.addText(runs, { x: 1.5, y: 1.85, w: item.icon ? 8.7 : 10.2, h: 4.6, valign: "top" });
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addStatSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
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
      fontSize: 15, color: theme.muted, fontFace: fontFor(item.stat.label), lineSpacingMultiple: 1.15,
    });
    const runs = bulletRuns(item.bullets, theme, 16);
    if (runs.length) s.addText(runs, { x: 6.7, y: 1.9, w: 5.1, h: 4.4, valign: "top" });
  } else {
    const runs = bulletRuns(item.bullets, theme, 18);
    if (runs.length) s.addText(runs, { x: 1.5, y: 1.85, w: 9.6, h: 4.6, valign: "top" });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addTwoColSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  topBar(s, pptx, theme);
  numberPill(s, pptx, theme, idx + 1);
  headingBlock(s, pptx, theme, item.heading);
  const bullets = item.bullets || [];
  const mid = Math.ceil(bullets.length / 2);
  const left = bullets.slice(0, mid);
  const right = bullets.slice(mid);
  // vertical divider
  s.addShape(pptx.ShapeType.rect, {
    x: 6.6, y: 1.9, w: 0.04, h: 4.4, fill: { color: theme.band }, line: { color: theme.band },
  });
  const lr = bulletRuns(left, theme, 16);
  const rr = bulletRuns(right, theme, 16);
  if (lr.length) s.addText(lr, { x: 1.5, y: 1.9, w: 4.7, h: 4.4, valign: "top" });
  if (rr.length) s.addText(rr, { x: 7.0, y: 1.9, w: 4.7, h: 4.4, valign: "top" });
  if (item.icon) {
    s.addText(item.icon, {
      x: 11.9, y: 6.2, w: 0.9, h: 0.9, align: "center", fontSize: 36, fontFace: "Segoe UI Emoji",
    });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addQuoteSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  topBar(s, pptx, theme);
  s.addShape(pptx.ShapeType.ellipse, {
    x: -1.8, y: -1.8, w: 5.4, h: 5.4,
    fill: { color: theme.band, transparency: 55 }, line: { color: theme.band, transparency: 100 },
  });
  numberPill(s, pptx, theme, idx + 1);
  s.addText("\u201C", {
    x: 1.5, y: 1.6, w: 10.3, h: 1.2, align: "center",
    fontSize: 110, bold: true, color: theme.accent, fontFace: "Calibri",
  });
  const quote = item.quote || item.bullets.join(" ") || item.heading;
  s.addText(quote, {
    x: 2.2, y: 2.9, w: 8.9, h: 2.4, align: "center",
    fontSize: 26, italic: true, color: theme.title, fontFace: fontFor(quote), lineSpacingMultiple: 1.3,
  });
  if (item.quoteBy) {
    s.addShape(pptx.ShapeType.rect, {
      x: 6.16, y: 5.5, w: 1.0, h: 0.06, fill: { color: theme.accent }, line: { color: theme.accent },
    });
    s.addText("— " + item.quoteBy, {
      x: 2.2, y: 5.7, w: 8.9, h: 0.6, align: "center",
      fontSize: 15, color: theme.muted, fontFace: fontFor(item.quoteBy),
    });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, theme, item.heading, `${idx + 1} / ${deck.slides.length}`);
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
    fontSize: 52, bold: true, color: theme.title, fontFace: fontFor(thanks),
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 6.16, y: 4.45, w: 1.0, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  s.addText(deck.title || "", {
    x: 1.5, y: 4.7, w: 10.33, h: 1.0, align: "center",
    fontSize: 20, color: theme.muted, fontFace: fontFor(deck.title),
  });
  const q = lang === "en" ? "Questions & Discussion" : "မေးခွန်းများ နှင့် ဆွေးနွေးခန်း";
  s.addText(q, {
    x: 1.5, y: 5.8, w: 10.33, h: 0.6, align: "center",
    fontSize: 15, italic: true, color: theme.accentSoft, fontFace: fontFor(q),
  });
}

// ---- slide transitions (post-process the pptx zip) --------------------------

const TRANSITIONS = [
  '<p:transition spd="med" advOnClk="1"><p:fade thruBlk="1"/></p:transition>',
  '<p:transition spd="med" advOnClk="1"><p:push thruBlk="1" dir="l"/></p:transition>',
  '<p:transition spd="med" advOnClk="1"><p:wipe thruBlk="1" dir="l"/></p:transition>',
  '<p:transition spd="med" advOnClk="1"><p:cover thruBlk="1" dir="d"/></p:transition>',
];

function applyTransitions(pptxBuffer) {
  const zip = new AdmZip(Buffer.from(pptxBuffer));
  const entries = zip.getEntries().filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName));
  entries.forEach((entry, i) => {
    let xml = entry.getData().toString("utf8");
    if (xml.includes("<p:transition")) return;
    const trans = TRANSITIONS[i % TRANSITIONS.length];
    if (!xml.includes("</p:cSld>")) return;
    xml = xml.replace("</p:cSld>", "</p:cSld>" + trans);
    zip.updateFile(entry.entryName, Buffer.from(xml, "utf8"));
  });
  return zip.toBuffer();
}

// ---- main -------------------------------------------------------------------

async function buildPptx(deck, themeKey, lang = "my") {
  const theme = THEMES[themeKey] ? THEMES[themeKey] : THEMES["navy-gold"];
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE169", width: 13.33, height: 7.5 });
  pptx.layout = "WIDE169";
  pptx.author = "PPT Slide Maker";
  pptx.title = deck.title || "Slide Deck";
  pptx.subject = deck.subtitle || "";

  addTitleSlide(pptx, theme, deck, lang);
  deck.slides.forEach((item, i) => {
    const layout = item.layout === "stat" && item.stat && item.stat.value
      ? "stat"
      : item.layout === "quote" && (item.quote || item.bullets.length)
        ? "quote"
        : item.layout === "two-col" && item.bullets.length > 2
          ? "two-col"
          : "bullets";
    if (layout === "stat") addStatSlide(pptx, theme, deck, i, item);
    else if (layout === "quote") addQuoteSlide(pptx, theme, deck, i, item);
    else if (layout === "two-col") addTwoColSlide(pptx, theme, deck, i, item);
    else addBulletsSlide(pptx, theme, deck, i, item);
  });
  addClosingSlide(pptx, theme, deck, lang);

  const raw = await pptx.write({ outputType: "nodebuffer" });
  return applyTransitions(raw);
}

module.exports = { buildPptx, THEME_KEYS };
