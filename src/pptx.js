"use strict";
// Deck JSON -> real .pptx bytes via pptxgenjs.
const PptxGenJS = require("pptxgenjs");
const { THEMES, THEME_KEYS } = require("./themes");

const MY_RE = /[\u1000-\u109F]/;
const fontFor = (s) => (MY_RE.test(String(s || "")) ? "Noto Sans Myanmar" : "Calibri");

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

function addTitleSlide(pptx, theme, deck, lang) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  // decorative glow circle, partially off-canvas
  s.addShape(pptx.ShapeType.ellipse, {
    x: 9.4, y: -2.6, w: 7.2, h: 7.2, fill: { color: theme.band }, line: { color: theme.band },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: -2.4, y: 5.4, w: 5.2, h: 5.2, fill: { color: theme.band }, line: { color: theme.band },
  });
  // accent rule
  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 2.35, w: 1.4, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  const kicker = lang === "en" ? "AI SLIDE DECK" : "AI ဆလိုက်ဒ်";
  s.addText(kicker, {
    x: 0.9, y: 1.85, w: 6, h: 0.4,
    fontSize: 13, bold: true, color: theme.accent, charSpacing: 6, fontFace: fontFor(kicker),
  });
  s.addText(deck.title || "", {
    x: 0.9, y: 2.6, w: 8.6, h: 2.2,
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
  s.addNotes(`Title slide. ${deck.subtitle || ""}`.trim());
}

function addContentSlide(pptx, theme, deck, idx, item) {
  const s = pptx.addSlide();
  s.background = { color: theme.bg };
  // top accent bar
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 13.33, h: 0.09, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  // slide number pill
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.55, y: 0.42, w: 0.72, h: 0.42, rectRadius: 0.21,
    fill: { color: theme.accent }, line: { color: theme.accent },
  });
  s.addText(String(idx + 1).padStart(2, "0"), {
    x: 0.55, y: 0.42, w: 0.72, h: 0.42, align: "center",
    fontSize: 14, bold: true, color: theme.bgDeep === "F4F5F7" ? "FFFFFF" : theme.bg, fontFace: "Calibri",
  });
  const heading = item.heading || "";
  s.addText(heading, {
    x: 1.5, y: 0.32, w: 11.3, h: 1.1,
    fontSize: 30, bold: true, color: theme.title, fontFace: fontFor(heading), valign: "middle",
  });
  // underline accent
  s.addShape(pptx.ShapeType.rect, {
    x: 1.52, y: 1.32, w: 1.1, h: 0.05, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  if (item.bullets && item.bullets.length) {
    const runs = item.bullets.map((b, i) => ({
      text: b,
      options: {
        fontSize: 18,
        color: theme.text,
        fontFace: fontFor(b),
        bullet: { code: "2022", color: theme.accent, indent: 20 },
        paraSpaceAfter: 14,
        lineSpacingMultiple: 1.25,
        breakLine: i < item.bullets.length - 1,
      },
    }));
    s.addText(runs, { x: 1.5, y: 1.85, w: 10.4, h: 4.6, valign: "top" });
  }
  if (item.notes) s.addNotes(item.notes);
  addFooter(s, theme, deck.title, `${idx + 1} / ${deck.slides.length}`);
}

function addClosingSlide(pptx, theme, deck, lang) {
  const s = pptx.addSlide();
  s.background = { color: theme.bgDeep };
  s.addShape(pptx.ShapeType.ellipse, {
    x: 3.06, y: 1.4, w: 7.2, h: 7.2, fill: { color: theme.band }, line: { color: theme.band },
  });
  const thanks = lang === "en" ? "Thank You" : "ကျေးဇူးတင်ပါတယ်";
  s.addText(thanks, {
    x: 1.5, y: 2.7, w: 10.33, h: 1.3, align: "center",
    fontSize: 52, bold: true, color: theme.title, fontFace: fontFor(thanks),
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 6.16, y: 4.15, w: 1.0, h: 0.07, fill: { color: theme.accent }, line: { color: theme.accent },
  });
  s.addText(deck.title || "", {
    x: 1.5, y: 4.45, w: 10.33, h: 1.0, align: "center",
    fontSize: 20, color: theme.muted, fontFace: fontFor(deck.title),
  });
  const q = lang === "en" ? "Questions & Discussion" : "မေးခွန်းများ နှင့် ဆွေးနွေးခန်း";
  s.addText(q, {
    x: 1.5, y: 5.55, w: 10.33, h: 0.6, align: "center",
    fontSize: 15, italic: true, color: theme.accentSoft, fontFace: fontFor(q),
  });
}

async function buildPptx(deck, themeKey, lang = "my") {
  const theme = THEMES[themeKey] && THEMES[themeKey] ? THEMES[themeKey] : THEMES["navy-gold"];
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE169", width: 13.33, height: 7.5 });
  pptx.layout = "WIDE169";
  pptx.author = "PPT Slide Maker";
  pptx.title = deck.title || "Slide Deck";
  pptx.subject = deck.subtitle || "";

  addTitleSlide(pptx, theme, deck, lang);
  deck.slides.forEach((item, i) => addContentSlide(pptx, theme, deck, i, item));
  addClosingSlide(pptx, theme, deck, lang);

  return pptx.write({ outputType: "nodebuffer" });
}

module.exports = { buildPptx, THEME_KEYS };
