"use strict";
const express = require("express");
const path = require("path");
const store = require("./store");
const gw = require("./gateway");
const { buildPptx, THEME_KEYS } = require("./pptx");
const { THEMES } = require("./themes");

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "..", "public")));

function creds() {
  return {
    baseUrl: store.getSetting("gateway_base_url") || gw.DEFAULT_BASE_URL,
    apiKey: store.getSecret("gateway_api_key"),
    model: store.getSetting("gateway_model") || "claude-sonnet-5",
  };
}

app.get("/api/status", (req, res) => {
  const c = creds();
  res.json({
    gatewayConfigured: Boolean(c.apiKey),
    ephemeralKey: store.isEphemeralKey(),
    baseUrl: c.baseUrl,
    model: c.model,
    themes: THEME_KEYS.map((k) => ({ key: k, name: THEMES[k].name, nameMy: THEMES[k].nameMy, colors: THEMES[k] })),
    slideCounts: [5, 8, 10, 12],
  });
});

app.get("/api/settings", (req, res) => {
  const c = creds();
  res.json({ baseUrl: c.baseUrl, model: c.model, hasKey: Boolean(c.apiKey) });
});

app.post("/api/settings", (req, res) => {
  const { baseUrl, apiKey, model } = req.body || {};
  if (baseUrl) store.setSetting("gateway_base_url", String(baseUrl).trim());
  if (apiKey) store.setSecret("gateway_api_key", String(apiKey));
  if (model) store.setSetting("gateway_model", String(model));
  res.json({ ok: true, hasKey: Boolean(creds().apiKey) });
});

app.post("/api/settings/test", async (req, res) => {
  try {
    const c = creds();
    if (!c.apiKey) return res.status(400).json({ ok: false, error: "no-key" });
    const models = await gw.listModels(c.baseUrl, c.apiKey);
    res.json({ ok: true, count: models.length });
  } catch (e) {
    res.status(502).json({ ok: false, error: String(e.message || e).slice(0, 300) });
  }
});

app.get("/api/models", async (req, res) => {
  try {
    const c = creds();
    if (!c.apiKey) return res.status(400).json({ error: "no-key" });
    res.json({ models: await gw.listModels(c.baseUrl, c.apiKey) });
  } catch (e) {
    res.status(502).json({ error: String(e.message || e).slice(0, 300) });
  }
});

app.post("/api/generate", async (req, res) => {
  try {
    const { topic, detail, slides, theme, lang, model } = req.body || {};
    const cleanTopic = String(topic || "").trim();
    if (!cleanTopic) return res.status(400).json({ error: "empty-topic" });
    const count = [5, 8, 10, 12].includes(Number(slides)) ? Number(slides) : 8;
    const themeKey = THEME_KEYS.includes(theme) ? theme : "navy-gold";
    const useLang = lang === "en" ? "en" : "my";
    const c = creds();
    if (!c.apiKey) return res.status(400).json({ error: "no-key" });
    const useModel = model || c.model;

    const raw = await gw.chatCompletion(
      c.baseUrl,
      c.apiKey,
      useModel,
      gw.deckSystemPrompt(cleanTopic, detail, count, useLang),
      `Topic: ${cleanTopic}`,
      { maxTokens: 8000 }
    );
    const parsed = gw.extractDeckJson(raw);
    if (!parsed) return res.status(502).json({ error: "no-json", detail: raw.slice(0, 300) });
    const deck = gw.normalizeDeck(parsed, count);
    if (!deck.title) return res.status(502).json({ error: "bad-deck" });

    const id = store.saveDeck({
      title: deck.title, topic: cleanTopic, theme: themeKey, lang: useLang, model: useModel, deck,
    });
    res.json({ ok: true, id, deck, theme: themeKey, lang: useLang });
  } catch (e) {
    res.status(502).json({ error: String(e.message || e).slice(0, 300) });
  }
});

app.get("/api/decks", (req, res) => {
  res.json({ items: store.listDecks() });
});

function safeFilename(title) {
  const t = String(title || "slides").replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 60) || "slides";
  return t + ".pptx";
}

app.get("/api/decks/:id/download", async (req, res) => {
  try {
    const row = store.getDeck(Number(req.params.id));
    if (!row) return res.status(404).json({ error: "not-found" });
    const buf = await buildPptx(row.deck, row.theme, row.lang);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(safeFilename(row.title))}`);
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(500).json({ error: String(e.message || e).slice(0, 300) });
  }
});

app.delete("/api/decks/:id", (req, res) => {
  const ok = store.deleteDeck(Number(req.params.id));
  res.json({ ok });
});

const PORT = process.env.PORT || 3333;
app.listen(PORT, () => console.log(`[ppt-slide-maker] listening on :${PORT}`));
