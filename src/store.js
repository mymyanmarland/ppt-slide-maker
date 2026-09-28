"use strict";
// SQLite storage + AES-256-GCM encrypted settings (gateway key never stored in plain text).
const { DatabaseSync } = require("node:sqlite");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const dataDir = process.env.DATA_DIR || path.join(__dirname, "..", "data");
fs.mkdirSync(dataDir, { recursive: true });
const db = new DatabaseSync(path.join(dataDir, "app.db"));

db.exec(`
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS decks (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  topic TEXT NOT NULL DEFAULT '',
  theme TEXT NOT NULL DEFAULT '',
  lang TEXT NOT NULL DEFAULT 'my',
  model TEXT NOT NULL DEFAULT '',
  deck_json TEXT NOT NULL,
  created_at INTEGER NOT NULL
);`);

// ---- encryption (APP_SECRET env, per-boot ephemeral fallback with warning) ----
let _key = null;
let _ephemeral = false;
function appKey() {
  if (_key) return _key;
  const s = process.env.APP_SECRET;
  if (s && s.length >= 16) {
    _key = crypto.createHash("sha256").update(String(s)).digest();
  } else {
    _key = crypto.randomBytes(32);
    _ephemeral = true;
    console.warn(
      "[store] APP_SECRET not set — using an ephemeral key. The saved gateway key will NOT survive a restart."
    );
  }
  return _key;
}
function enc(plain) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv("aes-256-gcm", appKey(), iv);
  const ct = Buffer.concat([c.update(String(plain), "utf8"), c.final()]);
  return JSON.stringify({
    iv: iv.toString("base64"),
    tag: c.getAuthTag().toString("base64"),
    data: ct.toString("base64"),
  });
}
function dec(payload) {
  const o = JSON.parse(payload);
  const d = crypto.createDecipheriv(
    "aes-256-gcm",
    appKey(),
    Buffer.from(o.iv, "base64")
  );
  d.setAuthTag(Buffer.from(o.tag, "base64"));
  return Buffer.concat([
    d.update(Buffer.from(o.data, "base64")),
    d.final(),
  ]).toString("utf8");
}

function getSetting(key) {
  const row = db.prepare("SELECT value FROM settings WHERE key = ?").get(key);
  return row ? row.value : null;
}
function setSetting(key, value) {
  db.prepare(
    "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value"
  ).run(key, value);
}

const SECRET_PREFIX = "enc:";
function getSecret(name) {
  const v = getSetting("secret:" + name);
  if (!v) return null;
  try {
    return v.startsWith(SECRET_PREFIX) ? dec(v.slice(SECRET_PREFIX.length)) : v;
  } catch {
    return null;
  }
}
function setSecret(name, plain) {
  setSetting("secret:" + name, SECRET_PREFIX + enc(plain));
}

function saveDeck({ title, topic, theme, lang, model, deck }) {
  const r = db
    .prepare(
      "INSERT INTO decks (title, topic, theme, lang, model, deck_json, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)"
    )
    .run(title, topic || "", theme || "", lang || "my", model || "", JSON.stringify(deck), Date.now());
  return Number(r.lastInsertRowid);
}
function listDecks(limit = 60) {
  return db
    .prepare(
      "SELECT id, title, topic, theme, lang, model, created_at FROM decks ORDER BY id DESC LIMIT ?"
    )
    .all(limit);
}
function getDeck(id) {
  const row = db.prepare("SELECT * FROM decks WHERE id = ?").get(id);
  if (!row) return null;
  try {
    row.deck = JSON.parse(row.deck_json);
  } catch {
    return null;
  }
  return row;
}
function deleteDeck(id) {
  return db.prepare("DELETE FROM decks WHERE id = ?").run(id).changes > 0;
}

module.exports = {
  getSetting,
  setSetting,
  getSecret,
  setSecret,
  saveDeck,
  listDecks,
  getDeck,
  deleteDeck,
  isEphemeralKey: () => _ephemeral,
};
