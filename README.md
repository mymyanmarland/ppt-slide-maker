# PPT Slide Maker — AI ဆလိုက်ဒ်ဖန်တီးစက်

AI-powered PowerPoint slide generator. Type a topic → the AI writes the full deck outline (title, headings, bullets, speaker notes) via your Claude N Codex gateway → the server builds a **real `.pptx` file** you can open in PowerPoint / Google Slides / Keynote.

## Features

- 🧠 **AI deck writing** — topic + optional details → complete outline in strict JSON
- 🎨 **6 design themes** — Navy Gold, Modern Minimal, Bold Impact, Corporate Blue, Myanmar Royal, Violet Tech
- 📊 **Real `.pptx` output** — title slide, numbered content slides with accent bullets, closing "Thank You" slide, speaker notes, footers
- 🇲🇲 **Bilingual UI** — Myanmar-first + English, and decks can be generated in either language
- 🖼️ **Live preview** — see every slide styled in the browser before downloading
- 💾 **Deck gallery** — saved decks can be re-downloaded or deleted anytime
- 🔐 **Secure** — gateway API key stored AES-256-GCM encrypted (server-side only, never reaches the browser)

## Quick start

```bash
npm install
APP_SECRET="a-long-random-secret" PORT=3333 npm start
```

Open http://localhost:3333 → **Settings** → paste your gateway URL + API key → Test connection → start generating.

## Tech

- **Stack:** Node.js + Express + vanilla JS (no build step)
- **PPTX:** [pptxgenjs](https://github.com/gitbrent/pptxgenjs) (16:9 widescreen)
- **Storage:** `node:sqlite` (built-in) — settings + saved decks
- **AI:** any OpenAI-compatible `/v1/chat/completions` gateway (default: Claude N Codex)

## API

| Method | Route | Description |
|--------|-------|-------------|
| GET | `/api/status` | config state, themes, slide counts |
| GET/POST | `/api/settings` | gateway URL / key / default model |
| POST | `/api/settings/test` | test gateway connection |
| GET | `/api/models` | list gateway models |
| POST | `/api/generate` | `{topic, detail, slides, theme, lang, model}` → deck JSON |
| GET | `/api/decks` | saved decks |
| GET | `/api/decks/:id/download` | download `.pptx` |
| DELETE | `/api/decks/:id` | delete a deck |

## Notes

- `APP_SECRET` env: encrypts the saved gateway key. Without it, an ephemeral key is used and the saved key is lost on restart (free-tier SQLite also resets on redeploy).
- Gateway used for testing: `https://claude-n-codex.com:8443/v1` (chat models).

MIT License.
