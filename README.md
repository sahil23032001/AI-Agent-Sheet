<div align="center">

# Groq Sheets Autocomplete

**An AI writing agent that lives inside Google Sheets and finishes your sentences.**

It reads the cell you're on, the column around it, and a brief you write about your product —
then suggests how to complete the entry. Powered by the Groq API's free tier, running entirely
inside your own Apps Script project.

[![License: MIT](https://img.shields.io/badge/License-MIT-0f6b3f.svg)](LICENSE)
[![Apps Script](https://img.shields.io/badge/Google_Apps_Script-V8-4285F4.svg)](https://developers.google.com/apps-script)
[![Powered by Groq](https://img.shields.io/badge/LLM-Groq_free_tier-f55036.svg)](https://console.groq.com)
[![Tests](https://img.shields.io/badge/tests-53_passing-0f6b3f.svg)](test/harness.js)
[![No server](https://img.shields.io/badge/backend-none-lightgrey.svg)](#privacy)

</div>

---

## Why

Spreadsheets are where a lot of writing actually happens — bug tickets, CRM notes, product copy,
research logs — and none of it gets AI help, because the assistants all live in documents and chat
windows. This puts one in the sidebar.

It is not a generic "ask an LLM" box. Every suggestion is built from the specific cell you're in:
the column header, the other fields in that row, the neighbouring entries that establish house
style, and a standing brief about what your product is. The result reads like the rest of your
column instead of like an AI.

```
┌─ AI Autocomplete ──────────────────┐
│ ● Tickets!C5      Fix · 8 examples │
│ ┌────────────────────────────────┐ │
│ │ The pantry view drops line     │ │
│ └────────────────────────────────┘ │
│ [       Suggest       ] □ Auto     │
│                                    │
│ CONTINUE C5                        │
│ ┌────────────────────────────────┐ │
│ │ The pantry view drops line    1│ │
│ │  items when an order is saved  │ │
│ │  with more than one item.      │ │
│ └────────────────────────────────┘ │
│ ┌────────────────────────────────┐ │
│ │ The pantry view drops line    2│ │
│ │  items on save; only the first │ │
│ │  item persists.                │ │
│ └────────────────────────────────┘ │
│                                    │
│ SELECTION                          │
│ C5:C40 · 36 rows × 1 col           │
│ 36 cells · 22 empty · 14 filled    │
│ ☑ Skip cells with content          │
│ [ Complete 22 cells ]  [ Stop ]    │
│ ███████████░░░░░░░░░░░  11/22      │
└────────────────────────────────────┘
```

---

## Contents

- [Features](#features)
- [Install](#install)
- [Using it](#using-it)
- [Configuration](#configuration)
- [Presets](#presets)
- [How it works](#how-it-works)
- [Privacy](#privacy)
- [Tests](#tests)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Contributing](#contributing)
- [License](#license)

---

## Features

### Sentence completion, in context

Select a cell, half-type a sentence, press **Suggest** (or `Ctrl`/`Cmd`+`Enter`). You get up to 5
alternatives; click one or press `1`–`5` to write it in. Two modes, chosen automatically:

| Cell state | Mode | Behaviour |
|---|---|---|
| Partly typed | `continue` | Returns only the continuation. Any echo of your prefix is stripped server-side. |
| Empty | `fill` | Returns the whole value. |

### Context the model actually gets

Each toggle is independent, so you can trade quality against tokens:

| Source | What it does |
|---|---|
| **Project context** | Your brief about the product and sheet. Sent first, as background fact. |
| **Workbook structure** | Every tab name and its column headers. ~20–60 tokens. |
| **Column header** | Labels what this cell is for. |
| **Row across columns** | Other filled fields in the row, as `Header: value` facts. |
| **Neighbouring rows** | ±N entries from the same column. The single biggest quality lever — this is what teaches tone, length and format. |
| **Tone instruction** | Your style rules, injected at highest priority. |

### Draft your context automatically

**Draft from sheet** reads your workbook structure and six real rows, then writes a first version
of the brief. Its prompt forbids inventing company names, teams or purposes not evident in the
data. Treat it as a starting point and correct it.

### Multi-cell runs

Select any block — rows, columns, or a rectangle — and a **Selection** panel appears with the cell
count, how many are empty, a *skip cells that already have content* toggle, a live progress bar, a
per-cell log, and a **Stop** button.

The loop runs in the sidebar, one request per cell, which means real progress, a working stop, and
no exposure to Apps Script's 6-minute execution ceiling. Your cursor never moves. Whole-column
selections are clamped to the used range, so clicking a column header gives you the rows that
exist rather than a million empty ones.

### Model list that can't go stale

Groq retires models on a rolling schedule. The sidebar fetches the live list from your key on open
and on demand, filters out anything that can't serve chat completions, and reports a 404 as
*"this model has most likely been decommissioned"* rather than a raw API error.

---

## Install

You need a free Groq API key: **https://console.groq.com/keys**

### Manual — about five minutes, no tooling

1. Open the Google Sheet you want this in.
2. **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with [`src/Code.gs`](src/Code.gs).
4. **＋ → HTML**, name it exactly `Sidebar`, replace the boilerplate with [`src/Sidebar.html`](src/Sidebar.html).
5. ⚙️ **Project Settings** → tick *Show `appsscript.json` manifest file in editor*. Open
   `appsscript.json`, replace with [`src/appsscript.json`](src/appsscript.json).
6. **Save**, then **reload the spreadsheet tab** — the menu is built on page load.
7. **AI Autocomplete → Open assistant**. Approve the authorisation prompt.
8. Paste your key into **Settings → Save key**, then **Test connection**.

> [!NOTE]
> Google will warn that the app "isn't verified". That's expected for a private script you own —
> click *Advanced → Go to (project name)*.

> [!IMPORTANT]
> Step 4 is where installs usually go wrong. If the sidebar opens but is blank, the `Sidebar` file
> is still Apps Script's empty boilerplate. Select all, delete, paste again.

### With clasp — for version control

```bash
npm install -g @google/clasp
clasp login

git clone https://github.com/<you>/groq-sheets-autocomplete
cd groq-sheets-autocomplete
cp .clasp.json.example .clasp.json    # add your Script ID
clasp push
```

Script ID lives in **Extensions → Apps Script → Project Settings**.

---

## Using it

**Start with the brief.** Open **Project context** and describe your product in six to ten lines,
including any word your sheet uses as jargon. This is what separates useful suggestions from
plausible-sounding filler:

```text
Flownix is a gym management app. This sheet is the QA bug backlog; one row is one defect.
Main flows: walk-in enquiry, membership signup, pantry orders, class booking, SSO login.
"Flow" means a guided multi-step screen sequence, not a general workflow.
Column B is the reported symptom, column C is the fix instruction for developers.
```

**Then set tone.** See [presets](#presets) for a ready-made one.

**Then work.** Type a few words, `Ctrl`/`Cmd`+`Enter`, press `1`. For bulk work, select a range and
use the Selection panel — but run five rows first and check the output before committing a
hundred requests to it.

---

## Configuration

Everything lives in the sidebar and is stored per-user in `PropertiesService.getUserProperties()`.
Collaborators on the same sheet never see your key or your settings.

| Setting | Default | Notes |
|---|---|---|
| Model | `openai/gpt-oss-20b` | Refreshed live from your key |
| Project context | empty | Max 4000 chars, sent first in every prompt |
| Send workbook structure | on | Tab names + column headers |
| Tone / style instructions | empty | Highest-priority instruction |
| Use column header | on | |
| Use other values in the row | on | Sends `Header: value` pairs as facts |
| Use nearby entries in column | on | Biggest quality lever |
| Suggestions | 3 | 1–5 |
| Example rows ± | 8 | Lower it if your rows are long |
| Creativity | 0.4 | `temperature` |
| Response length | Normal (500) | Token ceiling; scales with suggestion count |
| Header row | 1 | Change if your table starts lower |

### Models and rate limits

**Never hardcode a Groq model ID.** `llama-3.1-8b-instant` and `llama-3.3-70b-versatile` were both
retired on 2026-08-16. The `MODELS` array in `Code.gs` is only a fallback for when
`GET /openai/v1/models` fails.

Free tier is roughly 30 requests/minute. Batch runs space calls ~250 ms apart and back off 4 s on
a 429. Reasoning models (`gpt-oss`, `qwen3`) get `reasoning_effort: low` so hidden thinking
doesn't eat the token budget before the answer.

---

## Presets

[`presets/product-manager-bugs.md`](presets/product-manager-bugs.md) — a tone instruction for PMs
writing bug tickets developers can pick up. Includes the reasoning behind each rule and variants
for terse boards, acceptance criteria, and reproduction-step columns.

PRs adding presets for other roles are welcome.

---

## How it works

```
Sidebar (HTML/JS)  ──poll 900ms──►  getCellContext()
      │                              ├─ header row, row values, ±N column neighbours
      │                              └─ selectionInfo_()  clamped to the used range
      │  suggest(settings)
      ▼
 sheetMap_()       ─►  tab names + headers, cached 5 min in CacheService
      ▼
 buildMessages_()  ─►  system: "cell autocomplete, strict JSON, N alternatives"
      │                user:   project brief → workbook map → column header
      │                        → row facts → column examples → tone → the cell
      ▼
 callGroq_()       ─►  POST /openai/v1/chat/completions
      │                response_format: json_object, reasoning_effort: low
      │                on a JSON-validation 400 → retry once in plain-text mode
      ▼
 parseCompletions_()  ─►  tolerant parse: fences, bare arrays, {"1":…} objects
      │                   salvage_() recovers whole strings from a truncated reply
      ▼
 applyCompletion() ─►  appends to what you typed, or replaces an empty cell
```

Multi-cell runs call `completeCell(sheet, a1, settings, skipFilled)` once per cell from the
client. That's deliberate: a server-side loop would hit the 6-minute ceiling, show no progress,
and couldn't be stopped.

### Project structure

| Path | Purpose |
|---|---|
| `src/Code.gs` | Context engine, prompt builder, Groq client, settings, batch |
| `src/Sidebar.html` | Entire UI — one file, no build step, no external assets |
| `src/appsscript.json` | Manifest and OAuth scopes |
| `test/harness.js` | Runs `Code.gs` in a Node VM with stubbed Apps Script services |
| `presets/` | Ready-made tone instructions |

---

## Privacy

There is no backend. Your sheet data goes from your Apps Script project straight to Groq and
nowhere else — no proxy, no analytics, no telemetry. Your API key is stored in your own
`UserProperties` and is not visible to other editors of the sheet.

What leaves your sheet on each request: the active cell, its column header, the filled values in
that row, up to ±N entries from that column, the workbook's tab and column names, and your brief.
Turn any of those off in Settings.

Three OAuth scopes, no more:

| Scope | Why |
|---|---|
| `spreadsheets.currentonly` | Read/write **only** the sheet it's installed in — not your Drive |
| `script.container.ui` | Draw the menu and sidebar |
| `script.external_request` | Call `api.groq.com` |

---

## Tests

```bash
npm test    # 53 checks, no network, no Google account
```

`test/harness.js` runs `Code.gs` in a Node VM with `SpreadsheetApp`, `PropertiesService`,
`CacheService` and `UrlFetchApp` stubbed, then asserts the context engine, selection clamping,
prompt assembly and ordering, the tolerant parser and truncation salvage, error mapping, the
JSON-400 retry, and cell-write spacing rules.

Run it before changing the prompt — prompt regressions are otherwise invisible until they reach
your sheet.

---

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| Menu doesn't appear | Reload the **spreadsheet** tab. `onOpen()` only fires on page load. |
| Sidebar opens but is blank | The `Sidebar` file is still empty boilerplate. Select all → delete → paste. |
| `401 rejected the API key` | Re-paste it; Groq keys start with `gsk_`. |
| `404 model does not exist` | Groq retired it. Settings → **Refresh models**. |
| `400 Failed to validate JSON` | Reply truncated mid-JSON. Raise **Response length** or ask for fewer suggestions. |
| `429 rate limit` | Over 30 req/min or the model's daily cap. Turn Auto off, or wait. |
| Suggestions repeat my text | Raise Creativity, or add *"never restate the input"* to the tone box. |
| Suggestions ignore my style | Turn on *Use nearby entries in column* and raise **Example rows ±**. |
| Suggestions don't know my product | Fill in **Project context**. Column examples teach style; only the brief teaches domain. |

---

## Limitations

**No inline ghost text.** Apps Script cannot paint into a cell while you type — the grid is
Google's canvas and add-ons have no API for it. The sidebar is the closest thing that works in
real Google Sheets. Copilot-style inline completion would require a self-hosted spreadsheet app.

**Desktop browser only.** Custom menus, sidebars and add-ons don't run in the Google Sheets mobile
apps at all.

**Polling, not events.** There is no selection-change event for sidebars, so the active cell is
polled every 900 ms.

**No streaming.** `UrlFetchApp` has no streaming support. Short `max_tokens` keeps it responsive.

---

## Contributing

PRs welcome. Please run `npm test` first, and add a check for any behaviour you change — the
harness needs no Google account or network.

Good first contributions:

- Presets for other roles (sales notes, research logs, content calendars)
- A `=AI_COMPLETE(A2)` custom function for fill-down
- Undo for a batch run
- Per-sheet rather than per-user project briefs

Everything provider-specific is `GROQ_ENDPOINT`, `MODELS_ENDPOINT` and `callGroq_`. Any
OpenAI-compatible endpoint drops in unchanged.

---

## License

MIT — see [LICENSE](LICENSE). Use it, fork it, ship it.
