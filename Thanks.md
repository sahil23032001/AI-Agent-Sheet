# Contributing

Thanks for taking a look.

## Setup

No build step. `src/Code.gs` and `src/Sidebar.html` are the whole app.

```bash
git clone https://github.com/<you>/groq-sheets-autocomplete
cd groq-sheets-autocomplete
npm test          # runs the harness; no Google account or network needed
```

To try changes in a real sheet, either paste the two files into the Apps Script editor, or use
clasp:

```bash
cp .clasp.json.example .clasp.json    # add your Script ID
clasp push
```

## Ground rules

- **Run `npm test` before opening a PR**, and add a check for any behaviour you change.
  Prompt regressions are invisible until they reach someone's spreadsheet.
- **Keep it dependency-free.** Apps Script has no package manager at runtime; the sidebar loads
  no external CSS or JS on purpose.
- **ES5-flavoured JS in `Code.gs`.** The V8 runtime supports modern syntax, but the file is
  written with `var` and `function` throughout — match the surrounding style.
- **Don't widen the OAuth scopes.** `spreadsheets.currentonly` is a feature, not an oversight.
  A change that needs full Drive access needs a discussion first.
- **No telemetry, ever.**

## Testing notes

`test/harness.js` runs `Code.gs` in a Node VM with `SpreadsheetApp`, `PropertiesService`,
`CacheService` and `UrlFetchApp` stubbed out. The fake sheet is the `GRID` array at the top —
add rows there if your feature needs different fixture data.

Two gotchas:

- Arrays created inside the VM have a different `Array` prototype, so use the `eq()` helper
  rather than `assert.deepStrictEqual` for anything returned by `Code.gs`.
- `fetchResponse` is a mutable function — set it to whatever `{code, text}` your test needs, and
  remember later tests inherit it unless you reset it.

## Good first issues

- Presets for other roles (sales notes, research logs, content calendars)
- A `=AI_COMPLETE(A2)` custom function for fill-down
- Undo for a batch run
- Per-sheet rather than per-user project briefs
- A provider adapter so any OpenAI-compatible endpoint works from the UI
