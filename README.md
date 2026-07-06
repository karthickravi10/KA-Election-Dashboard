# Karnataka Election Dashboard

A single-page, client-side dashboard for exploring Karnataka Assembly
Constituency (AC) profiles — winner/runner-up results, dominant caste
composition, and 2024 GE booth gradation — with a searchable AC-name
selector. Pure HTML/CSS/JS, no backend, no build step.

---

## 1. Project structure

```
election-dashboard/
├── index.html            # Markup: header, search panel, card grid
├── css/
│   └── styles.css        # All styling (design tokens, layout, responsive rules)
├── js/
│   ├── data.js            # Default dataset, embedded as a plain JS object
│   └── app.js             # Data loading, parsing, search, rendering logic
├── data/
│   ├── KA_Raw_Data.xlsx   # Bundled dataset (cleaned: single header row) — for reference / re-upload
│   └── KA_Raw_Data.csv    # Same data, CSV format — for reference / re-upload
└── README.md
```

## 2. How it works

- `js/data.js` embeds the full default dataset as a plain JS variable
  (`window.__ELECTION_RECORDS__`), loaded with a normal `<script>` tag.
  **This is why the dashboard works with zero setup** — just open
  `index.html` — even by double-clicking it. No server, no network
  call, and no dependency on any external library is needed for the
  default view.
- Every row is normalized into a plain JS object keyed by short field
  names (see `COLUMN_MAP` in `app.js`), so the rest of the app never has
  to deal with raw spreadsheet column headers.
- The **only** thing that needs internet access is uploading a *new*
  `.xlsx`/`.xls` file via "Load a different file" — that goes through
  the [SheetJS](https://sheetjs.com/) library, loaded from a CDN.
  Uploading a `.csv` instead needs no external library at all (parsed
  with a small built-in CSV parser), so it always works offline.
- The **search box** filters constituencies by `AC Name` as you type
  (case-insensitive, substring match), with full keyboard support
  (`↑ ↓` to move, `Enter` to select, `Esc` to close).
- Selecting a constituency calls `renderDashboard()`, which fills in
  every card. Two cells get special treatment because they contain
  multi-line, delimited text in the source file:
  - **Dominant Caste 1–5** — split into a headline (caste + % share)
    and any sub-breakdown lines, shown as a compact card per caste.
  - **2024 GE Booth Gradation** — parsed into a total booth count and
    Grade A–D counts, rendered as a proportional stacked bar + legend.
- You can also load a **different** `.xlsx`/`.csv` file at any time via
  "Load a different file" — it must use the same column headers.

## 3. Running locally

**Just double-click `index.html`.** The default dataset is embedded
directly in `js/data.js`, so the dashboard opens and works fully
offline — no server required.

If you prefer serving it (e.g. for the cleanest experience while also
testing file uploads), any of these work too:

```bash
cd election-dashboard
python3 -m http.server 8080   # then open http://localhost:8080
```
```bash
npx serve election-dashboard
```
Or, in VS Code: install the "Live Server" extension, right-click
`index.html` → *Open with Live Server*.

No server-side code, database, or API key is required anywhere in this project.

### If the dropdown still looks empty

- Open your browser's DevTools console (F12) and check for a red error
  mentioning `js/data.js` — that means the file didn't load. Make sure
  the `js/` folder sits next to `index.html` and wasn't renamed or left
  out when copying the project.
- If you're using "Load a different file" to upload an `.xlsx` and
  nothing happens, your network may be blocking the SheetJS CDN script.
  Save the file as `.csv` instead — CSV uploads don't need that library.

## 4. Using your own data

The app expects these exact column headers (order doesn't matter):

```
Sl No, AC No, AC Name, Zone, PC, District, Org District,
Dominant Caste 1, Dominant Caste 2, Dominant Caste 3, Dominant Caste 4, Dominant Caste 5,
Winner - Name, Winner - Party, Winner - Category, Winner - Caste,
Runner Up - Name, Runner Up - Party, Runner Up - Category, Runner Party - Caste,
2024 GE Booth Gradation
```

To use a different dataset by default, replace `data/KA_Raw_Data.xlsx`
with your own file (same name), or change `DEFAULT_DATA_URL` at the top
of `js/app.js`. To load ad-hoc without touching files, use the in-app
"Load a different file" picker — it accepts `.xlsx`, `.xls`, and `.csv`.

Missing/blank cells are handled gracefully and shown as "Not available"
rather than breaking the layout.

## 5. Customizing the look

All design tokens (colors, fonts, radii, shadows) live at the top of
`css/styles.css` under `:root`. The palette follows the brief:

| Token | Value | Usage |
|---|---|---|
| `--color-primary` | `#FF6B00` | Accents, primary card, active states |
| `--color-secondary` | `#1F2937` | Header, headings, secondary accents |
| `--color-bg` | `#F8FAFC` | Page background |
| `--color-card` | `#FFFFFF` | Card surfaces |

Fonts are loaded from Google Fonts: **Poppins** (headings/labels),
**Inter** (body text), **JetBrains Mono** (AC numbers, booth counts —
anything numeric/tabular benefits from a monospaced figure style).

## 6. Browser support

Any modern evergreen browser (Chrome, Edge, Firefox, Safari). Uses
standard ES6 (arrow functions, `const`/`let`, template literals,
`Array` methods) — no transpilation needed.

## 7. Notes on the SheetJS dependency

`index.html` loads SheetJS from a CDN:
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
```
For a fully offline/air-gapped deployment, download that file and save
it locally (e.g. `js/xlsx.full.min.js`), then update the `<script src="...">`
path in `index.html` to point to the local copy.
