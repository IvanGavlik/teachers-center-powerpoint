# POC 1: Slide layout rendering results

Branch `poc/slide-layouts` (throwaway). Code: `src/taskpane/poc/`. Plan: AI picks the layout, the add-in fits the text. See `teachers-center-backend/doc/slide-UIUX-investigation.md` Part 5.

## How to run

1. `npm run dev-server`, then `npm run start:dev` (desktop), or sideload `manifest.dev.xml` in PowerPoint for the web.
2. Open a test deck and complete the add-in's settings modal.
3. In the chat, type:

   ```
   /poc-layouts <variant> <size>
   ```

   - `variant`: `typical` (B1 English), `max` (German, at the proposed limits) or `overflow` (over the limits).
   - `size`: `match` (built at the deck's detected size) or `default` (PptxGenJS 10×5.625in, which is what the add-in does today).
4. The slides are added at the end of the deck. The chat shows a summary; the full JSON is in the console and in `window.__pocLastReport`.
5. Each slide has a footer: `POC · layout · variant · size · title pt · body pt · action`. The action is what the fit logic did: `fit`, `reflow-2col`, `reflow-vertical` or `split-N`.
6. The last slide is **Calibration**. Its boxes start at our estimated height; PowerPoint then resizes them to fit their text. `errPct` is how far the estimate was off (+ means our estimate was taller).
7. To test theme switching: Design tab → pick another theme → screenshot again.

Screenshots go in `doc/poc-1/` named `<platform>-<deck>-<variant>-<size>[-switched].png`.

## Test decks

| Deck | How to make it |
|---|---|
| **A: Default 16:9** | New blank presentation, Office Theme (13.33×7.5in) |
| **B: Dark theme** | New presentation with a dark built-in theme (e.g. *Ion* or *Mesh*) |
| **C: 4:3** | Design → Slide Size → Standard (4:3) (10×7.5in) |

For each deck and platform, run: `typical match`, `max match`, `overflow match`, `max default`.

## Results

Fill with ✅ / ❌ / values. "Scale" and "Fit error" come from the chat summary.

| Platform | Deck | Run | Size detected (source) | Colours follow theme | Fonts follow theme | Re-colour on switch | Geometry max dev / scale | Table OK | Fit error (calibration) | Visible overflow | Insert ms |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Desktop | A | typical match | | | | | | | | | |
| Desktop | A | max match | | | | | | | | | |
| Desktop | A | overflow match | | | | | | | | | |
| Desktop | A | max default | | | | | | | | | |
| Desktop | B | max match | | | | | | | | | |
| Desktop | B | max default | | | | | | | | | |
| Desktop | C | max match | | | | | | | | | |
| Desktop | C | max default | | | | | | | | | |
| Web | A | typical match | | | | | | | | | |
| Web | A | max match | | | | | | | | | |
| Web | A | overflow match | | | | | | | | | |
| Web | A | max default | | | | | | | | | |
| Web | B | max match | | | | | | | | | |
| Web | B | max default | | | | | | | | | |
| Web | C | max match | | | | | | | | | |
| Web | C | max default | | | | | | | | | |

## Questions to answer

1. **Theme colours**: do `accent1`, `tx1` and `bg2` take the deck's colours on insert, on both platforms? Do they change when the theme is switched?
2. **Theme fonts**: do `+mj-lt` and `+mn-lt` show the deck's heading and body fonts?
3. **Slide size**: which source detected the size on web (`pageSetup` or fallback)? With `default`, are slides scaled (scale ≈ deck/10in) or pinned to the top-left (scale 1.0)?
4. **PptxGenJS on desktop**: do all six layouts, including the table, insert without errors?
5. **Fit**: is the calibration error within about ±10%? Does any `max` slide visibly overflow? Are the `overflow` splits and reflows sensible?

## Recorded results (2026-09-26)

Only deck A (default 16:9) with `max match` was recorded with numbers. Earlier runs of all variants were checked by eye: everything looked OK except `overflow match`, where the last content slide spilled slightly.

| Platform | Deck | Run | Size detected (source) | Geometry | Fit error (calibration) | Responsive actions | Build / insert ms |
|---|---|---|---|---|---|---|---|
| Desktop | A | max match | 13.33×7.50in via office.js `pageSetup` | 64 shapes, max dev 0pt, scale 1×1, 0 missing | 0% ×6, confirmed by hand* | examples split-2 (18pt) | 202 / 149 |
| Web | A | max match | 13.33×7.50in via office.js `pageSetup` | 64 shapes, max dev 0pt, scale 1×1, 0 missing | 0% ×6 | examples split-2 (18pt) | 150 / 1300 |

\* The read-back looked suspicious (all exactly 0%), so it was checked by hand on desktop: each calibration box was dragged taller, then set to "Resize shape to fit text". Every box snapped back to its original (estimated) height.

Not run: dark theme (B), 4:3 (C), theme switching, `default` size. `default` no longer matters, because the production renderer always builds at the detected size.

## Decision

- **Use PptxGenJS on both platforms, with theme colours: go.** Rendering, tables and geometry are identical on desktop and web.
- **Slide size on web:** the office.js `pageSetup` API works on both platforms. Use it; keep the presentation.xml and 16:9 fallbacks only as a safety net.
- **Text estimate:** it's exact when the line count is predicted correctly (the height formula matches PowerPoint's). When it's wrong, it's off by a whole line. The likely cause of the small `overflow` spill is a borderline word that wrapped in PowerPoint but not in our estimate. For production: measure against a ~3–5% narrower width as a safety margin.
- **Limits:** keep the POC 2 backend limits (5 bullets, 2–4 pattern parts, 4 per comparison side, 8 table rows, 5 examples). At the limit, 5 long German examples with translations split onto 2 slides at 18pt. That's acceptable; the AI's real examples in the eval were much shorter.
- **Surprises:** the add-in's own calibration read-back can't be trusted to show a resize, so it has to be checked by hand. Web insert is ~1.3 s vs ~0.15 s on desktop, which is fine.

## Step 4: production renderer (2026-09-26)

The POC code (`src/taskpane/poc/`, `/poc-layouts`) has been removed. The renderer now lives in `src/taskpane/slides/`:

| File | What it does |
|---|---|
| `measure.js` | Canvas text measurement and `fitParas`. All sizing constants are in `FIT`, including the new `widthSafety: 0.96` margin. |
| `slideSize.js` | `detectSlideSize` (pageSetup → presentation.xml → 16:9), `toSlideFormat` (sent to the backend as `slide-format`), `readDocumentZip` |
| `layouts.js` | The five layout drawers + the title slide (no subtitle). A comparison now uses the longer side's item count, so no items are dropped. |
| `renderDeck.js` | `buildDeckBase64(deck, size, themeFonts)` → `{base64, slideCount, splits, skipped}`, and `insertDeckBase64` |
| `previewCards.js` | Mini-slide HTML preview per layout |
| `devFixtures.js` | Sample decks for the dev-only `/dev-deck <typical\|max\|overflow>` command, which runs them through the real preview and insert |

Desktop and web use the same insert path. The interactivity (QR) slide is now built at the deck's real size.
