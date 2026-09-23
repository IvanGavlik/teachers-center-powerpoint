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

## Decision (fill in after testing)

- Use PptxGenJS on both platforms, with theme colours: **go / no-go**
- How to get the slide size on web:
- Is the canvas text estimate good enough, or do the schema limits need tightening?
- Limits to use in POC 2 (bullets × chars, table rows, items per side):
- Surprises:
