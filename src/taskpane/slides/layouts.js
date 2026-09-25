/**
 * Slide layout drawers (moved from the POC 1 spike, see doc/poc-1-rendering-results.md).
 *
 * The backend's AI picks a `layout` per slide; each drawer turns that data into PptxGenJS shapes
 * sized to the real slide, using only theme references (colours `tx1`/`accent1`/…, fonts
 * `+mj-lt`/`+mn-lt`) so the destination deck's theme styles them on insert. Text that doesn't fit
 * goes through: shrink to the floor size → reflow (2 columns / vertical) → split onto
 * continuation slides "(2/2)".
 */

import PptxGenJS from 'pptxgenjs';
import { FIT, fitParas } from './measure';

const { SIZE, TEXT_MARGIN_PT, LINE_HEIGHT, PARA_AFTER_PT, BULLET_INDENT_PT, CELL_MARGIN_IN, MIN_COLUMN_IN } = FIT;

// Theme references only — resolved by the destination deck's theme on insert.
const FONT = { heading: '+mj-lt', body: '+mn-lt' };
const COLOR = {
    text: 'tx1',
    muted: 'tx2',
    onAccent: 'bg1',
    panel: 'bg2',
    accent: 'accent1',
    accent2: 'accent2',
    positive: '2E7D32', // fixed on purpose: "correct" must never turn purple with a theme
    negative: 'C62828',
    onFixed: 'FFFFFF',
};

// ── Geometry ─────────────────────────────────────────────────────────────────

function gridFor(W, H) {
    const m = Math.max(0.35, W * 0.05);
    const title = { x: m, y: m * 0.6, w: W - 2 * m, h: H * 0.15 };
    const bottom = H - 0.45; // bottom margin (the POC kept a footer strip here)
    const top = title.y + title.h + H * 0.04;
    const content = { x: m, y: top, w: W - 2 * m, h: bottom - top };
    return { W, H, m, title, content, gap: W * 0.025 };
}

// ── Deck builder ─────────────────────────────────────────────────────────────

export class Deck {
    constructor({ W, H, measurer }) {
        this.pptx = new PptxGenJS();
        this.pptx.defineLayout({ name: 'DEST', width: W, height: H });
        this.pptx.layout = 'DEST';
        this.W = W;
        this.H = H;
        this.grid = gridFor(W, H);
        this.m = measurer;
        this.slideCount = 0;
    }

    newSlide() {
        this.slideCount++;
        return this.pptx.addSlide();
    }

    text(slide, content, opts) {
        slide.addText(content, { margin: TEXT_MARGIN_PT, valign: 'top', wrap: true, ...opts });
    }

    // Title band shared by content layouts.
    frame(title) {
        const slide = this.newSlide();
        const { title: box } = this.grid;
        const fit = fitParas(this.m, [{ text: title }], box, { role: 'heading', size: SIZE.title });
        this.text(slide, title, {
            ...box, valign: 'bottom', fontFace: FONT.heading, fontSize: fit.pt, color: COLOR.text,
        });
        slide.addShape(this.pptx.ShapeType.rect, {
            x: box.x + TEXT_MARGIN_PT / 72, y: box.y + box.h + 0.04, w: Math.min(1.4, box.w * 0.15), h: 0.07,
            fill: { color: COLOR.accent }, line: { type: 'none' },
        });
        return slide;
    }
}

// ── Text-run helpers ─────────────────────────────────────────────────────────

// Emphasise the first occurrence of `highlight` (case-sensitive — the backend sends the exact substring).
function runsFor(text, highlight, base, highlightColor = COLOR.accent) {
    const i = highlight ? text.indexOf(highlight) : -1;
    if (i < 0) return [{ text, options: { ...base } }];
    return [
        i > 0 && { text: text.slice(0, i), options: { ...base } },
        { text: highlight, options: { ...base, bold: true, color: highlightColor } },
        i + highlight.length < text.length && { text: text.slice(i + highlight.length), options: { ...base } },
    ].filter(Boolean);
}

// Join per-paragraph run arrays into one PptxGenJS run list (breakLine ends a paragraph).
function joinParagraphs(paragraphRuns) {
    const out = [];
    paragraphRuns.forEach((runs, i) => {
        runs.forEach((run, j) => {
            const isLast = j === runs.length - 1 && i < paragraphRuns.length - 1;
            out.push({ text: run.text, options: { ...run.options, ...(isLast ? { breakLine: true } : {}) } });
        });
    });
    return out;
}

// Greedy split of items into pages that each fit `fitsPage(items)`.
function paginate(items, fitsPage) {
    const pages = [];
    let current = [];
    for (const item of items) {
        const candidate = [...current, item];
        if (current.length && !fitsPage(candidate)) {
            pages.push(current);
            current = [item];
        } else {
            current = candidate;
        }
    }
    if (current.length) pages.push(current);
    return pages;
}

const pageTitle = (title, i, n) => (n > 1 ? `${title} (${i + 1}/${n})` : title);

// ── Layout drawers ───────────────────────────────────────────────────────────

export function drawDeckTitle(deck, title) {
    const { W, H } = deck;
    const slide = deck.newSlide();
    slide.addShape(deck.pptx.ShapeType.rect, {
        x: 0, y: 0, w: Math.max(0.25, W * 0.035), h: H, fill: { color: COLOR.accent }, line: { type: 'none' },
    });
    const titleBox = { x: W * 0.1, y: H * 0.3, w: W * 0.8, h: H * 0.34 };
    const fit = fitParas(deck.m, [{ text: title }], titleBox, { role: 'heading', size: SIZE.deckTitle });
    deck.text(slide, title, {
        ...titleBox, valign: 'middle', fontFace: FONT.heading, fontSize: fit.pt, color: COLOR.text,
    });
}

function drawBullets(deck, data) {
    const { content, gap } = deck.grid;
    const barW = 0.08;
    const inner = { x: content.x + barW + 0.2, y: content.y, w: content.w - barW - 0.2, h: content.h };
    const paras = (items) => items.map((text) => ({ text }));
    const fitCol = (items, box) => fitParas(deck.m, paras(items), box, { indentPt: BULLET_INDENT_PT });

    const colW = (inner.w - gap) / 2;
    const col1 = { ...inner, w: colW };
    const col2 = { ...inner, x: inner.x + colW + gap, w: colW };

    let pages;
    const single = fitCol(data.items, inner);
    if (single.fits) {
        pages = [{ columns: [data.items], pt: single.pt }];
    } else {
        const half = Math.ceil(data.items.length / 2);
        const halves = [data.items.slice(0, half), data.items.slice(half)];
        const f1 = fitCol(halves[0], col1);
        const f2 = fitCol(halves[1], col2);
        if (colW >= MIN_COLUMN_IN && f1.fits && f2.fits) {
            pages = [{ columns: halves, pt: Math.min(f1.pt, f2.pt) }];
        } else {
            const chunks = paginate(data.items, (items) => fitCol(items, inner).fits);
            pages = chunks.map((items) => ({ columns: [items], pt: fitCol(items, inner).pt }));
        }
    }

    pages.forEach((page, i) => {
        const slide = deck.frame(pageTitle(data['slide-title'], i, pages.length));
        slide.addShape(deck.pptx.ShapeType.rect, {
            x: content.x, y: content.y, w: barW, h: content.h, fill: { color: COLOR.accent }, line: { type: 'none' },
        });
        const boxes = page.columns.length === 2 ? [col1, col2] : [inner];
        page.columns.forEach((items, c) => {
            const runs = joinParagraphs(items.map((text) => [{
                text,
                options: { bullet: { indent: BULLET_INDENT_PT }, paraSpaceAfter: PARA_AFTER_PT },
            }]));
            deck.text(slide, runs, { ...boxes[c], fontFace: FONT.body, fontSize: page.pt, color: COLOR.text });
        });
    });
}

function drawPattern(deck, data) {
    const { content, gap } = deck.grid;
    const n = data.parts.length;
    const arrowW = 0.45;
    const chipW = (content.w - (n - 1) * (arrowW + 2 * 0.1)) / n;
    const chipH = content.h * 0.34;
    const chipFit = (w, h) => {
        const fits = data.parts.map((p) => fitParas(deck.m, [{ text: p }], { w, h }, { size: SIZE.chip, bold: true }));
        return { pt: Math.min(...fits.map((f) => f.pt)), fits: fits.every((f) => f.fits) };
    };

    const horizontal = chipFit(chipW, chipH);
    const vertical = chipW < 1.6 || !horizontal.fits;
    const slide = deck.frame(data['slide-title']);
    const example = data.example || { text: '', highlight: null };

    let exampleBox;
    if (!vertical) {
        const y = content.y + content.h * 0.06;
        data.parts.forEach((part, i) => {
            const x = content.x + i * (chipW + arrowW + 0.2);
            deck.text(slide, part, {
                x, y, w: chipW, h: chipH, shape: deck.pptx.ShapeType.roundRect, rectRadius: 0.12,
                fill: { color: COLOR.accent }, align: 'center', valign: 'middle',
                fontFace: FONT.body, fontSize: horizontal.pt, bold: true, color: COLOR.onAccent,
            });
            if (i < n - 1) {
                slide.addShape(deck.pptx.ShapeType.rightArrow, {
                    x: x + chipW + 0.1, y: y + chipH / 2 - 0.2, w: arrowW, h: 0.4,
                    fill: { color: COLOR.accent2 }, line: { type: 'none' },
                });
            }
        });
        const top = y + chipH + content.h * 0.12;
        exampleBox = { x: content.x, y: top, w: content.w, h: content.y + content.h - top };
    } else {
        const colW = content.w * 0.45;
        const arrowH = 0.3;
        const vChipH = (content.h - (n - 1) * (arrowH + 0.1)) / n;
        const chipPt = chipFit(colW, vChipH).pt;
        data.parts.forEach((part, i) => {
            const y = content.y + i * (vChipH + arrowH + 0.1);
            deck.text(slide, part, {
                x: content.x, y, w: colW, h: vChipH, shape: deck.pptx.ShapeType.roundRect, rectRadius: 0.1,
                fill: { color: COLOR.accent }, align: 'center', valign: 'middle',
                fontFace: FONT.body, fontSize: chipPt, bold: true, color: COLOR.onAccent,
            });
            if (i < n - 1) {
                slide.addShape(deck.pptx.ShapeType.downArrow, {
                    x: content.x + colW / 2 - 0.2, y: y + vChipH + 0.05, w: 0.4, h: arrowH,
                    fill: { color: COLOR.accent2 }, line: { type: 'none' },
                });
            }
        });
        exampleBox = { x: content.x + colW + gap, y: content.y, w: content.w - colW - gap, h: content.h };
    }

    if (!example.text) return;
    const exFit = fitParas(deck.m, [{ text: example.text }], { ...exampleBox, w: exampleBox.w - 0.4 }, {});
    slide.addShape(deck.pptx.ShapeType.roundRect, {
        ...exampleBox, rectRadius: 0.1, fill: { color: COLOR.panel }, line: { type: 'none' },
    });
    deck.text(slide, runsFor(example.text, example.highlight, { fontSize: exFit.pt }), {
        x: exampleBox.x + 0.2, y: exampleBox.y, w: exampleBox.w - 0.4, h: exampleBox.h, valign: 'middle',
        fontFace: FONT.body, fontSize: exFit.pt, italic: true, color: COLOR.text,
    });
}

function drawComparison(deck, data) {
    const { content, gap } = deck.grid;
    const colW = (content.w - gap) / 2;
    const headH = Math.min(0.7, content.h * 0.16);
    const bodyBox = (c) => ({ x: content.x + c * (colW + gap) + 0.15, y: content.y + headH + 0.1, w: colW - 0.3, h: content.h - headH - 0.2 });
    const sides = [data.left, data.right];
    const paras = (items) => items.map((it) => ({ text: it.text }));

    // Rows = the longer side, so no item is dropped when the AI sends uneven sides.
    const rowCount = Math.max(data.left.items.length, data.right.items.length);
    const rows = Array.from({ length: rowCount }, (_, i) => i);
    const fitRows = (idx) => sides.map((side, c) => fitParas(deck.m, paras(idx.map((i) => side.items[i]).filter(Boolean)), bodyBox(c), {}));
    const allFit = (idx) => fitRows(idx).every((f) => f.fits);
    const pages = allFit(rows) ? [rows] : paginate(rows, allFit);

    const headerStyle = (side, c) => {
        if (side.polarity === 'positive') return { fill: COLOR.positive, color: COLOR.onFixed, mark: '✓ ', hl: COLOR.positive };
        if (side.polarity === 'negative') return { fill: COLOR.negative, color: COLOR.onFixed, mark: '✗ ', hl: COLOR.negative };
        return { fill: c === 0 ? COLOR.accent : COLOR.accent2, color: COLOR.onAccent, mark: '', hl: COLOR.accent };
    };

    pages.forEach((idx, p) => {
        const slide = deck.frame(pageTitle(data['slide-title'], p, pages.length));
        const pt = Math.min(...fitRows(idx).map((f) => f.pt));
        sides.forEach((side, c) => {
            const st = headerStyle(side, c);
            const x = content.x + c * (colW + gap);
            deck.text(slide, `${st.mark}${side.heading}`, {
                x, y: content.y, w: colW, h: headH, fill: { color: st.fill }, align: 'center', valign: 'middle',
                fontFace: FONT.heading, fontSize: 24, bold: true, color: st.color,
            });
            slide.addShape(deck.pptx.ShapeType.rect, {
                x, y: content.y + headH, w: colW, h: content.h - headH, fill: { color: COLOR.panel }, line: { type: 'none' },
            });
            const items = idx.map((i) => side.items[i]).filter(Boolean);
            if (!items.length) return;
            const runs = joinParagraphs(items.map((it) => runsFor(it.text, it.highlight, { paraSpaceAfter: PARA_AFTER_PT }, st.hl)));
            deck.text(slide, runs, { ...bodyBox(c), fontFace: FONT.body, fontSize: pt, color: COLOR.text });
        });
    });
}

function drawVocabTable(deck, data) {
    const { content } = deck.grid;

    // Column widths proportional to the longest text per column, clamped to ≥ 20%.
    const natural = data.columns.map((_, c) => {
        const texts = [data.columns[c], ...data.rows.map((r) => r[c] ?? '')];
        return Math.max(...texts.map((t) => deck.m.width(t, 'body', SIZE.table.start, c === 0)));
    });
    const raw = natural.map((n) => Math.max(0.2, n / natural.reduce((a, b) => a + b, 0)));
    const colW = raw.map((r) => (r / raw.reduce((a, b) => a + b, 0)) * content.w);

    const rowHeightIn = (row, pt, bold) => {
        const lines = row.map((cell, c) => deck.m.lineCount(cell, 'body', pt, colW[c] * 72 - 2 * CELL_MARGIN_IN.x * 72, bold));
        return (Math.max(...lines) * pt * LINE_HEIGHT) / 72 + 2 * CELL_MARGIN_IN.y;
    };
    const tableHeight = (rows, pt) => rowHeightIn(data.columns, pt, true) + rows.reduce((a, r) => a + rowHeightIn(r, pt, false), 0);

    let pt = SIZE.table.start;
    while (pt > SIZE.table.floor && tableHeight(data.rows, pt) > content.h) pt -= 2;
    const pages = tableHeight(data.rows, pt) <= content.h
        ? [data.rows]
        : paginate(data.rows, (rows) => tableHeight(rows, SIZE.table.floor) <= content.h);
    if (pages.length > 1) pt = SIZE.table.floor;

    pages.forEach((rows, p) => {
        const slide = deck.frame(pageTitle(data['slide-title'], p, pages.length));
        const header = data.columns.map((text) => ({
            text, options: { bold: true, fill: { color: COLOR.accent }, color: COLOR.onAccent },
        }));
        const body = rows.map((row, r) => data.columns.map((_, c) => ({
            text: row[c] ?? '',
            options: {
                fill: { color: r % 2 ? COLOR.panel : COLOR.onAccent },
                color: c === 0 ? COLOR.accent : COLOR.text,
                bold: c === 0,
            },
        })));
        const rowH = [rowHeightIn(data.columns, pt, true), ...rows.map((r) => rowHeightIn(r, pt, false))];
        slide.addTable([header, ...body], {
            x: content.x, y: content.y, w: content.w, colW, rowH,
            fontFace: FONT.body, fontSize: pt, valign: 'middle', border: { type: 'none' },
        });
    });
}

function drawExamples(deck, data) {
    const { content } = deck.grid;
    const gapY = 0.15;
    const pad = 0.12;
    const textW = content.w - 0.5;
    const itemHeightIn = (it, pt) => {
        const paras = [{ text: it.text }];
        if (it.translation) paras.push({ text: it.translation, pt: Math.round(pt * 0.8) });
        return deck.m.heightPt(paras, 'body', pt, textW * 72 - 2 * TEXT_MARGIN_PT) / 72 + 2 * TEXT_MARGIN_PT / 72 + 2 * pad;
    };
    const stackHeight = (items, pt) => items.reduce((a, it) => a + itemHeightIn(it, pt), 0) + gapY * (items.length - 1);

    let pt = SIZE.body.start;
    while (pt > SIZE.body.floor && stackHeight(data.items, pt) > content.h) pt -= 2;
    const pages = stackHeight(data.items, pt) <= content.h
        ? [data.items]
        : paginate(data.items, (items) => stackHeight(items, SIZE.body.floor) <= content.h);
    if (pages.length > 1) pt = SIZE.body.floor;

    pages.forEach((items, p) => {
        const slide = deck.frame(pageTitle(data['slide-title'], p, pages.length));
        let y = content.y;
        items.forEach((it) => {
            const h = itemHeightIn(it, pt);
            slide.addShape(deck.pptx.ShapeType.rect, {
                x: content.x, y, w: content.w, h, fill: { color: COLOR.panel }, line: { type: 'none' },
            });
            slide.addShape(deck.pptx.ShapeType.rect, {
                x: content.x, y, w: 0.08, h, fill: { color: COLOR.accent }, line: { type: 'none' },
            });
            const runs = [runsFor(it.text, it.highlight, { paraSpaceAfter: PARA_AFTER_PT })];
            if (it.translation) {
                runs.push([{ text: it.translation, options: { fontSize: Math.round(pt * 0.8), italic: true, color: COLOR.muted } }]);
            }
            deck.text(slide, joinParagraphs(runs), {
                x: content.x + 0.3, y: y + pad, w: textW, h: h - 2 * pad, valign: 'middle',
                fontFace: FONT.body, fontSize: pt, color: COLOR.text,
            });
            y += h + gapY;
        });
    });
}

export const DRAWERS = {
    bullets: drawBullets,
    pattern: drawPattern,
    comparison: drawComparison,
    'vocab-table': drawVocabTable,
    examples: drawExamples,
};
