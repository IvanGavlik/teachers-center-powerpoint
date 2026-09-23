/* global PowerPoint, Office */

/**
 * POC 1 — rendering spike for structured slide layouts (throwaway, dev builds only).
 *
 * Fixture JSON in the candidate contract ({layout, …fields}) is drawn with PptxGenJS using only
 * theme scheme colours + theme fonts, fitted to the real slide size by a responsive fit pipeline
 * (estimate → shrink to floor → reflow → split), inserted with UseDestinationTheme, then checked
 * back through Office.js. See doc/poc-1-rendering-results.md.
 */

import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { FIXTURES, CALIBRATION } from './fixtures';

const EMU_PER_IN = 914400;
const PPTXGEN_DEFAULT = { w: 10, h: 5.625 };
const FALLBACK_SIZE = { w: 13.333, h: 7.5 };

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

// Font sizes (pt): start = preferred, floor = smallest readable from the back of a classroom.
const SIZE = {
    deckTitle: { start: 48, floor: 32 },
    subtitle: { start: 24, floor: 18 },
    title: { start: 36, floor: 28 },
    body: { start: 24, floor: 18 },
    table: { start: 20, floor: 16 },
    chip: { start: 24, floor: 16 },
    footer: 10,
};

const TEXT_MARGIN_PT = 5;       // inset we set on every text box (all sides)
const LINE_HEIGHT = 1.2;        // single spacing ≈ 1.2 × font size
const PARA_AFTER_PT = 6;
const BULLET_INDENT_PT = 27;    // PptxGenJS DEF_BULLET_MARGIN
const CELL_MARGIN_IN = { x: 0.1, y: 0.05 }; // PptxGenJS DEF_CELL_MARGIN_IN
const MIN_COLUMN_IN = 3.5;      // narrower than this, reflow to 2 columns is not allowed

// ── Text measurement ─────────────────────────────────────────────────────────

const measureCtx = document.createElement('canvas').getContext('2d');
const fontCache = {};

function fontAvailable(name) {
    if (!name) return false;
    if (name in fontCache) return fontCache[name];
    const probe = 'mmmmmmmmmmlliWWQ@#äöü';
    const available = ['monospace', 'serif', 'sans-serif'].some((base) => {
        measureCtx.font = `72px ${base}`;
        const baseWidth = measureCtx.measureText(probe).width;
        measureCtx.font = `72px "${name}", ${base}`;
        return measureCtx.measureText(probe).width !== baseWidth;
    });
    fontCache[name] = available;
    return available;
}

function makeMeasurer(themeFonts) {
    const resolve = (name) => {
        if (fontAvailable(name)) return { css: `"${name}"`, used: name };
        if (fontAvailable('Calibri')) return { css: '"Calibri"', used: `Calibri (fallback for ${name})` };
        return { css: 'Arial, sans-serif', used: `Arial (fallback for ${name})` };
    };
    const family = { heading: resolve(themeFonts.heading), body: resolve(themeFonts.body) };

    function width(text, role, pt, bold) {
        measureCtx.font = `${bold ? 'bold ' : ''}${pt}px ${family[role].css}`;
        return measureCtx.measureText(text).width; // px at `pt`px == points at `pt`pt
    }

    // Greedy word wrap, like PowerPoint; a single word wider than the line is broken by characters.
    function lineCount(text, role, pt, maxWidthPt, bold) {
        const fits = (s) => width(s, role, pt, bold) <= maxWidthPt;
        let total = 0;
        for (const para of String(text).split('\n')) {
            const words = para.split(/\s+/).filter(Boolean);
            let n = 1;
            let line = '';
            for (const word of words) {
                const candidate = line ? `${line} ${word}` : word;
                if (fits(candidate)) { line = candidate; continue; }
                if (line) { n++; line = ''; }
                let rest = word;
                while (!fits(rest) && rest.length > 1) {
                    let i = rest.length - 1;
                    while (i > 1 && !fits(rest.slice(0, i))) i--;
                    rest = rest.slice(i);
                    n++;
                }
                line = rest;
            }
            total += n;
        }
        return total;
    }

    // paras: [{ text, pt?, bold? }] — pt defaults to the size being tried.
    function heightPt(paras, role, pt, widthPt, bold = false) {
        const lines = paras.map((p) => {
            const size = p.pt || pt;
            return lineCount(p.text, role, size, widthPt, p.bold ?? bold) * size * LINE_HEIGHT;
        });
        return lines.reduce((a, b) => a + b, 0) + PARA_AFTER_PT * Math.max(0, paras.length - 1);
    }

    return { width, lineCount, heightPt, fonts: { heading: family.heading.used, body: family.body.used } };
}

// Try sizes from start down to floor; returns the first that fits, or the floor with fits=false.
function fitParas(m, paras, box, { role = 'body', size = SIZE.body, bold = false, indentPt = 0, scale = null }) {
    const widthPt = box.w * 72 - 2 * TEXT_MARGIN_PT - indentPt;
    const availPt = box.h * 72 - 2 * TEXT_MARGIN_PT;
    const sized = (pt) => (scale ? scale(pt) : paras);
    for (let pt = size.start; pt >= size.floor; pt -= 2) {
        const est = m.heightPt(sized(pt), role, pt, widthPt, bold);
        if (est <= availPt) return { pt, fits: true, estPt: est, availPt };
    }
    const est = m.heightPt(sized(size.floor), role, size.floor, widthPt, bold);
    return { pt: size.floor, fits: false, estPt: est, availPt };
}

// ── Geometry ─────────────────────────────────────────────────────────────────

function gridFor(W, H) {
    const m = Math.max(0.35, W * 0.05);
    const title = { x: m, y: m * 0.6, w: W - 2 * m, h: H * 0.15 };
    const footer = { x: m, y: H - 0.35, w: W - 2 * m, h: 0.3 };
    const top = title.y + title.h + H * 0.04;
    const content = { x: m, y: top, w: W - 2 * m, h: footer.y - top - 0.1 };
    return { W, H, m, title, content, footer, gap: W * 0.025 };
}

// ── Deck builder: every object is named so it can be found again after insert ─

class Deck {
    constructor({ W, H, sizeMode, variant, measurer }) {
        this.pptx = new PptxGenJS();
        if (sizeMode === 'match') {
            this.pptx.defineLayout({ name: 'POC_DEST', width: W, height: H });
            this.pptx.layout = 'POC_DEST';
        }
        this.W = W;
        this.H = H;
        this.variant = variant;
        this.grid = gridFor(W, H);
        this.m = measurer;
        this.slideCount = 0;
        this.expected = [];   // { slide, name, kind, x, y, w, h } in inches
        this.decisions = [];  // one per rendered slide
        this.calibration = [];
    }

    newSlide(layout) {
        const idx = this.slideCount++;
        return { idx, layout, slide: this.pptx.addSlide(), n: 0 };
    }

    name(s, role) {
        s.n++;
        return `poc-${s.idx}-${s.n}-${role}`;
    }

    text(s, role, content, opts) {
        const objectName = opts.objectName || this.name(s, role);
        s.slide.addText(content, { margin: TEXT_MARGIN_PT, valign: 'top', wrap: true, ...opts, objectName });
        this.expected.push({ slide: s.idx, name: objectName, kind: 'text', x: opts.x, y: opts.y, w: opts.w, h: opts.h });
    }

    shape(s, role, type, opts) {
        const objectName = this.name(s, role);
        s.slide.addShape(type, { ...opts, objectName });
        this.expected.push({ slide: s.idx, name: objectName, kind: 'shape', x: opts.x, y: opts.y, w: opts.w, h: opts.h });
    }

    table(s, role, rows, opts) {
        const objectName = this.name(s, role);
        s.slide.addTable(rows, { ...opts, objectName });
        const h = (opts.rowH || []).reduce((a, b) => a + b, 0);
        this.expected.push({ slide: s.idx, name: objectName, kind: 'table', x: opts.x, y: opts.y, w: opts.w, h });
    }

    // Title band shared by content layouts.
    frame(layout, title) {
        const s = this.newSlide(layout);
        const { title: box } = this.grid;
        const fit = fitParas(this.m, [{ text: title }], box, { role: 'heading', size: SIZE.title });
        this.text(s, 'title', title, {
            ...box, valign: 'bottom', fontFace: FONT.heading, fontSize: fit.pt, color: COLOR.text,
        });
        this.shape(s, 'title-bar', this.pptx.ShapeType.rect, {
            x: box.x + TEXT_MARGIN_PT / 72, y: box.y + box.h + 0.04, w: Math.min(1.4, box.w * 0.15), h: 0.07,
            fill: { color: COLOR.accent }, line: { type: 'none' },
        });
        s.titleFit = fit;
        return s;
    }

    footer(s, bodyPt, action) {
        const tag = [
            'POC', s.layout, this.variant, `${this.W.toFixed(2)}×${this.H.toFixed(2)}in`,
            `title ${s.titleFit ? s.titleFit.pt : '-'}pt`, `body ${bodyPt}pt`, action,
        ].filter(Boolean).join(' · ');
        this.text(s, 'footer', tag, {
            ...this.grid.footer, fontFace: FONT.body, fontSize: SIZE.footer, color: COLOR.muted, valign: 'middle',
        });
        this.decisions.push({
            slide: s.idx + 1,
            layout: s.layout,
            titlePt: s.titleFit ? s.titleFit.pt : null,
            titleFits: s.titleFit ? s.titleFit.fits : true,
            bodyPt,
            action: action || 'fit',
        });
    }
}

// ── Text-run helpers ─────────────────────────────────────────────────────────

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

function drawDeckTitle(deck, title, subtitle) {
    const { W, H } = deck;
    const s = deck.newSlide('title');
    deck.shape(s, 'accent-band', deck.pptx.ShapeType.rect, {
        x: 0, y: 0, w: Math.max(0.25, W * 0.035), h: H, fill: { color: COLOR.accent }, line: { type: 'none' },
    });
    const x = W * 0.1;
    const w = W * 0.8;
    const titleBox = { x, y: H * 0.26, w, h: H * 0.28 };
    const subBox = { x, y: H * 0.58, w, h: H * 0.14 };
    const tFit = fitParas(deck.m, [{ text: title }], titleBox, { role: 'heading', size: SIZE.deckTitle });
    const sFit = fitParas(deck.m, [{ text: subtitle }], subBox, { size: SIZE.subtitle });
    deck.text(s, 'deck-title', title, {
        ...titleBox, valign: 'bottom', fontFace: FONT.heading, fontSize: tFit.pt, color: COLOR.text,
    });
    deck.text(s, 'deck-subtitle', subtitle, {
        ...subBox, fontFace: FONT.body, fontSize: sFit.pt, color: COLOR.muted,
    });
    s.titleFit = tFit;
    deck.footer(s, sFit.pt, [!tFit.fits && 'title-overflow', !sFit.fits && 'subtitle-overflow'].filter(Boolean).join(',') || null);
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
    let action = null;
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
            action = 'reflow-2col';
        } else {
            const chunks = paginate(data.items, (items) => fitCol(items, inner).fits);
            pages = chunks.map((items) => ({ columns: [items], pt: fitCol(items, inner).pt }));
            action = `split-${chunks.length}`;
        }
    }

    pages.forEach((page, i) => {
        const s = deck.frame('bullets', pageTitle(data['slide-title'], i, pages.length));
        deck.shape(s, 'accent-bar', deck.pptx.ShapeType.rect, {
            x: content.x, y: content.y, w: barW, h: content.h, fill: { color: COLOR.accent }, line: { type: 'none' },
        });
        const boxes = page.columns.length === 2 ? [col1, col2] : [inner];
        page.columns.forEach((items, c) => {
            const runs = joinParagraphs(items.map((text) => [{
                text,
                options: { bullet: { indent: BULLET_INDENT_PT }, paraSpaceAfter: PARA_AFTER_PT },
            }]));
            deck.text(s, `bullets-col${c + 1}`, runs, {
                ...boxes[c], fontFace: FONT.body, fontSize: page.pt, color: COLOR.text,
            });
        });
        deck.footer(s, page.pt, action);
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
    const s = deck.frame('pattern', data['slide-title']);
    const exampleRuns = (pt) => runsFor(data.example.text, data.example.highlight, { fontSize: pt });

    let chipPt;
    let exampleBox;
    if (!vertical) {
        chipPt = horizontal.pt;
        const y = content.y + content.h * 0.06;
        data.parts.forEach((part, i) => {
            const x = content.x + i * (chipW + arrowW + 0.2);
            deck.text(s, `chip${i + 1}`, part, {
                x, y, w: chipW, h: chipH, shape: deck.pptx.ShapeType.roundRect, rectRadius: 0.12,
                fill: { color: COLOR.accent }, align: 'center', valign: 'middle',
                fontFace: FONT.body, fontSize: chipPt, bold: true, color: COLOR.onAccent,
            });
            if (i < n - 1) {
                deck.shape(s, `arrow${i + 1}`, deck.pptx.ShapeType.rightArrow, {
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
        chipPt = chipFit(colW, vChipH).pt;
        data.parts.forEach((part, i) => {
            const y = content.y + i * (vChipH + arrowH + 0.1);
            deck.text(s, `chip${i + 1}`, part, {
                x: content.x, y, w: colW, h: vChipH, shape: deck.pptx.ShapeType.roundRect, rectRadius: 0.1,
                fill: { color: COLOR.accent }, align: 'center', valign: 'middle',
                fontFace: FONT.body, fontSize: chipPt, bold: true, color: COLOR.onAccent,
            });
            if (i < n - 1) {
                deck.shape(s, `arrow${i + 1}`, deck.pptx.ShapeType.downArrow, {
                    x: content.x + colW / 2 - 0.2, y: y + vChipH + 0.05, w: 0.4, h: arrowH,
                    fill: { color: COLOR.accent2 }, line: { type: 'none' },
                });
            }
        });
        exampleBox = { x: content.x + colW + gap, y: content.y, w: content.w - colW - gap, h: content.h };
    }

    const exFit = fitParas(deck.m, [{ text: data.example.text }], { ...exampleBox, w: exampleBox.w - 0.4 }, {});
    deck.shape(s, 'example-panel', deck.pptx.ShapeType.roundRect, {
        ...exampleBox, rectRadius: 0.1, fill: { color: COLOR.panel }, line: { type: 'none' },
    });
    deck.text(s, 'example', exampleRuns(exFit.pt), {
        x: exampleBox.x + 0.2, y: exampleBox.y, w: exampleBox.w - 0.4, h: exampleBox.h, valign: 'middle',
        fontFace: FONT.body, fontSize: exFit.pt, italic: true, color: COLOR.text,
    });
    const action = [vertical && 'reflow-vertical', !exFit.fits && 'example-overflow'].filter(Boolean).join(',');
    deck.footer(s, `${chipPt}/${exFit.pt}`, action || null);
}

function drawComparison(deck, data) {
    const { content, gap } = deck.grid;
    const colW = (content.w - gap) / 2;
    const headH = Math.min(0.7, content.h * 0.16);
    const bodyBox = (c) => ({ x: content.x + c * (colW + gap) + 0.15, y: content.y + headH + 0.1, w: colW - 0.3, h: content.h - headH - 0.2 });
    const sides = [data.left, data.right];
    const paras = (items) => items.map((it) => ({ text: it.text }));

    const rows = data.left.items.map((_, i) => i);
    const fitRows = (idx) => sides.map((side, c) => fitParas(deck.m, paras(idx.map((i) => side.items[i]).filter(Boolean)), bodyBox(c), {}));
    const allFit = (idx) => fitRows(idx).every((f) => f.fits);
    const pages = allFit(rows) ? [rows] : paginate(rows, allFit);
    const action = pages.length > 1 ? `split-${pages.length}` : null;

    const headerStyle = (side, c) => {
        if (side.polarity === 'positive') return { fill: COLOR.positive, color: COLOR.onFixed, mark: '✓ ', hl: COLOR.positive };
        if (side.polarity === 'negative') return { fill: COLOR.negative, color: COLOR.onFixed, mark: '✗ ', hl: COLOR.negative };
        return { fill: c === 0 ? COLOR.accent : COLOR.accent2, color: COLOR.onAccent, mark: '', hl: COLOR.accent };
    };

    pages.forEach((idx, p) => {
        const s = deck.frame('comparison', pageTitle(data['slide-title'], p, pages.length));
        const pt = Math.min(...fitRows(idx).map((f) => f.pt));
        sides.forEach((side, c) => {
            const st = headerStyle(side, c);
            const x = content.x + c * (colW + gap);
            deck.text(s, `head${c + 1}`, `${st.mark}${side.heading}`, {
                x, y: content.y, w: colW, h: headH, fill: { color: st.fill }, align: 'center', valign: 'middle',
                fontFace: FONT.heading, fontSize: 24, bold: true, color: st.color,
            });
            deck.shape(s, `panel${c + 1}`, deck.pptx.ShapeType.rect, {
                x, y: content.y + headH, w: colW, h: content.h - headH, fill: { color: COLOR.panel }, line: { type: 'none' },
            });
            const items = idx.map((i) => side.items[i]).filter(Boolean);
            const runs = joinParagraphs(items.map((it) => runsFor(it.text, it.highlight, { paraSpaceAfter: PARA_AFTER_PT }, st.hl)));
            deck.text(s, `items${c + 1}`, runs, {
                ...bodyBox(c), fontFace: FONT.body, fontSize: pt, color: COLOR.text,
            });
        });
        deck.footer(s, pt, action);
    });
}

function drawVocabTable(deck, data) {
    const { content } = deck.grid;

    // Column widths proportional to the longest text per column, clamped to ≥ 20%.
    const natural = data.columns.map((_, c) => {
        const texts = [data.columns[c], ...data.rows.map((r) => r[c])];
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
    const action = pages.length > 1 ? `split-${pages.length}` : null;

    pages.forEach((rows, p) => {
        const s = deck.frame('vocab-table', pageTitle(data['slide-title'], p, pages.length));
        const header = data.columns.map((text) => ({
            text, options: { bold: true, fill: { color: COLOR.accent }, color: COLOR.onAccent },
        }));
        const body = rows.map((row, r) => row.map((text, c) => ({
            text,
            options: {
                fill: { color: r % 2 ? COLOR.panel : COLOR.onAccent },
                color: c === 0 ? COLOR.accent : COLOR.text,
                bold: c === 0,
            },
        })));
        const rowH = [rowHeightIn(data.columns, pt, true), ...rows.map((r) => rowHeightIn(r, pt, false))];
        deck.table(s, 'table', [header, ...body], {
            x: content.x, y: content.y, w: content.w, colW, rowH,
            fontFace: FONT.body, fontSize: pt, valign: 'middle', border: { type: 'none' },
        });
        deck.footer(s, pt, action);
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
    const action = pages.length > 1 ? `split-${pages.length}` : null;

    pages.forEach((items, p) => {
        const s = deck.frame('examples', pageTitle(data['slide-title'], p, pages.length));
        let y = content.y;
        items.forEach((it, i) => {
            const h = itemHeightIn(it, pt);
            deck.shape(s, `callout${i + 1}`, deck.pptx.ShapeType.rect, {
                x: content.x, y, w: content.w, h, fill: { color: COLOR.panel }, line: { type: 'none' },
            });
            deck.shape(s, `callout-bar${i + 1}`, deck.pptx.ShapeType.rect, {
                x: content.x, y, w: 0.08, h, fill: { color: COLOR.accent }, line: { type: 'none' },
            });
            const runs = [runsFor(it.text, it.highlight, { paraSpaceAfter: PARA_AFTER_PT })];
            if (it.translation) {
                runs.push([{ text: it.translation, options: { fontSize: Math.round(pt * 0.8), italic: true, color: COLOR.muted } }]);
            }
            deck.text(s, `example${i + 1}`, joinParagraphs(runs), {
                x: content.x + 0.3, y: y + pad, w: textW, h: h - 2 * pad, valign: 'middle',
                fontFace: FONT.body, fontSize: pt, color: COLOR.text,
            });
            y += h + gapY;
        });
        deck.footer(s, pt, action);
    });
}

// Boxes sized to our estimate; after insert they get autosize-to-fit so PowerPoint reports the real height.
function drawCalibration(deck) {
    const { content, W } = deck.grid;
    const s = deck.frame('calibration', 'Calibration: estimate vs. PowerPoint');
    const colX = [content.x, content.x + W * 0.42];
    const colY = [content.y, content.y];
    CALIBRATION.forEach((c, i) => {
        const col = i < 3 ? 0 : 1;
        const w = W * c.widthFrac;
        const widthPt = w * 72 - 2 * TEXT_MARGIN_PT;
        const estPt = deck.m.heightPt([{ text: c.text }], 'body', c.pt, widthPt);
        const h = (estPt + 2 * TEXT_MARGIN_PT) / 72;
        const objectName = `poc-cal-${i}`;
        deck.text(s, 'cal', c.text, {
            objectName, x: colX[col], y: colY[col], w, h, fontFace: FONT.body, fontSize: c.pt, color: COLOR.text,
            line: { color: COLOR.accent2, width: 0.75 },
        });
        deck.calibration.push({ name: objectName, pt: c.pt, widthIn: +w.toFixed(2), estHeightPt: +(h * 72).toFixed(1) });
        colY[col] += h + 0.15;
    });
    deck.footer(s, '-', 'calibration');
}

const DRAWERS = {
    bullets: drawBullets,
    pattern: drawPattern,
    comparison: drawComparison,
    'vocab-table': drawVocabTable,
    examples: drawExamples,
};

// ── Slide size detection ─────────────────────────────────────────────────────

function toUint8Array(data) {
    if (data instanceof Uint8Array) return data;
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    if (Array.isArray(data)) return new Uint8Array(data);
    if (typeof data === 'string') {
        const binary = atob(data);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return bytes;
    }
    throw new Error(`Unexpected slice data type: ${Object.prototype.toString.call(data)}`);
}

function readDocumentZip() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, async (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) return reject(result.error);
            try {
                const file = result.value;
                const slices = [];
                for (let i = 0; i < file.sliceCount; i++) {
                    const data = await new Promise((res, rej) => {
                        file.getSliceAsync(i, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? res(r.value.data) : rej(r.error)));
                    });
                    slices.push(toUint8Array(data));
                }
                file.closeAsync();
                const combined = new Uint8Array(slices.reduce((sum, s) => sum + s.length, 0));
                let offset = 0;
                for (const slice of slices) { combined.set(slice, offset); offset += slice.length; }
                resolve(await JSZip.loadAsync(combined));
            } catch (err) {
                reject(err);
            }
        });
    });
}

async function detectSlideSize(isWeb) {
    const attempts = [];

    // (a) Office.js page setup — newer requirement sets only; not in our typings, so feature-detect.
    try {
        const r = await PowerPoint.run(async (ctx) => {
            const ps = ctx.presentation.pageSetup;
            if (!ps) return null;
            ps.load('slideWidth,slideHeight');
            await ctx.sync();
            return { rawW: ps.slideWidth, rawH: ps.slideHeight };
        });
        if (r && r.rawW && r.rawH) {
            attempts.push(`pageSetup: ${r.rawW}×${r.rawH}`);
            return { w: r.rawW / 72, h: r.rawH / 72, source: 'office.js pageSetup (assumed points)', attempts };
        }
        attempts.push('pageSetup: not available');
    } catch (err) {
        attempts.push(`pageSetup: ${err.message || err}`);
    }

    // (b) Desktop: read <p:sldSz> from ppt/presentation.xml.
    if (!isWeb) {
        try {
            const zip = await readDocumentZip();
            const xml = await zip.file('ppt/presentation.xml').async('string');
            const match = xml.match(/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
            if (match) {
                attempts.push(`presentation.xml: cx=${match[1]} cy=${match[2]}`);
                return { w: +match[1] / EMU_PER_IN, h: +match[2] / EMU_PER_IN, source: 'presentation.xml', attempts };
            }
            attempts.push('presentation.xml: no <p:sldSz>');
        } catch (err) {
            attempts.push(`presentation.xml: ${err.message || err}`);
        }
    } else {
        attempts.push('presentation.xml: skipped on web (getFileAsync Compressed unsupported)');
    }

    return { ...FALLBACK_SIZE, source: 'fallback (unknown)', attempts };
}

// ── Insert + inspect ─────────────────────────────────────────────────────────

async function insertDeck(base64) {
    return PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        const prevCount = slides.items.length;
        const lastSlideId = prevCount > 0 ? slides.items[prevCount - 1].id : undefined;
        ctx.presentation.insertSlidesFromBase64(base64, { formatting: 'UseDestinationTheme', targetSlideId: lastSlideId });
        await ctx.sync();
        return prevCount;
    });
}

const median = (xs) => {
    if (!xs.length) return null;
    const s = [...xs].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
};

async function inspectInserted(prevCount, deck) {
    return PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        const newSlides = slides.items.slice(prevCount);
        const collections = newSlides.map((slide) => {
            const shapes = slide.shapes;
            shapes.load('items/name,items/left,items/top,items/width,items/height,items/type');
            return shapes;
        });
        await ctx.sync();

        const actual = {};
        collections.forEach((c) => c.items.forEach((sh) => { actual[sh.name] = sh; }));

        // Geometry: expected (in) vs actual (pt). Table height excluded from max deviation — rows may grow.
        const missing = [];
        const deviations = [];
        const scaleW = [];
        const scaleH = [];
        const tables = [];
        for (const e of deck.expected) {
            const a = actual[e.name];
            if (!a) { missing.push(e.name); continue; }
            const ex = { left: e.x * 72, top: e.y * 72, width: e.w * 72, height: e.h * 72 };
            if (ex.width > 10) scaleW.push(a.width / ex.width);
            if (ex.height > 10 && e.kind !== 'table') scaleH.push(a.height / ex.height);
            const dev = Math.max(
                Math.abs(a.left - ex.left), Math.abs(a.top - ex.top), Math.abs(a.width - ex.width),
                e.kind === 'table' ? 0 : Math.abs(a.height - ex.height),
            );
            deviations.push({ name: e.name, devPt: +dev.toFixed(1) });
            if (e.kind === 'table') {
                tables.push({ name: e.name, estHeightPt: +ex.height.toFixed(1), actualHeightPt: +a.height.toFixed(1) });
            }
        }
        deviations.sort((a, b) => b.devPt - a.devPt);

        // Calibration: let PowerPoint size the boxes to their text, then read the real heights.
        const calibration = [];
        const calShapes = deck.calibration.map((c) => actual[c.name]).filter(Boolean);
        let calError = null;
        try {
            calShapes.forEach((sh) => { sh.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText'; });
            await ctx.sync();
            calShapes.forEach((sh) => sh.load('name,height'));
            await ctx.sync();
            for (const c of deck.calibration) {
                const sh = calShapes.find((x) => x.name === c.name);
                if (!sh) continue;
                const errPct = ((c.estHeightPt - sh.height) / sh.height) * 100;
                calibration.push({ ...c, actualHeightPt: +sh.height.toFixed(1), errPct: +errPct.toFixed(1) });
            }
        } catch (err) {
            calError = err.message || String(err);
        }

        return {
            newSlides: newSlides.length,
            geometry: {
                matched: deviations.length,
                missing,
                maxDevPt: deviations.length ? deviations[0].devPt : null,
                worst: deviations.slice(0, 5),
                medianScaleW: median(scaleW) && +median(scaleW).toFixed(3),
                medianScaleH: median(scaleH) && +median(scaleH).toFixed(3),
            },
            tables,
            calibration,
            calError,
        };
    });
}

// ── Entry point ──────────────────────────────────────────────────────────────

export async function runLayoutPoc({ variant = 'typical', sizeMode = 'match', isWeb = false, themeFonts = {} }) {
    const fixture = FIXTURES[variant];
    if (!fixture) throw new Error(`Unknown variant "${variant}". Use typical | max | overflow.`);
    if (!['match', 'default'].includes(sizeMode)) throw new Error(`Unknown size "${sizeMode}". Use match | default.`);

    const report = { variant, sizeMode, platform: isWeb ? 'web' : 'desktop', errors: [] };

    const detected = await detectSlideSize(isWeb);
    report.detected = detected;
    const built = sizeMode === 'match' ? { w: detected.w, h: detected.h } : { ...PPTXGEN_DEFAULT };
    report.built = { w: +built.w.toFixed(3), h: +built.h.toFixed(3) };

    const measurer = makeMeasurer({ heading: themeFonts.heading || 'Calibri Light', body: themeFonts.body || 'Calibri' });
    report.measuredWith = measurer.fonts;

    const t0 = performance.now();
    const deck = new Deck({ W: built.w, H: built.h, sizeMode, variant, measurer });
    drawDeckTitle(deck, fixture.title, fixture.subtitle);
    for (const slide of fixture.slides) {
        const draw = DRAWERS[slide.layout];
        if (!draw) { report.errors.push(`no drawer for layout "${slide.layout}"`); continue; }
        draw(deck, slide);
    }
    drawCalibration(deck);
    const base64 = await deck.pptx.write('base64');
    const t1 = performance.now();

    const prevCount = await insertDeck(base64);
    const t2 = performance.now();

    report.decisions = deck.decisions;
    report.timings = { buildMs: Math.round(t1 - t0), insertMs: Math.round(t2 - t1) };

    try {
        Object.assign(report, await inspectInserted(prevCount, deck));
    } catch (err) {
        report.errors.push(`inspect: ${err.message || err}`);
    }

    window.__pocLastReport = report;
    console.log('[POC] layout report', report);
    console.log('[POC] layout report (JSON)\n' + JSON.stringify(report, null, 2));
    return { report, summary: summarize(report) };
}

function summarize(r) {
    const cal = (r.calibration || []).map((c) => `${c.errPct > 0 ? '+' : ''}${c.errPct}%`).join(' ');
    const actions = (r.decisions || []).filter((d) => d.action !== 'fit').map((d) => `#${d.slide} ${d.layout}: ${d.action}`);
    return [
        `POC ${r.variant}/${r.sizeMode} on ${r.platform}: ${r.newSlides ?? '?'} slides inserted`,
        `Deck size ${r.detected.w.toFixed(2)}×${r.detected.h.toFixed(2)}in via ${r.detected.source}; built ${r.built.w}×${r.built.h}in`,
        `Build ${r.timings.buildMs}ms, insert ${r.timings.insertMs}ms`,
        r.geometry && `Geometry: ${r.geometry.matched} shapes, max dev ${r.geometry.maxDevPt}pt, scale ${r.geometry.medianScaleW}×${r.geometry.medianScaleH}, missing ${r.geometry.missing.length}`,
        cal && `Fit estimate error (est vs real height): ${cal}`,
        r.calError && `Calibration failed: ${r.calError}`,
        actions.length && `Responsive actions: ${actions.join('; ')}`,
        `Measured with ${r.measuredWith.heading} / ${r.measuredWith.body}`,
        r.errors.length && `Errors: ${r.errors.join('; ')}`,
        'Full report: console / window.__pocLastReport',
    ].filter(Boolean).join('\n');
}
