/**
 * Text measurement for slide fitting. Neither Office.js nor PptxGenJS can measure text, so we
 * estimate wrapped height with a hidden canvas: lines × font size × line height + margins — the
 * same formula PowerPoint uses. The estimate is exact when the predicted line count is right
 * (POC 1 calibration, doc/poc-1-rendering-results.md); when it is wrong it is off by a whole line,
 * so words are wrapped against a slightly narrower width (FIT.widthSafety) to err on the safe side.
 */

export const FIT = {
    // Font sizes (pt): start = preferred, floor = smallest readable from the back of a classroom.
    SIZE: {
        deckTitle: { start: 48, floor: 32 },
        title: { start: 36, floor: 28 },
        body: { start: 24, floor: 18 },
        table: { start: 20, floor: 16 },
        chip: { start: 24, floor: 16 },
    },
    TEXT_MARGIN_PT: 5,          // inset we set on every text box (all sides)
    LINE_HEIGHT: 1.2,           // single spacing ≈ 1.2 × font size
    PARA_AFTER_PT: 6,
    BULLET_INDENT_PT: 27,       // PptxGenJS DEF_BULLET_MARGIN
    CELL_MARGIN_IN: { x: 0.1, y: 0.05 }, // PptxGenJS DEF_CELL_MARGIN_IN
    MIN_COLUMN_IN: 3.5,         // narrower than this, reflow to 2 columns is not allowed
    widthSafety: 0.96,          // wrap against 96% of the real width — borderline words wrap early in our estimate
};

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

// themeFonts: { heading, body } — the deck's theme fonts (desktop reads them from the file).
export function makeMeasurer(themeFonts) {
    const resolve = (name) => {
        if (fontAvailable(name)) return { css: `"${name}"`, used: name };
        if (fontAvailable('Calibri')) return { css: '"Calibri"', used: `Calibri (fallback for ${name})` };
        return { css: 'Arial, sans-serif', used: `Arial (fallback for ${name})` };
    };
    const family = {
        heading: resolve(themeFonts.heading || 'Calibri Light'),
        body: resolve(themeFonts.body || 'Calibri'),
    };

    function width(text, role, pt, bold) {
        measureCtx.font = `${bold ? 'bold ' : ''}${pt}px ${family[role].css}`;
        return measureCtx.measureText(text).width; // px at `pt`px == points at `pt`pt
    }

    // Greedy word wrap, like PowerPoint; a single word wider than the line is broken by characters.
    function lineCount(text, role, pt, maxWidthPt, bold) {
        const limit = maxWidthPt * FIT.widthSafety;
        const fits = (s) => width(s, role, pt, bold) <= limit;
        let total = 0;
        for (const para of String(text ?? '').split('\n')) {
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
            return lineCount(p.text, role, size, widthPt, p.bold ?? bold) * size * FIT.LINE_HEIGHT;
        });
        return lines.reduce((a, b) => a + b, 0) + FIT.PARA_AFTER_PT * Math.max(0, paras.length - 1);
    }

    return { width, lineCount, heightPt, fonts: { heading: family.heading.used, body: family.body.used } };
}

// Try sizes from start down to floor; returns the first that fits, or the floor with fits=false.
export function fitParas(m, paras, box, { role = 'body', size = FIT.SIZE.body, bold = false, indentPt = 0 } = {}) {
    const widthPt = box.w * 72 - 2 * FIT.TEXT_MARGIN_PT - indentPt;
    const availPt = box.h * 72 - 2 * FIT.TEXT_MARGIN_PT;
    for (let pt = size.start; pt >= size.floor; pt -= 2) {
        const est = m.heightPt(paras, role, pt, widthPt, bold);
        if (est <= availPt) return { pt, fits: true, estPt: est, availPt };
    }
    const est = m.heightPt(paras, role, size.floor, widthPt, bold);
    return { pt: size.floor, fits: false, estPt: est, availPt };
}
