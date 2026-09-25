/**
 * Taskpane preview: a mini 16:9 slide per layout that looks like what `layouts.js` inserts
 * (not pixel-exact — it's HTML, readable at taskpane width). Colours come from the CSS variables
 * `applyCSSVariables` sets from the deck theme. AI text is only ever set via textContent.
 *
 * A card is { kind: 'title', title } or { kind: 'layout', slide: { layout, … } }.
 */

export const LAYOUT_LABELS = {
    title: 'Title',
    bullets: 'Bullets',
    pattern: 'Pattern',
    comparison: 'Comparison',
    'vocab-table': 'Table',
    examples: 'Examples',
};

export function cardLabel(card) {
    return card.kind === 'title' ? LAYOUT_LABELS.title : (LAYOUT_LABELS[card.slide.layout] || card.slide.layout);
}

function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
}

// Same rule as the inserted slide: first case-sensitive occurrence of `highlight` is emphasised.
function appendHighlighted(parent, text, highlight) {
    const i = highlight ? text.indexOf(highlight) : -1;
    if (i < 0) {
        parent.appendChild(document.createTextNode(text));
        return parent;
    }
    if (i > 0) parent.appendChild(document.createTextNode(text.slice(0, i)));
    parent.appendChild(el('strong', 'sm-hl', highlight));
    if (i + highlight.length < text.length) parent.appendChild(document.createTextNode(text.slice(i + highlight.length)));
    return parent;
}

function renderBullets(body, slide) {
    const list = el('ul', 'sm-bullets');
    (slide.items || []).forEach((text) => list.appendChild(el('li', null, text)));
    body.appendChild(list);
}

function renderPattern(body, slide) {
    const chips = el('div', 'sm-chips');
    (slide.parts || []).forEach((part, i) => {
        if (i > 0) chips.appendChild(el('span', 'sm-arrow', '→'));
        chips.appendChild(el('span', 'sm-chip', part));
    });
    body.appendChild(chips);
    if (slide.example?.text) {
        body.appendChild(appendHighlighted(el('div', 'sm-example'), slide.example.text, slide.example.highlight));
    }
}

function renderComparison(body, slide) {
    const wrap = el('div', 'sm-compare');
    [slide.left, slide.right].forEach((side, c) => {
        if (!side) return;
        const polarity = side.polarity === 'positive' || side.polarity === 'negative' ? side.polarity : `neutral-${c}`;
        const col = el('div', `sm-side sm-side--${polarity}`);
        const mark = side.polarity === 'positive' ? '✓ ' : side.polarity === 'negative' ? '✗ ' : '';
        col.appendChild(el('div', 'sm-side-head', `${mark}${side.heading || ''}`));
        const list = el('ul', 'sm-side-items');
        (side.items || []).forEach((it) => list.appendChild(appendHighlighted(el('li'), it.text || '', it.highlight)));
        col.appendChild(list);
        wrap.appendChild(col);
    });
    body.appendChild(wrap);
}

function renderTable(body, slide) {
    const table = el('table', 'sm-table');
    const head = el('tr');
    (slide.columns || []).forEach((c) => head.appendChild(el('th', null, c)));
    table.appendChild(el('thead')).appendChild(head);
    const tbody = el('tbody');
    (slide.rows || []).forEach((row) => {
        const tr = el('tr');
        (slide.columns || []).forEach((_, c) => tr.appendChild(el('td', null, row[c] ?? '')));
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    body.appendChild(table);
}

function renderExamples(body, slide) {
    const list = el('div', 'sm-examples');
    (slide.items || []).forEach((it) => {
        const callout = el('div', 'sm-callout');
        callout.appendChild(appendHighlighted(el('div'), it.text || '', it.highlight));
        if (it.translation) callout.appendChild(el('div', 'sm-translation', it.translation));
        list.appendChild(callout);
    });
    body.appendChild(list);
}

const RENDERERS = {
    bullets: renderBullets,
    pattern: renderPattern,
    comparison: renderComparison,
    'vocab-table': renderTable,
    examples: renderExamples,
};

export function renderSlideCard(container, card) {
    container.replaceChildren();

    if (card.kind === 'title') {
        const slide = el('div', 'slide-mini slide-mini--title');
        slide.appendChild(el('div', 'sm-band'));
        slide.appendChild(el('div', 'sm-deck-title', card.title));
        container.appendChild(slide);
        return;
    }

    const data = card.slide;
    const slide = el('div', `slide-mini slide-mini--${data.layout}`);
    slide.appendChild(el('div', 'sm-title', data['slide-title'] || ''));
    slide.appendChild(el('div', 'sm-title-bar'));
    const body = el('div', 'sm-body');
    const render = RENDERERS[data.layout];
    if (render) render(body, data);
    else body.appendChild(el('div', 'sm-unknown', `Unsupported layout: ${data.layout}`));
    slide.appendChild(body);
    container.appendChild(slide);
}
