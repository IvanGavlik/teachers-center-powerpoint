/* global PowerPoint */

import { makeMeasurer } from './measure';
import { Deck, DRAWERS, drawDeckTitle } from './layouts';

/**
 * Build a .pptx (base64) for a backend deck { title, slides: [{ layout, … }] } at the deck's real
 * slide size. Same path on desktop and web.
 *
 * Returns { base64, slideCount, splits, skipped } — slideCount is what will actually be inserted
 * (continuation slides included), splits = AI slides that needed more than one PowerPoint slide.
 */
export async function buildDeckBase64(deckData, size, themeFonts) {
    const deck = new Deck({ W: size.w, H: size.h, measurer: makeMeasurer(themeFonts || {}) });
    let splits = 0;
    let skipped = 0;

    if (deckData.title) drawDeckTitle(deck, deckData.title);

    for (const slide of deckData.slides || []) {
        const draw = DRAWERS[slide.layout];
        if (!draw) {
            console.warn(`[Slides] No drawer for layout "${slide.layout}" — slide skipped`, slide);
            skipped++;
            continue;
        }
        const before = deck.slideCount;
        draw(deck, slide);
        if (deck.slideCount - before > 1) splits++;
    }

    const base64 = await deck.pptx.write('base64');
    return { base64, slideCount: deck.slideCount, splits, skipped };
}

// Insert after the last slide, styled by the destination deck's theme.
export async function insertDeckBase64(base64) {
    await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load('items/id');
        await ctx.sync();
        const lastSlideId = slides.items.length > 0 ? slides.items[slides.items.length - 1].id : undefined;
        ctx.presentation.insertSlidesFromBase64(base64, { formatting: 'UseDestinationTheme', targetSlideId: lastSlideId });
        await ctx.sync();
    });
}
