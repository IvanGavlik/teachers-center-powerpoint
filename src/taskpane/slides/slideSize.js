/* global PowerPoint, Office */

import JSZip from 'jszip';

const EMU_PER_IN = 914400;
const FALLBACK_SIZE = { w: 13.333, h: 7.5 }; // a new PowerPoint deck (16:9)

// Office JS returns slice data as a plain number array on Win32, ArrayBuffer elsewhere.
export function toUint8Array(data) {
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

// Reads the open presentation as a ZIP. Desktop only — getFileAsync(Compressed) is not
// supported in PowerPoint for the web.
export function readDocumentZip() {
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

/**
 * The deck's slide size in inches: { w, h, source }.
 * Office.js pageSetup works on desktop and web (POC 1); presentation.xml (desktop) and a 16:9
 * default are kept as a safety net for older Office builds.
 */
export async function detectSlideSize(isWeb) {
    try {
        const r = await PowerPoint.run(async (ctx) => {
            const ps = ctx.presentation.pageSetup;
            if (!ps) return null;
            ps.load('slideWidth,slideHeight');
            await ctx.sync();
            return { rawW: ps.slideWidth, rawH: ps.slideHeight };
        });
        if (r && r.rawW && r.rawH) return { w: r.rawW / 72, h: r.rawH / 72, source: 'pageSetup' };
    } catch (err) {
        console.warn('[SlideSize] pageSetup unavailable:', err.message || err);
    }

    if (!isWeb) {
        try {
            const zip = await readDocumentZip();
            const xml = await zip.file('ppt/presentation.xml').async('string');
            const match = xml.match(/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
            if (match) return { w: +match[1] / EMU_PER_IN, h: +match[2] / EMU_PER_IN, source: 'presentation.xml' };
        } catch (err) {
            console.warn('[SlideSize] presentation.xml unreadable:', err.message || err);
        }
    }

    return { ...FALLBACK_SIZE, source: 'fallback' };
}

// Layout hint for the backend: {"aspect":"16:9","width-in":13.33,"height-in":7.5}
export function toSlideFormat({ w, h }) {
    const ratio = w / h;
    const aspect = Math.abs(ratio - 16 / 9) < 0.02 ? '16:9'
        : Math.abs(ratio - 4 / 3) < 0.02 ? '4:3'
        : `${+w.toFixed(2)}:${+h.toFixed(2)}`;
    return { aspect, 'width-in': +w.toFixed(2), 'height-in': +h.toFixed(2) };
}
