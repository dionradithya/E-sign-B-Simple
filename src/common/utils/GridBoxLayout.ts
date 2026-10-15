import { IGridBoxCell } from '../models/IEsignState';

/**
 * Shared look of the requestor grid box (approval stamp).
 * Used by both the on-screen preview (PdfViewer) and the PDF output (PdfUtils.drawGridBoxes),
 * so what the requestor sees is what gets burned into the document.
 */
export const GRID_BOX_COLORS = {
    band: '#D6E6F2',       // Light blue header band
    line: '#1F6491',       // Blue lines
    text: '#1F6491',       // Blue labels
    title: '#D4151C'       // Red title
};

export const DEFAULT_GRID_BOX_TITLE = 'Received From PIC';

export const DEFAULT_GRID_BOX_CELLS: IGridBoxCell[] = [
    { header: 'Validation By', footer: 'Treasury' },
    { header: 'Checked By', footer: 'Budget Division' },
    { header: 'Checked By', footer: 'Tax Division' },
    { header: 'Approved By', footer: 'GH Finance' }
];

/**
 * Keeps existing cell labels when rows/cols change; new cells start empty.
 */
export const resizeGridBoxCells = (cells: IGridBoxCell[], rows: number, cols: number): IGridBoxCell[] =>
    Array.from({ length: rows * cols }, (_, i) => cells[i] || { header: '', footer: '' });

export interface IGridBoxRowLayout {
    top: number;
    headerH: number;
    bottom: number;
}

/**
 * All values are in the same unit as w/h (px on screen, points in the PDF),
 * measured from the box's top-left corner.
 */
export interface IGridBoxLayout {
    lineWidth: number;
    padX: number;
    titleH: number;
    titleSize: number;
    titleBaseline: number;
    /** Baseline of the typed dots in the title area */
    dotY: number;
    /** Fill-in segments [x1, x2]: left, middle (after the colon) and right (under the title) */
    dotLines: [number, number][];
    colonX: number;
    rows: IGridBoxRowLayout[];
    colW: number;
    textSize: number;
}

export const computeGridBoxLayout = (w: number, h: number, rows: number, cols: number): IGridBoxLayout => {
    const titleH = h * 0.16;
    const rowH = (h - titleH) / rows;
    const headerH = Math.min(rowH * 0.15, h * 0.07);

    return {
        lineWidth: Math.max(0.75, h * 0.004),
        padX: w * 0.03,
        titleH,
        titleSize: titleH * 0.32,
        titleBaseline: titleH * 0.36,
        dotY: titleH * 0.86,
        dotLines: [[w * 0.03, w * 0.25], [w * 0.29, w * 0.49], [w * 0.55, w * 0.97]],
        colonX: w * 0.27,
        rows: Array.from({ length: rows }, (_, r) => ({
            top: titleH + rowH * r,
            headerH,
            bottom: titleH + rowH * (r + 1)
        })),
        colW: w / cols,
        textSize: headerH * 0.75
    };
};

/**
 * Shrinks the font size so the text fits in maxWidth.
 * measure(text, size) returns the text width at that size.
 */
export const fitTextSize = (text: string, maxWidth: number, size: number, measure: (text: string, size: number) => number): number => {
    const width = measure(text, size);
    return width > maxWidth && width > 0 ? size * (maxWidth / width) : size;
};

/**
 * Typed dots ("......") that fill the given width, like a printed fill-in line.
 */
export const dotLeader = (width: number, size: number, measure: (text: string, size: number) => number): string => {
    const dotWidth = measure('.', size);
    return dotWidth > 0 ? '.'.repeat(Math.max(0, Math.floor(width / dotWidth))) : '';
};

/**
 * Standard PDF fonts (Helvetica) only support WinAnsi characters; replace anything else so drawing never fails.
 */
export const toWinAnsiText = (text: string): string => text.replace(/[^\n\x20-\x7E\xA0-\xFF]/g, '?');
