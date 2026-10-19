import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { IGridBox } from '../../../common/models/IEsignState';
import { GRID_BOX_COLORS, computeGridBoxLayout, dotLeader, fitTextSize } from '../../../common/utils/GridBoxLayout';

// Helvetica is what PdfUtils embeds; Arial has the same metrics as a fallback
const FONT_FAMILY = 'Helvetica, Arial, sans-serif';

let measureCanvas: HTMLCanvasElement | undefined;
const measureBold = (text: string, size: number): number => {
    if (!measureCanvas) measureCanvas = document.createElement('canvas');
    const ctx = measureCanvas.getContext('2d');
    if (!ctx) return text.length * size * 0.6;
    ctx.font = `bold ${size}px ${FONT_FAMILY}`;
    return ctx.measureText(text).width;
};

/**
 * On-screen preview of a grid box. Uses the same layout as PdfUtils.drawGridBoxes,
 * so the preview matches what is burned into the PDF.
 */
const GridBoxPreview: React.FC<{ box: IGridBox }> = ({ box }) => {
    const ref = useRef<HTMLDivElement>(null);
    const [size, setSize] = useState<{ w: number; h: number }>({ w: 0, h: 0 });

    useEffect(() => {
        if (!ref.current) return;
        const observer = new ResizeObserver(entries => {
            const { width, height } = entries[0].contentRect;
            setSize({ w: width, h: height });
        });
        observer.observe(ref.current);
        return () => observer.disconnect();
    }, []);

    const { w, h } = size;
    const L = w > 0 && h > 0 ? computeGridBoxLayout(w, h, box.rows, box.cols) : undefined;

    const textEl = (key: string, str: string, x: number, baseline: number, fontSize: number, color: string): React.ReactElement => (
        <text key={key} x={x} y={baseline} fontSize={fontSize} fontFamily={FONT_FAMILY} fontWeight={700} fill={color}>{str}</text>
    );

    const centeredText = (key: string, str: string, left: number, width: number, baseline: number): React.ReactElement | null => {
        if (!L) return null;
        const clean = str.trim();
        if (!clean) return null;
        const fontSize = fitTextSize(clean, width - L.padX * 2, L.textSize, measureBold);
        return textEl(key, clean, left + (width - measureBold(clean, fontSize)) / 2, baseline, fontSize, GRID_BOX_COLORS.text);
    };

    const title = box.title.trim();
    const titleSize = L && title ? fitTextSize(title, w * 0.6, L.titleSize, measureBold) : 0;

    return (
        <div ref={ref} style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, pointerEvents: 'none' }}>
            {L && (
                <svg width="100%" height="100%" viewBox={`0 0 ${w} ${h}`} style={{ display: 'block', overflow: 'visible' }}>
                    {/* Header bands (no background fill, same as the PDF) */}
                    {L.rows.map((r, i) => <rect key={`band-${i}`} x={0} y={r.top} width={w} height={r.headerH} fill={GRID_BOX_COLORS.band} />)}

                    <g stroke={GRID_BOX_COLORS.line} strokeWidth={L.lineWidth}>
                        {/* Frame */}
                        <line x1={0} y1={0} x2={w} y2={0} />
                        <line x1={0} y1={0} x2={0} y2={h} />
                        <line x1={w} y1={0} x2={w} y2={h} />
                        <line x1={0} y1={h} x2={w} y2={h} />

                        {/* Header band borders & column dividers */}
                        {L.rows.map((r, i) => (
                            <g key={`row-${i}`}>
                                <line x1={0} y1={r.top} x2={w} y2={r.top} />
                                <line x1={0} y1={r.top + r.headerH} x2={w} y2={r.top + r.headerH} />
                                {Array.from({ length: box.cols - 1 }, (_, c) => (
                                    <line key={c} x1={L.colW * (c + 1)} y1={r.top + r.headerH} x2={L.colW * (c + 1)} y2={r.bottom} />
                                ))}
                            </g>
                        ))}

                    </g>

                    {/* Typed dots with optional fill-in values */}
                    {L.dotLines.map(([x1, x2], i) => {
                        const value = (box.fillIns?.[i] || '').trim();
                        const valueSize = value ? fitTextSize(value, x2 - x1, L.textSize * 0.9, measureBold) : 0;
                        return (
                            <g key={`fill-${i}`}>
                                {textEl('dots', dotLeader(x2 - x1, L.textSize, measureBold), x1, L.dotY, L.textSize, GRID_BOX_COLORS.text)}
                                {value && textEl('value', value, x1 + (x2 - x1 - measureBold(value, valueSize)) / 2, L.dotY - L.textSize * 0.3, valueSize, GRID_BOX_COLORS.text)}
                            </g>
                        );
                    })}
                    {textEl('colon', ':', L.colonX - measureBold(':', L.textSize) / 2, L.dotY, L.textSize, GRID_BOX_COLORS.text)}
                    {title && textEl('title', title, w - L.padX - measureBold(title, titleSize), L.titleBaseline, titleSize, GRID_BOX_COLORS.title)}

                    {/* Cell labels */}
                    {L.rows.map((r, ri) => Array.from({ length: box.cols }, (_, c) => {
                        const cell = box.cells[ri * box.cols + c];
                        if (!cell) return null;
                        return (
                            <g key={`cell-${ri}-${c}`}>
                                {centeredText('h', cell.header, L.colW * c, L.colW, r.top + r.headerH * 0.72)}
                                {centeredText('f', cell.footer, L.colW * c, L.colW, r.bottom - r.headerH * 0.35)}
                            </g>
                        );
                    }))}
                </svg>
            )}
        </div>
    );
};

export default GridBoxPreview;
