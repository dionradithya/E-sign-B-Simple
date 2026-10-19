import * as React from 'react';
import { useState, useEffect } from 'react';
import { Document, Page, pdfjs } from 'react-pdf';
import 'react-pdf/dist/Page/AnnotationLayer.css';
import 'react-pdf/dist/Page/TextLayer.css';
import { sanitizeApproverName, formatBadgeNumber } from '../../../common/utils/helper';
import { useTouchDrag } from './TouchDragContext';
import { IGridBox, ITextAnnotation } from '../../../common/models/IEsignState';
import { FREE_TEXT_LINE_HEIGHT } from '../../../common/utils/PdfUtils';
import GridBoxPreview from './GridBoxPreview';

// Worker Setup - Create Blob URL from bundled worker content
// eslint-disable-next-line @typescript-eslint/no-var-requires
const workerContent = require('../../../common/assets/pdf.worker.min.js');
const workerBlob = new Blob([typeof workerContent === 'string' ? workerContent : workerContent.default || ''], { type: 'application/javascript' });
pdfjs.GlobalWorkerOptions.workerSrc = URL.createObjectURL(workerBlob);

export interface ISignaturePlaceholder {
  id: string; // unique internal ID
  page: number;
  x: number;
  y: number;
  approverId: number;
  checklistName: boolean;
  checklistDate: boolean;
  checklistBadge: boolean;
  type: 'initial' | 'signature';
  width: number;
  height: number;
  groupId?: string;
}

/** Position and size (in %) shared by signature placeholders and grid boxes */
type IBoxRect = Pick<ISignaturePlaceholder, 'x' | 'y' | 'width' | 'height'> & { size?: number };

// Free text size limits (% of page width), ~4pt to ~96pt on A4 portrait
const MIN_TEXT_SIZE = (4 / 595) * 100;
const MAX_TEXT_SIZE = (96 / 595) * 100;

type InteractionKind = 'placeholder' | 'grid' | 'text';

interface IPdfViewerProps {
  fileUrl: string;
  placeholders: ISignaturePlaceholder[];
  approvers: { id: number; displayName: string; jobTitle?: string; badgeNumber?: string }[];
  onDrop: (page: number, x: number, y: number, approverId: number, type: 'initial' | 'signature', checklistName: boolean, checklistDate: boolean, checklistBadge: boolean, widthPercent?: number, heightPercent?: number) => void;
  onUpdatePlaceholder: (id: string, updates: Partial<ISignaturePlaceholder>) => void;
  onRemovePlaceholder: (id: string) => void;
  gridBoxes: IGridBox[];
  onUpdateGridBox: (id: string, updates: Partial<IGridBox>) => void;
  onRemoveGridBox: (id: string) => void;
  texts: ITextAnnotation[];
  onUpdateText: (id: string, updates: Partial<ITextAnnotation>) => void;
  onRemoveText: (id: string) => void;
  /** Grid box or text currently being edited in the sidebar */
  selectedId: string | null;
  onSelect: (id: string | null) => void;
  onDocumentLoaded?: (numPages: number) => void;
  /** Called with the page that takes up most of the screen while the user scrolls */
  onVisiblePageChange?: (page: number) => void;
}

const PdfViewer: React.FC<IPdfViewerProps> = ({
  fileUrl,
  placeholders,
  approvers,
  onDrop,
  onUpdatePlaceholder,
  onRemovePlaceholder,
  gridBoxes,
  onUpdateGridBox,
  onRemoveGridBox,
  texts,
  onUpdateText,
  onRemoveText,
  selectedId,
  onSelect,
  onDocumentLoaded,
  onVisiblePageChange
}) => {
  const [numPages, setNumPages] = useState<number | null>(null);
  const [pdfWrapperWidth, setPdfWrapperWidth] = useState<number>(600); // Default start
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const pageRefs = React.useRef<Map<number, HTMLDivElement>>(new Map());

  // Track which page is most visible on screen (used by "Current Page" placement).
  // root: null measures against the viewport, clipped by the scroll container.
  useEffect(() => {
    if (!numPages || !onVisiblePageChange || typeof IntersectionObserver === 'undefined') return;

    const visibleHeights = new Map<number, number>();
    let lastPage = 0;
    const thresholds = Array.from({ length: 21 }, (_, i) => i / 20);

    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        const page = Number((entry.target as HTMLElement).dataset.page);
        visibleHeights.set(page, entry.isIntersecting ? entry.intersectionRect.height : 0);
      });

      let bestPage = lastPage || 1;
      let bestHeight = -1;
      visibleHeights.forEach((height, page) => {
        if (height > bestHeight) {
          bestHeight = height;
          bestPage = page;
        }
      });

      if (bestPage !== lastPage) {
        lastPage = bestPage;
        onVisiblePageChange(bestPage);
      }
    }, { root: null, threshold: thresholds });

    pageRefs.current.forEach(el => observer.observe(el));
    return () => observer.disconnect();
  }, [numPages, pdfWrapperWidth, onVisiblePageChange]);

  // Interaction State
  const [interactingId, setInteractingId] = useState<string | null>(null);
  const [interactionType, setInteractionType] = useState<'move' | 'resize' | null>(null);
  const [interactingKind, setInteractingKind] = useState<InteractionKind>('placeholder');
  // Store initial state for delta calculation
  const [startState, setStartState] = useState<{ xPx: number, yPx: number, wPx: number, hPx: number, mouseX: number, mouseY: number, parentW: number, parentH: number, size?: number } | null>(null);

  // Touch drag context for mobile support
  const { dragData, setDragData, isDragging, setIsDragging } = useTouchDrag();

  // Responsive Width Observer
  useEffect(() => {
    if (!wrapperRef.current) return;
    const resizeObserver = new ResizeObserver((entries) => {
      for (const entry of entries) {
        if (entry.contentRect.width) {
          // Full width, practically no buffer needed if padding is removed
          const width = entry.contentRect.width - 10; // 20px buffer just for scrollbar safety
          setPdfWrapperWidth(width > 200 ? width : 200);
        }
      }
    });

    resizeObserver.observe(wrapperRef.current);
    return () => resizeObserver.disconnect();
  }, []);

  useEffect(() => {
    // Shared move/resize logic for mouse and touch, for both signature placeholders and grid boxes
    const applyDelta = (deltaX: number, deltaY: number): void => {
      if (!interactingId || !startState || !interactionType) return;

      const onUpdate = interactingKind === 'grid' ? onUpdateGridBox : (interactingKind === 'text' ? onUpdateText : onUpdatePlaceholder);

      if (interactionType === 'move') {
        let newXPx = startState.xPx + deltaX;
        let newYPx = startState.yPx + deltaY;

        // Clamp positions in Px
        newXPx = Math.max(0, Math.min(newXPx, startState.parentW - startState.wPx));
        newYPx = Math.max(0, Math.min(newYPx, startState.parentH - startState.hPx));

        // Convert back to %
        onUpdate(interactingId, {
          x: (newXPx / startState.parentW) * 100,
          y: (newYPx / startState.parentH) * 100
        });
        return;
      }

      // Calculate max available dimensions in Px
      const maxWPx = startState.parentW - startState.xPx;
      const maxHPx = startState.parentH - startState.yPx;

      if (interactingKind === 'text') {
        // Text resizes by scaling its font size with the dragged width
        if (!startState.size || startState.wPx <= 0) return;
        const newWidthPx = Math.min(Math.max(10, startState.wPx + deltaX), maxWPx);
        const newSize = startState.size * (newWidthPx / startState.wPx);

        onUpdateText(interactingId, { size: Math.min(MAX_TEXT_SIZE, Math.max(MIN_TEXT_SIZE, newSize)) });
        return;
      }

      if (interactingKind === 'grid') {
        // Grid boxes resize freely (no fixed aspect ratio)
        const newWidthPx = Math.min(Math.max(30, startState.wPx + deltaX), maxWPx);
        const newHeightPx = Math.min(Math.max(20, startState.hPx + deltaY), maxHPx);

        onUpdateGridBox(interactingId, {
          width: (newWidthPx / startState.parentW) * 100,
          height: (newHeightPx / startState.parentH) * 100
        });
        return;
      }

      const rawNewWidthPx = Math.max(50, startState.wPx + deltaX);

      // Constrain width by available width AND available height (taking 3:2 ratio)
      // Height = Width / 1.5
      // Width <= maxHPx * 1.5
      const constrainedWidthPx = Math.min(rawNewWidthPx, maxWPx, maxHPx * 1.5);
      const newHeightPx = constrainedWidthPx / 1.5;

      onUpdatePlaceholder(interactingId, {
        width: (constrainedWidthPx / startState.parentW) * 100,
        height: (newHeightPx / startState.parentH) * 100
      });
    };

    const handleWindowMouseMove = (e: MouseEvent): void => {
      if (!interactingId || !startState) return;
      applyDelta(e.clientX - startState.mouseX, e.clientY - startState.mouseY);
    };

    // Touch move handler for mobile
    const handleWindowTouchMove = (e: TouchEvent): void => {
      if (!interactingId || !startState) return;

      e.preventDefault(); // Prevent scrolling while moving/resizing
      const touch = e.touches[0];
      applyDelta(touch.clientX - startState.mouseX, touch.clientY - startState.mouseY);
    };

    const handleWindowMouseUp = (): void => {
      setInteractingId(null);
      setInteractionType(null);
      setStartState(null);
    };

    // Touch end handler for mobile
    const handleWindowTouchEnd = (): void => {
      setInteractingId(null);
      setInteractionType(null);
      setStartState(null);
    };

    if (interactingId) {
      window.addEventListener('mousemove', handleWindowMouseMove);
      window.addEventListener('mouseup', handleWindowMouseUp);
      window.addEventListener('touchmove', handleWindowTouchMove, { passive: false });
      window.addEventListener('touchend', handleWindowTouchEnd);
    }

    return () => {
      window.removeEventListener('mousemove', handleWindowMouseMove);
      window.removeEventListener('mouseup', handleWindowMouseUp);
      window.removeEventListener('touchmove', handleWindowTouchMove);
      window.removeEventListener('touchend', handleWindowTouchEnd);
    };
  }, [interactingId, interactingKind, interactionType, startState, onUpdatePlaceholder, onUpdateGridBox, onUpdateText]);


  const handleMouseDown = (e: React.MouseEvent, id: string, type: 'move' | 'resize', placeholder: IBoxRect, kind: InteractionKind = 'placeholder'): void => {
    e.stopPropagation(); // Prevent drag drop interference
    e.preventDefault();  // Prevent text selection or native drag

    let parentElement = e.currentTarget.parentElement;

    // If resizing, the handle is the target, so parent is the box. 
    // We need the box's parent (the page container).
    if (type === 'resize' && parentElement) {
      parentElement = parentElement.parentElement;
    }

    if (!parentElement) return;

    const parentW = parentElement.offsetWidth;
    const parentH = parentElement.offsetHeight;

    // Convert stored % to Px for startState interaction
    const xPx = (placeholder.x / 100) * parentW;
    const yPx = (placeholder.y / 100) * parentH;
    const wPx = (placeholder.width / 100) * parentW;
    const hPx = (placeholder.height / 100) * parentH;

    setInteractingId(id);
    setInteractingKind(kind);
    if (kind !== 'placeholder') onSelect(id);
    setInteractionType(type);
    setStartState({
      xPx,
      yPx,
      wPx,
      hPx,
      mouseX: e.clientX,
      mouseY: e.clientY,
      parentW,
      parentH,
      size: placeholder.size
    });
  };

  // Touch start handler for mobile move/resize
  const handleTouchStart = (e: React.TouchEvent, id: string, type: 'move' | 'resize', placeholder: IBoxRect, kind: InteractionKind = 'placeholder'): void => {
    e.stopPropagation();
    // Don't preventDefault here to allow delete button to work

    const touch = e.touches[0];
    let parentElement = e.currentTarget.parentElement;

    if (type === 'resize' && parentElement) {
      parentElement = parentElement.parentElement;
    }

    if (!parentElement) return;

    const parentW = parentElement.offsetWidth;
    const parentH = parentElement.offsetHeight;

    const xPx = (placeholder.x / 100) * parentW;
    const yPx = (placeholder.y / 100) * parentH;
    const wPx = (placeholder.width / 100) * parentW;
    const hPx = (placeholder.height / 100) * parentH;

    setInteractingId(id);
    setInteractingKind(kind);
    if (kind !== 'placeholder') onSelect(id);
    setInteractionType(type);
    setStartState({
      xPx,
      yPx,
      wPx,
      hPx,
      mouseX: touch.clientX,
      mouseY: touch.clientY,
      parentW,
      parentH,
      size: placeholder.size
    });
  };

  const getElementRect = (el: HTMLElement, t: ITextAnnotation): IBoxRect => {
    const parent = el.parentElement;
    const parentW = parent?.offsetWidth || 1;
    const parentH = parent?.offsetHeight || 1;
    return { x: t.x, y: t.y, width: (el.offsetWidth / parentW) * 100, height: (el.offsetHeight / parentH) * 100, size: t.size };
  };

  function onDocumentLoadSuccess({ numPages }: { numPages: number }): void {
    console.log("PDF LOADED:", numPages);
    setNumPages(numPages);
    onDocumentLoaded?.(numPages);
  }

  const handleDrop = (e: React.DragEvent<HTMLDivElement>, pageNumber: number): void => {
    e.preventDefault();
    const type = e.dataTransfer.getData("type") as 'initial' | 'signature';
    if (!type || (type !== "initial" && type !== "signature")) return;

    const approverId = parseInt(e.dataTransfer.getData("approverId"));
    if (!approverId) return;

    const checklistName = e.dataTransfer.getData("checklistName") === "true";
    const checklistDate = e.dataTransfer.getData("checklistDate") === "true";
    const checklistBadge = e.dataTransfer.getData("checklistBadge") === "true";

    const rect = e.currentTarget.getBoundingClientRect();
    const containerWidth = rect.width;
    const containerHeight = rect.height;

    // Standard initial size in PX (e.g. 150x100)
    // We convert this to % so its relative size is preserved on this screen width
    const targetWidthPx = 150;
    const targetHeightPx = 100;

    // Relative mouse position in PX
    const mouseX = e.clientX - rect.left;
    const mouseY = e.clientY - rect.top;

    // Centered position in PX
    let xPx = mouseX - (targetWidthPx / 2);
    let yPx = mouseY - (targetHeightPx / 2);

    // Clamp in PX
    xPx = Math.max(0, Math.min(xPx, containerWidth - targetWidthPx));
    yPx = Math.max(0, Math.min(yPx, containerHeight - targetHeightPx));

    // Convert to %
    const xPercent = (xPx / containerWidth) * 100;
    const yPercent = (yPx / containerHeight) * 100;
    const wPercent = (targetWidthPx / containerWidth) * 100;
    const hPercent = (targetHeightPx / containerHeight) * 100;

    onDrop(pageNumber, xPercent, yPercent, approverId, type, checklistName, checklistDate, checklistBadge, wPercent, hPercent);
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  };

  // Handle touch tap on PDF page for mobile - places signature box when isDragging is active
  const handleTouchTap = (e: React.TouchEvent<HTMLDivElement>, pageNumber: number): void => {
    if (!isDragging || !dragData) return;

    const touch = e.changedTouches[0];
    const rect = e.currentTarget.getBoundingClientRect();
    const containerWidth = rect.width;
    const containerHeight = rect.height;

    // Standard initial size in PX (e.g. 150x100)
    const targetWidthPx = 150;
    const targetHeightPx = 100;

    // Relative touch position in PX
    const touchX = touch.clientX - rect.left;
    const touchY = touch.clientY - rect.top;

    // Centered position in PX
    let xPx = touchX - (targetWidthPx / 2);
    let yPx = touchY - (targetHeightPx / 2);

    // Clamp in PX
    xPx = Math.max(0, Math.min(xPx, containerWidth - targetWidthPx));
    yPx = Math.max(0, Math.min(yPx, containerHeight - targetHeightPx));

    // Convert to %
    const xPercent = (xPx / containerWidth) * 100;
    const yPercent = (yPx / containerHeight) * 100;
    const wPercent = (targetWidthPx / containerWidth) * 100;
    const hPercent = (targetHeightPx / containerHeight) * 100;

    onDrop(pageNumber, xPercent, yPercent, dragData.approverId, dragData.type, dragData.checklistName, dragData.checklistDate, dragData.checklistBadge, wPercent, hPercent);

    // Reset touch drag state
    setDragData(undefined);
    setIsDragging(false);
  };

  const getApproverName = (id: number): string => {
    const app = approvers.find(a => a.id === id);
    return app ? sanitizeApproverName(app.displayName) : "Unknown";
  };

  return (
    <div
      ref={wrapperRef}
      className="d-flex flex-column align-items-center bg-secondary bg-opacity-10"
      style={{ width: '100%', flexGrow: 1, overflowY: 'auto', minHeight: 0 }}
    >
      <Document
        file={fileUrl}
        onLoadSuccess={onDocumentLoadSuccess}
        loading={<div className="p-5">Loading Document...</div>}
        error={<div className="alert alert-danger">Failed to load PDF.</div>}
      >
        {Array.from(new Array(numPages), (el, index) => {
          const pageNum = index + 1;
          const pagePlaceholders = placeholders.filter(p => p.page === pageNum);
          const pageGridBoxes = gridBoxes.filter(b => b.page === pageNum);
          const pageTexts = texts.filter(t => t.page === pageNum);

          return (
            <div
              key={`page_${pageNum}`}
              data-page={pageNum}
              ref={(el) => {
                if (el) pageRefs.current.set(pageNum, el);
                else pageRefs.current.delete(pageNum);
              }}
              className={`mb-4 shadow position-relative bg-white ${isDragging ? 'touch-drop-active' : ''}`}
              onDrop={(e) => handleDrop(e, pageNum)}
              onDragOver={handleDragOver}
              onMouseDown={() => onSelect(null)}
              onTouchEnd={(e) => handleTouchTap(e, pageNum)}
              style={{ width: pdfWrapperWidth, cursor: isDragging ? 'crosshair' : 'default' }}
            >
              <Page pageNumber={pageNum} renderTextLayer={false} renderAnnotationLayer={false} width={pdfWrapperWidth} />

            <div
                style={{
                    position: 'absolute',
                    bottom: '10px',
                    left: '10px',
                    backgroundColor: 'rgba(255,255,255,0.9)',
                    border: '1px solid #ddd',
                    borderRadius: '4px',
                    padding: '2px 8px',
                    fontSize: '12px',
                    color: '#666',
                    zIndex: 9999
                }}
            >
                {pageNum} / {numPages}
            </div>

              {/* Render Grid Boxes (burned into the PDF on Process Workflow) */}
              {pageGridBoxes.map(b => {
                const isSelected = selectedId === b.id;
                return (
                  <div
                    key={b.id}
                    className="position-absolute"
                    style={{
                      left: `${b.x}%`,
                      top: `${b.y}%`,
                      width: `${b.width}%`,
                      height: `${b.height}%`,
                      zIndex: 9,
                      outline: `${isSelected ? 2 : 1}px dashed ${isSelected ? '#0d6efd' : 'rgba(108, 117, 125, 0.8)'}`,
                      outlineOffset: 3,
                      cursor: 'move',
                      touchAction: 'none'
                    }}
                    onMouseDown={(e) => handleMouseDown(e, b.id, 'move', b, 'grid')}
                    onTouchStart={(e) => handleTouchStart(e, b.id, 'move', b, 'grid')}
                  >
                    <GridBoxPreview box={b} />

                    {/* Delete Button */}
                    <div
                      className="position-absolute bg-white rounded-circle shadow border d-flex align-items-center justify-content-center"
                      style={{ top: -12, right: -12, width: 20, height: 20, fontSize: 12, color: 'red', zIndex: 20, cursor: 'pointer' }}
                      onMouseDown={(e) => e.stopPropagation()}
                      onTouchStart={(e) => e.stopPropagation()}
                      onClick={(e) => { e.stopPropagation(); onRemoveGridBox(b.id); }}
                      role="button"
                      title="Remove"
                    >×</div>

                    {/* Resize Handle (Bottom Right) */}
                    <div
                      className="position-absolute d-flex align-items-end justify-content-end"
                      style={{ right: 0, bottom: 0, width: 25, height: 25, cursor: 'se-resize', zIndex: 15, padding: '3px', touchAction: 'none' }}
                      onMouseDown={(e) => handleMouseDown(e, b.id, 'resize', b, 'grid')}
                      onTouchStart={(e) => handleTouchStart(e, b.id, 'resize', b, 'grid')}
                    >
                      <div style={{ width: 0, height: 0, borderBottom: '10px solid #6c757d', borderLeft: '10px solid transparent' }} />
                    </div>
                  </div>
                );
              })}

              {/* Render Free Texts (burned into the PDF on Process Workflow) */}
              {pageTexts.map(t => {
                const isSelected = selectedId === t.id;
                return (
                  <div
                    key={t.id}
                    className="position-absolute"
                    style={{
                      left: `${t.x}%`,
                      top: `${t.y}%`,
                      zIndex: 12,
                      fontFamily: 'Helvetica, Arial, sans-serif',
                      fontSize: (t.size / 100) * pdfWrapperWidth,
                      lineHeight: FREE_TEXT_LINE_HEIGHT,
                      color: t.color,
                      whiteSpace: 'pre',
                      outline: `${isSelected ? 2 : 1}px dashed ${isSelected ? '#0d6efd' : 'rgba(108, 117, 125, 0.6)'}`,
                      outlineOffset: 2,
                      cursor: 'move',
                      touchAction: 'none',
                      userSelect: 'none'
                    }}
                    onMouseDown={(e) => handleMouseDown(e, t.id, 'move', getElementRect(e.currentTarget, t), 'text')}
                    onTouchStart={(e) => handleTouchStart(e, t.id, 'move', getElementRect(e.currentTarget, t), 'text')}
                  >
                    {t.text.trim() ? t.text : <span style={{ color: '#adb5bd', fontStyle: 'italic' }}>Empty text</span>}

                    {/* Delete Button */}
                    <div
                      className="position-absolute bg-white rounded-circle shadow border d-flex align-items-center justify-content-center"
                      style={{ top: -12, right: -12, width: 18, height: 18, fontSize: 11, lineHeight: 1, color: 'red', zIndex: 20, cursor: 'pointer', fontFamily: 'sans-serif' }}
                      onMouseDown={(e) => e.stopPropagation()}
                      onTouchStart={(e) => e.stopPropagation()}
                      onClick={(e) => { e.stopPropagation(); onRemoveText(t.id); }}
                      role="button"
                      title="Remove"
                    >×</div>

                    {/* Resize Handle (Bottom Right): drag to make the text bigger or smaller */}
                    <div
                      className="position-absolute"
                      style={{ right: -7, bottom: -7, width: 12, height: 12, backgroundColor: '#fff', border: '2px solid #0d6efd', borderRadius: 2, cursor: 'se-resize', zIndex: 20, touchAction: 'none' }}
                      onMouseDown={(e) => handleMouseDown(e, t.id, 'resize', getElementRect(e.currentTarget.parentElement as HTMLElement, t), 'text')}
                      onTouchStart={(e) => handleTouchStart(e, t.id, 'resize', getElementRect(e.currentTarget.parentElement as HTMLElement, t), 'text')}
                      title="Drag to resize"
                    />
                  </div>
                );
              })}

              {/* Render Placeholders */}
              {pagePlaceholders.map(p => {
                const isInitial = p.type === 'initial';
                const mainColor = isInitial ? '#0000ff' : '#0000ff'; // Blue vs Green
                const mainColorBorder = isInitial ? 'rgba(45, 137, 239, 0.8)' : 'rgba(0, 163, 0, 0.8)'; // Blue vs Green
                const bgColor = 'transparent'; // Transparent Box
                const labelBg = isInitial ? 'rgba(45, 137, 239, 0.8)' : 'rgba(0, 163, 0, 0.8)'; // Transparent Label
                const labelPrefix = isInitial ? 'Init' : 'Sig';

                return (
                  <div
                    key={p.id}
                    className={`position-absolute shadow-sm`}
                    style={{
                      left: `${p.x}%`,   // Percentage
                      top: `${p.y}%`,    // Percentage
                      width: `${p.width}%`, // Percentage
                      height: `${p.height}%`, // Percentage
                      zIndex: 10,
                      border: `2px dashed ${mainColorBorder}`,
                      backgroundColor: bgColor,
                      overflow: 'visible',
                      cursor: 'move',
                      touchAction: 'none' // Prevent scroll while dragging
                    }}
                    onMouseDown={(e) => handleMouseDown(e, p.id, 'move', p)}
                    onTouchStart={(e) => handleTouchStart(e, p.id, 'move', p)}
                  >
                    {/* Floating Label: Top-Left (Outside) */}
                    <div
                      className="position-absolute text-white px-2 py-0 shadow-sm"
                      style={{
                        top: -20,
                        left: -2,
                        backgroundColor: labelBg,
                        fontSize: '9px',
                        lineHeight: '18px',
                        borderTopLeftRadius: '4px',
                        borderTopRightRadius: '4px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      {labelPrefix}: {getApproverName(p.approverId)}
                    </div>

                    {/* Delete Button: Top-Right (Outside or inside corner) */}
                    <div
                      className="position-absolute bg-white rounded-circle shadow border d-flex align-items-center justify-content-center cursor-pointer"
                      style={{ top: -8, right: -8, width: 20, height: 20, fontSize: 12, color: 'red', zIndex: 20, cursor: 'pointer' }}
                      onClick={(e) => { e.stopPropagation(); onRemovePlaceholder(p.id); }}
                      role="button"
                      title="Remove"
                    >×</div>


                    {/* Conditional Date Label (Inside Top-Centered) */}
                    {p.checklistDate && (
                      <div
                        className="position-absolute"
                        style={{
                          top: Math.min(5, Math.max(0, (p.width / 100 * pdfWrapperWidth) * 0.05 - 2)),
                          left: '50%',
                          transform: 'translateX(-50%)',
                          fontSize: `${Math.min(14, Math.max(7, (p.width / 100 * pdfWrapperWidth) * 0.06))}px`,
                          color: 'black',
                          fontWeight: 600,
                          zIndex: 11,
                          whiteSpace: 'nowrap',
                          textAlign: 'center'
                        }}
                      >
                        {new Date().toLocaleDateString('en-GB')}
                      </div>
                    )}

                    {/* Conditional Name/Title/Badge Label (Inside Bottom-Centered or Outside if small) */}
                    {(p.checklistName || p.checklistBadge) && (
                      <div
                        className="position-absolute"
                        style={{
                          bottom: (p.width / 100 * pdfWrapperWidth) < 100 ? -5 : 5,
                          left: '50%',
                          transform: (p.width / 100 * pdfWrapperWidth) < 100 ? 'translate(-50%, 100%)' : 'translate(-50%, 0)',
                          fontSize: `${Math.min(14, Math.max(9, (p.width / 100 * pdfWrapperWidth) * 0.035))}px`,
                          color: 'black',
                          fontWeight: 600,
                          zIndex: 11,
                          whiteSpace: 'nowrap',
                          textAlign: 'center'
                        }}
                      >
                        {(() => {
                          const title = getApproverName(p.approverId);
                          const app = approvers.find(a => a.id === p.approverId);
                          const badge = formatBadgeNumber(app?.badgeNumber || "-");

                          return (
                            <div className="d-flex flex-column align-items-center">
                              {p.checklistName && <div>{title}</div>}
                              {p.checklistBadge && <div>{badge}</div>}
                            </div>
                          );
                        })()}
                      </div>
                    )}

                    {/* Body: Signature Preview Image */}
                    <div className="d-flex align-items-center justify-content-center h-100 w-100" style={{ pointerEvents: 'none' }}>
                      {isInitial ? (
                        <svg viewBox="0 0 200 100" width="100%" height="100%" preserveAspectRatio="none">
                          <path d="M20,50 Q50,20 80,50 T140,50 T180,40" stroke={mainColor} strokeWidth="5" fill="none" strokeLinecap="round" />
                        </svg>
                      ) : (
                        <svg viewBox="0 0 200 100" width="100%" height="100%" preserveAspectRatio="none">
                          <path d="M10,50 Q40,10 60,50 T110,50 T160,20 M120,60 Q150,80 180,40" stroke={mainColor} strokeWidth="5" fill="none" strokeLinecap="round" />
                        </svg>
                      )}
                    </div>

                    {/* Resize Handle Visual (Bottom Right) */}
                    <div
                      className="position-absolute bottom-0 end-0 d-flex align-items-end justify-content-end"
                      style={{
                        width: 25,
                        height: 25,
                        cursor: 'se-resize',
                        zIndex: 15,
                        padding: '3px',
                        touchAction: 'none' // Prevent scroll while resizing
                      }}
                      onMouseDown={(e) => handleMouseDown(e, p.id, 'resize', p)}
                      onTouchStart={(e) => handleTouchStart(e, p.id, 'resize', p)}
                    >
                      {/* Visual triangle */}
                      <div style={{
                        width: 0,
                        height: 0,
                        borderBottom: '10px solid #ccc',
                        borderLeft: '10px solid transparent'
                      }} />
                    </div>
                  </div>
                );
              })}
            </div >
          );
        })}
      </Document >
    </div >
  );
};

export default PdfViewer;
