import { PDFDocument, rgb, degrees, PDFPage, BlendMode } from 'pdf-lib';
import { IPathAnnotation, ISignaturePlaceholder, ITextAnnotation } from '../models/IEsignState';

export class PdfUtils {

    /**
     * Burns annotations (paths, texts, signatures) into the PDF.
     * Handles coordinate transformation based on page rotation.
     */
    public static async burnAnnotations(
        fileUrl: string,
        paths: IPathAnnotation[],
        texts: ITextAnnotation[],
        placeholders: ISignaturePlaceholder[],
        burnSignatures: boolean
    ): Promise<Blob> {
        if (!fileUrl) throw new Error("File URL is missing");

        // 1. Load PDF
        const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
        const pdfDoc = await PDFDocument.load(existingPdfBytes);
        const pages = pdfDoc.getPages();

        // 2. Draw Paths (Drawing)
        paths.forEach(p => {
            if (p.page > pages.length) return;
            const page = pages[p.page - 1];
            
            // Convert hex color to RGB (0-1 range for pdf-lib)
            const hexToRgb = (hex: string): { r: number, g: number, b: number } => {
                const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
                if (result) {
                    return {
                        r: parseInt(result[1], 16) / 255,
                        g: parseInt(result[2], 16) / 255,
                        b: parseInt(result[3], 16) / 255
                    };
                }
                // Fallback to red if invalid hex
                return { r: 1, g: 0, b: 0 };
            };
            
            const pathColor = hexToRgb(p.color);
            
            // Normalize path points based on rotation
            if (p.points && p.points.length > 0) {
                 // For paths, we simpler transform each point
                 const pdfPoints = p.points.map(pt => this._transformCoordinate(pt.x, pt.y, page));

                 for (let i = 0; i < pdfPoints.length - 1; i++) {
                    page.drawLine({
                        start: pdfPoints[i],
                        end: pdfPoints[i + 1],
                        thickness: p.thickness,
                        color: rgb(pathColor.r, pathColor.g, pathColor.b),
                    });
                }
            }
        });

        // 3. Draw Text
        texts.forEach(t => {
            if (t.page > pages.length) return;
            const page = pages[t.page - 1];
            const { width, height } = page.getCropBox() ?? page.getMediaBox(); // Physical size
            const rotation = page.getRotation().angle;

            // Calculate Font Size (Visual -> Physical)
            // t.size is % of Visual Width.
            // If Rot=0/180: VisWidth = PhysWidth. Size = (t.size/100) * W.
            // If Rot=90/270: VisWidth = PhysHeight. Size = (t.size/100) * H.
            let visualWidth = width;
            if (rotation === 90 || rotation === 270) {
                visualWidth = height;
            }
            const fontSize = (t.size / 100) * visualWidth;

            // Anchor logic for Text:
            // Text is drawn from Bottom-Left (ignoring descenders usually)
            // But our (x,y) is Top-Left of the text box visually.
            // We need to adjust y by fontSize usually to get baseline.
            // And then rotate.
            
            // Let's rely on transformCoordinate for the anchor.
            // The coordinate t.x, t.y is the Top-Left of the text block visually.
            // For simple text, let's just transform the point and draw.
            // We might need to adjust for height of text if we want WYSIWYG precision for multiline/large text.
            // For now, let's use the transformation and rotation.
            
            const anchor = this._transformCoordinate(t.x, t.y, page);
            
            // Adjustment for text baseline vs Top-Left
            // Visual Top-Left (t.x, t.y).
            // PDF Text Anchor is Bottom-Left of the first char.
            // So we need to move "Down" visually by approx 80% of FontSize.
            // "Down" depends on rotation.
            
            // Rot 0: Down is -Y.
            // Rot 90: Down is +X.
            // Rot 180: Down is +Y.
            // Rot 270: Down is -X.

            let textX = anchor.x;
            let textY = anchor.y;
            const textOffset = fontSize * 0.8;

            if (rotation === 0) textY -= textOffset;
            else if (rotation === 90) textX += textOffset;
            else if (rotation === 180) textY += textOffset;
            else if (rotation === 270) textX -= textOffset;

            page.drawText(t.text, {
                x: textX,
                y: textY,
                size: fontSize,
                color: rgb(t.color === 'red' ? 1 : 0, 0, 0),
                rotate: degrees(rotation) // Match page rotation
            });
        });

        // 4. Burn Signatures
        if (burnSignatures) {
            for (const p of placeholders) {
                if (p.page > pages.length || !p.signatureData) continue;
                
                const page = pages[p.page - 1];
                const { width, height } = page.getCropBox() ?? page.getMediaBox();
                const rotation = page.getRotation().angle;

                if (p.signatureData.startsWith('data:image/') || p.signatureData.startsWith('data:application/octet-stream')) {
                    try {
                        const imageBytes = await fetch(p.signatureData).then(res => res.arrayBuffer());
                        let embeddedImage;
                        if (p.signatureData.startsWith('data:image/png') || p.signatureData.startsWith('data:application/octet-stream')) {
                             try {
                                embeddedImage = await pdfDoc.embedPng(imageBytes);
                             } catch {
                                embeddedImage = await pdfDoc.embedJpg(imageBytes);
                             }
                        } else {
                            embeddedImage = await pdfDoc.embedJpg(imageBytes);
                        }

                        if (embeddedImage) {
                            // Calculate Visual Dimensions
                            let visualPageWidth = width;
                            let visualPageHeight = height;
                            if (rotation === 90 || rotation === 270) {
                                visualPageWidth = height;
                                visualPageHeight = width;
                            }

                            const boxWidthVis = p.width * visualPageWidth;
                            const boxHeightVis = p.height * visualPageHeight;

                            // Scale Image to fit Box (Maintain Aspect Ratio)
                            const dims = embeddedImage.scaleToFit(boxWidthVis, boxHeightVis);



                            // We need the Top-Left of the IMAGE within the box.
                            // Image Width/Height in Visual terms:
                            const imgWVis = dims.width;
                            const imgHVis = dims.height;

                            // Visual Top-Left of the Centered Image
                            const imgXVis = p.x + ( (p.width * visualPageWidth - imgWVis) / 2 ) / visualPageWidth;
                            const imgYVis = p.y + ( (p.height * visualPageHeight - imgHVis) / 2 ) / visualPageHeight;

                            // Calculate the Physical Coordinates of the Image Corners to determine Anchor
                            // We need to pass x,y to drawImage such that it draws at the correct spot.
                            
                            // Visual Top-Left of Image
                            // const tl = this._transformCoordinate(imgXVis, imgYVis, page);
                            // Visual Bottom-Right of Image
                            // Visual Bottom-Right of Image
                            const imgBottomVis = imgYVis + (imgHVis / visualPageHeight);
                            
                            // We also need TR and BL for other rotations
                            const bl = this._transformCoordinate(imgXVis, imgBottomVis, page);

                            // Determine Anchor (See Plan)
                            // Rot 0: Anchor is BL.
                            // Rot 90: Anchor is BR.
                            // Rot 180: Anchor is TR.
                            // Rot 270: Anchor is TL.
                            
                            // Determine Anchor:
                            // We want the image to extend "Up" and "Right" visually from the Bottom-Left corner.
                            // When we rotate the image to match the page, the "Up" and "Right" axes of the image 
                            // align with the Visual Up and Right.
                            // Therefore, the anchor point passed to drawImage (at rotation=pageRotation)
                            // must ALWAYS be the physical coordinate corresponding to the Visual Bottom-Left corner.
                            
                            const anchorX = bl.x;
                            const anchorY = bl.y;

                            page.drawImage(embeddedImage, {
                                x: anchorX,
                                y: anchorY,
                                width: dims.width,
                                height: dims.height,
                                rotate: degrees(rotation),
                                blendMode: BlendMode.Multiply
                            });
                        }
                    } catch (e) {
                        console.error("Failed to embed image", e);
                    }
                }
            }
        }

        const pdfBytes = await pdfDoc.save();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return new Blob([pdfBytes as any], { type: 'application/pdf' });
    }

    /**
     * Transforms normalized visual coordinates (0-1) to PDF physical coordinates (points).
     * Handles 0, 90, 180, 270 degree page rotations.
     * 
     * @param normX - Normalized X coordinate (0 to 1) from Left
     * @param normY - Normalized Y coordinate (0 to 1) from Top
     * @param page - The PDFPage object
     * @returns { x: number, y: number } - Physical PDF coordinates
     */
    private static _transformCoordinate(normX: number, normY: number, page: PDFPage): { x: number, y: number } {
        const { x: boxX, y: boxY, width, height } = page.getCropBox() ?? page.getMediaBox();
        const rotation = page.getRotation().angle;

        let x = 0;
        let y = 0;

        // Visual Coordinates: Origin Top-Left (0-1), +X Right, +Y Down.
        // PDF Coordinates: Origin Bottom-Left of CropBox (boxX, boxY), +X Right, +Y Up.

        if (rotation === 0) {
            // Standard Portrait
            // Visual X (0->1) => Phys X (boxX -> boxX+width)
            // Visual Y (0->1) => Phys Y (boxY+height -> boxY)
            x = boxX + (normX * width);
            y = boxY + height - (normY * height);
        } else if (rotation === 90) {
            // 90 Deg Clockwise
            // Phys Left (boxX) -> Vis Top (0).
            // Phys Right (boxX+width) -> Vis Bottom (1).
            // Phys Bottom (boxY) -> Vis Left (0).
            // Phys Top (boxY+height) -> Vis Right (1).
            
            // x = boxX + (VisY * width)
            x = boxX + (normY * width);
            
            // y = boxY + (VisX * height)
            y = boxY + (normX * height);

        } else if (rotation === 180) {
            // 180 Deg Upside Down
            // Vis X (0->1) : Phys X (boxX+width -> boxX).
            x = boxX + width - (normX * width);

            // Vis Y (0->1) : Phys Y (boxY -> boxY+height).
            y = boxY + (normY * height);

        } else if (rotation === 270) {
            // 270 Deg CW (90 CCW)
            // Vis X (0->1) : Phys Y (boxY+height -> boxY).
            y = boxY + height - (normX * height);

            // Vis Y (0->1) : Phys X (boxX+width -> boxX).
            x = boxX + width - (normY * width);
        }

        return { x, y };
    }
}
