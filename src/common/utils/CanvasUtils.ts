import * as React from "react";
import { IPathAnnotation, ISignaturePlaceholder } from "../models/IEsignState";

/**
 * Utility functions for canvas operations and coordinate calculations
 */
export class CanvasUtils {
    
    /**
     * Converts mouse/touch events to relative coordinates within a container
     * 
     * @param e - Mouse or Touch event
     * @param container - HTML element container
     * @returns Relative x and y coordinates
     */
    public static getRelativeCoords(
        e: React.MouseEvent | React.TouchEvent | MouseEvent | TouchEvent,
        container: HTMLElement
    ): { x: number; y: number } {
        const rect = container.getBoundingClientRect();
        let clientX: number;
        let clientY: number;
        
        // Check for touches
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const touchEvt = e as any;

        if (touchEvt.touches && touchEvt.touches.length > 0) {
            clientX = touchEvt.touches[0].clientX;
            clientY = touchEvt.touches[0].clientY;
        } else if (touchEvt.changedTouches && touchEvt.changedTouches.length > 0) {
            clientX = touchEvt.changedTouches[0].clientX;
            clientY = touchEvt.changedTouches[0].clientY;
        } else {
            // Fallback to mouse
            clientX = (e as React.MouseEvent).clientX;
            clientY = (e as React.MouseEvent).clientY;
        }

        return {
            x: clientX - rect.left,
            y: clientY - rect.top
        };
    }

    /**
     * Renders annotation paths on a canvas element
     * 
     * @param canvas - Canvas element to draw on
     * @param paths - Array of path annotations
     * @param pageNum - Page number to filter paths
     */
    public static drawPathsOnCanvas(
        canvas: HTMLCanvasElement,
        paths: IPathAnnotation[],
        pageNum: number
    ): void {
        const ctx = canvas.getContext('2d');
        if (!ctx) return;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        
        const pagePaths = paths.filter(p => p.page === pageNum);
        pagePaths.forEach(path => {
            if (path.points.length < 2) return;
            
            ctx.beginPath();
            ctx.strokeStyle = path.color;
            ctx.lineWidth = path.thickness;
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
            
            // De-normalize coordinates
            const start = path.points[0];
            ctx.moveTo(start.x * canvas.width, start.y * canvas.height);
            
            for (let i = 1; i < path.points.length; i++) {
                const p = path.points[i];
                ctx.lineTo(p.x * canvas.width, p.y * canvas.height);
            }
            ctx.stroke();
        });
    }

    /**
     * Creates a composite signature image with name, date, and badge overlays
     * 
     * @param base64Sig - Base64 encoded signature image
     * @param placeholder - Signature placeholder configuration
     * @param currentUserId - Current user's ID for badge
     * @returns Promise resolving to composite signature as base64 data URL
     */
    public static async createCompositeSignature(
        base64Sig: string,
        placeholder: ISignaturePlaceholder,
        currentUserId?: string
    ): Promise<string> {
        // If no badges required, return original
        if (!placeholder.includeName && !placeholder.includeDate && !placeholder.includeBadge) {
            return base64Sig;
        }

        return new Promise((resolve) => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            if (!ctx) return resolve(base64Sig);

            const img = new Image();
            img.onload = () => {
                // 1. Determine Canvas Size
                // Use a high-resolution base size
                const baseWidth = 600;
                
                 // Determine aspect ratio from placeholder if possible, or use signature image ratio
                const imgRatio = img.width / img.height;
                const placeholderRatio = placeholder.width / placeholder.height;
                
                // If placeholder is landscape (wider), use that ratio. Else use image ratio or standard 3:2
                const targetRatio = placeholderRatio > 1 ? placeholderRatio : imgRatio;

                const canvasWidth = baseWidth;
                const canvasHeight = baseWidth / targetRatio;

                canvas.width = canvasWidth;
                canvas.height = canvasHeight;

                // 2. Prepare Text
                const fontSize = 36;
                const isLandscape = targetRatio > 1; // Use targetRatio to determine landscape
                
                // Dynamic padding: 8% of height for landscape (min 20px), 10px constant for portrait
                const padding = isLandscape ? Math.max(20, canvasHeight * 0.08) : 1;

                
                let topText = "";
                if (placeholder.includeDate) {
                    topText = new Date().toLocaleDateString('en-GB', { 
                        day: '2-digit', 
                        month: '2-digit', 
                        year: 'numeric' 
                    });
                }

                const nameText = (placeholder.includeName && placeholder.approverName) ? placeholder.approverName : "";
                const badgeText = (placeholder.includeBadge && currentUserId) ? currentUserId : "";

                // Set font for measurement
                ctx.font = `${fontSize}px Arial`;



                // 3. Draw Signature (Maximize Size, Allow Overlap)
                // We want the signature to populate the box as much as possible
                
                // Fill vertically or horizontally? 
                // "image itu penuh" -> Cover/Contain with bias towards full size
                // Standard contain logic:
                let drawW = canvasWidth;
                let drawH = canvasWidth / imgRatio;

                if (drawH > canvasHeight) {
                    drawH = canvasHeight;
                    drawW = drawH * imgRatio;
                }

                // Center it
                const sigX = (canvasWidth - drawW) / 2;
                const sigY = (canvasHeight - drawH) / 2;
                
                // Draw
                ctx.drawImage(img, sigX, sigY, drawW, drawH);

                // 4. Draw Text Overlay (On top of image - "gkpp kena")
                ctx.fillStyle = '#000';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                
                const centerX = canvasWidth / 2;

                // Top Text (Date)
                if (topText) {
                    ctx.font = `${fontSize}px Arial`;
                    // Position at absolute top with padding
                    ctx.fillText(topText, centerX, padding + (fontSize / 2));
                }

                // Bottom Text 1: Name (Bold)
                // If badge is present, we need to shift the Name UP to make room for the Badge below it
                const badgeShift = (placeholder.includeBadge && currentUserId) ? 5 : 0;

                if (nameText) {
                    ctx.font = `700 ${fontSize}px Arial`; // Bold
                    // Position at absolute bottom with padding, shifted up if badge exists
                    ctx.fillText(nameText, centerX, canvasHeight - padding - (fontSize / 2) - badgeShift);
                }

                // Bottom Text 2: Badge/ID (Regular/Smaller) - Positioned BELOW Name
                if (badgeText) {
                    const badgeFontSize = 28; 
                    const offset = 32; // Distance from Name center to Badge center
                    
                    ctx.font = `${badgeFontSize}px Arial`; 
                    // Calculate based on shifted Name position
                    const nameY = canvasHeight - padding - (fontSize / 2) - badgeShift;
                    ctx.fillText(badgeText, centerX, nameY + offset);
                }

                resolve(canvas.toDataURL('image/png')); 
            };
            img.src = base64Sig;
        });
    }
}
