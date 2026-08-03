import QrCodeWithLogo from 'qrcode-with-logos';
import { logoBase64 as logoImage } from '../../webparts/initiateSignature/assets/logoBase64';

/**
 * Generates a QR code image with an embedded logo as a Uint8Array (PNG bytes).
 * @param content - The text/URL to encode in the QR code.
 * @returns Uint8Array of PNG image bytes.
 */
export async function generateQrCodeImageBytes(content: string): Promise<Uint8Array> {
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = 300;
    tempCanvas.height = 300;

    const qrcode = new QrCodeWithLogo({
        canvas: tempCanvas,
        content: content,
        width: 300,
        nodeQrCodeOptions: {
            errorCorrectionLevel: 'H',
            margin: 2,
            color: {
                dark: '#000000',
                light: '#FFFFFF'
            }
        },
        logo: {
            src: logoImage,
            bgColor: '#FFFFFF',
            borderWidth: 5,
            borderRadius: 15,
            logoRadius: 15
        }
    });

    const drawnCanvas = await qrcode.getCanvas();
    const qrBase64 = drawnCanvas.toDataURL('image/png');
    const base64Data = qrBase64.split(',')[1];
    const binaryString = atob(base64Data);
    const len = binaryString.length;
    const qrImageBytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        qrImageBytes[i] = binaryString.charCodeAt(i);
    }

    return qrImageBytes;
}
