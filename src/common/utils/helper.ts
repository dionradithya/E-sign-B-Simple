
export function dataURLToBlob(dataURL: string): Blob {
  const parts = dataURL.split(",");
  const byteString = atob(parts[1]);
  const mimeString = parts[0].match(/:(.*?);/)?.[1] || "image/png";
  const ab = new ArrayBuffer(byteString.length);
  const ia = new Uint8Array(ab);
  for (let i = 0; i < byteString.length; i++) {
    ia[i] = byteString.charCodeAt(i);
  }
  return new Blob([ab], { type: mimeString });
}

export function sanitizeApproverName(displayName: string): string {
  if (!displayName) return "";
  return displayName.replace(/\s*\(.*?\)\s*/g, '').trim();
}

export function getEncodedFolderUrl(fileRef: string, origin: string): string {
  if (!fileRef) return "";

  const lastSlashIndex = fileRef.lastIndexOf('/');
  const folderPath = lastSlashIndex > -1 ? fileRef.substring(0, lastSlashIndex) : "";

  if (!folderPath) return "";

  const encodedFolder = folderPath.split('/').map((s: string) => encodeURIComponent(s)).join('/');
  return origin + encodedFolder;
}

export function formatBadgeNumber(badge: string): string {
  if (!badge) return "";
  return badge.toUpperCase().startsWith("U") ? badge.substring(1) : badge;
}
