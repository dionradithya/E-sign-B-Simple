export const generatePDFHash = async (fileBlob: Blob): Promise<string> => {
  // 1. Convert Blob/File ke ArrayBuffer
  const buffer = await fileBlob.arrayBuffer();

  // 2. Proses Hashing menggunakan Crypto API browser
  const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);

  // 3. Convert hasil buffer ke Hex String (biar bisa dibaca/disimpan sbg text)
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');

  return hashHex;
};
