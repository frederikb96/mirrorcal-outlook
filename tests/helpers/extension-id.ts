/** The id Chromium derives from an extension's public key: SHA-256 of the DER, 16 bytes, a-p. */
export async function extensionIdFromKey(key: string): Promise<string> {
  const der = Uint8Array.from(atob(key), (c) => c.charCodeAt(0));
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', der));
  return Array.from(digest.slice(0, 16), (b) => b.toString(16).padStart(2, '0'))
    .join('')
    .replace(/[0-9a-f]/g, (c) => String.fromCharCode(97 + Number.parseInt(c, 16)));
}
