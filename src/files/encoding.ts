/**
 * Encoding helpers for workspace file content.
 */

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

export function encodeTextUtf8(text: string): Uint8Array {
  return textEncoder.encode(text);
}

export function decodeTextUtf8(bytes: Uint8Array): string {
  return textDecoder.decode(bytes);
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x4000;

  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    for (const byte of chunk) {
      binary += String.fromCharCode(byte);
    }
  }

  return btoa(binary);
}

export function base64ToBytes(base64: string): Uint8Array {
  const normalized = base64.trim();
  if (normalized.length === 0) {
    return new Uint8Array(0);
  }

  const binary = atob(normalized);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes;
}

export function truncateText(text: string, maxChars: number): {
  text: string;
  truncated: boolean;
} {
  if (maxChars <= 0) {
    return { text: "", truncated: text.length > 0 };
  }

  if (text.length <= maxChars) {
    return { text, truncated: false };
  }

  return {
    text: text.slice(0, maxChars),
    truncated: true,
  };
}

export function truncateBase64(base64: string, maxChars: number): {
  base64: string;
  truncated: boolean;
} {
  if (maxChars <= 0) {
    return { base64: "", truncated: base64.length > 0 };
  }

  if (base64.length <= maxChars) {
    return { base64, truncated: false };
  }

  return {
    base64: base64.slice(0, maxChars),
    truncated: true,
  };
}
