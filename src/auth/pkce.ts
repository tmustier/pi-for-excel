/**
 * Browser-safe PKCE utilities.
 *
 * Kept local so we do not depend on pi-ai package-internal deep imports.
 */

const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

const SHA256_H0 = new Uint32Array([
  0x6a09e667,
  0xbb67ae85,
  0x3c6ef372,
  0xa54ff53a,
  0x510e527f,
  0x9b05688c,
  0x1f83d9ab,
  0x5be0cd19,
]);

function base64urlEncode(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}

function rightRotate(value: number, bits: number): number {
  return (value >>> bits) | (value << (32 - bits));
}

function readPkceByte(bytes: Uint8Array, index: number): number {
  const value = bytes[index];
  if (value === undefined) {
    throw new Error("PKCE SHA-256 byte index out of bounds");
  }
  return value;
}

function readSha256Word(words: Uint32Array, index: number): number {
  const value = words[index];
  if (value === undefined) {
    throw new Error("PKCE SHA-256 word index out of bounds");
  }
  return value;
}

function sha256Fallback(bytes: Uint8Array): Uint8Array {
  const paddedLength = Math.ceil((bytes.length + 9) / 64) * 64;
  const padded = new Uint8Array(paddedLength);
  padded.set(bytes);
  padded[bytes.length] = 0x80;

  const bitLengthHigh = Math.floor(bytes.length / 0x20000000);
  const bitLengthLow = (bytes.length << 3) >>> 0;
  const lengthOffset = paddedLength - 8;
  padded[lengthOffset] = (bitLengthHigh >>> 24) & 0xff;
  padded[lengthOffset + 1] = (bitLengthHigh >>> 16) & 0xff;
  padded[lengthOffset + 2] = (bitLengthHigh >>> 8) & 0xff;
  padded[lengthOffset + 3] = bitLengthHigh & 0xff;
  padded[lengthOffset + 4] = (bitLengthLow >>> 24) & 0xff;
  padded[lengthOffset + 5] = (bitLengthLow >>> 16) & 0xff;
  padded[lengthOffset + 6] = (bitLengthLow >>> 8) & 0xff;
  padded[lengthOffset + 7] = bitLengthLow & 0xff;

  const h = new Uint32Array(SHA256_H0);
  const w = new Uint32Array(64);

  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let i = 0; i < 16; i++) {
      const j = offset + i * 4;
      w[i] = (
        (readPkceByte(padded, j) << 24) |
        (readPkceByte(padded, j + 1) << 16) |
        (readPkceByte(padded, j + 2) << 8) |
        readPkceByte(padded, j + 3)
      ) >>> 0;
    }

    for (let i = 16; i < 64; i++) {
      const wordMinus15 = readSha256Word(w, i - 15);
      const wordMinus2 = readSha256Word(w, i - 2);
      const s0 = (rightRotate(wordMinus15, 7) ^ rightRotate(wordMinus15, 18) ^ (wordMinus15 >>> 3)) >>> 0;
      const s1 = (rightRotate(wordMinus2, 17) ^ rightRotate(wordMinus2, 19) ^ (wordMinus2 >>> 10)) >>> 0;
      w[i] = (readSha256Word(w, i - 16) + s0 + readSha256Word(w, i - 7) + s1) >>> 0;
    }

    let a = readSha256Word(h, 0);
    let b = readSha256Word(h, 1);
    let c = readSha256Word(h, 2);
    let d = readSha256Word(h, 3);
    let e = readSha256Word(h, 4);
    let f = readSha256Word(h, 5);
    let g = readSha256Word(h, 6);
    let hash = readSha256Word(h, 7);

    for (let i = 0; i < 64; i++) {
      const s1 = (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25)) >>> 0;
      const ch = ((e & f) ^ (~e & g)) >>> 0;
      const temp1 = (hash + s1 + ch + readSha256Word(SHA256_K, i) + readSha256Word(w, i)) >>> 0;
      const s0 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22)) >>> 0;
      const maj = ((a & b) ^ (a & c) ^ (b & c)) >>> 0;
      const temp2 = (s0 + maj) >>> 0;

      hash = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    h[0] = (readSha256Word(h, 0) + a) >>> 0;
    h[1] = (readSha256Word(h, 1) + b) >>> 0;
    h[2] = (readSha256Word(h, 2) + c) >>> 0;
    h[3] = (readSha256Word(h, 3) + d) >>> 0;
    h[4] = (readSha256Word(h, 4) + e) >>> 0;
    h[5] = (readSha256Word(h, 5) + f) >>> 0;
    h[6] = (readSha256Word(h, 6) + g) >>> 0;
    h[7] = (readSha256Word(h, 7) + hash) >>> 0;
  }

  const digest = new Uint8Array(32);
  for (let i = 0; i < h.length; i++) {
    const value = readSha256Word(h, i);
    const offset = i * 4;
    digest[offset] = (value >>> 24) & 0xff;
    digest[offset + 1] = (value >>> 16) & 0xff;
    digest[offset + 2] = (value >>> 8) & 0xff;
    digest[offset + 3] = value & 0xff;
  }

  return digest;
}

export async function sha256ForPkce(bytes: Uint8Array): Promise<Uint8Array> {
  const subtle = globalThis.crypto?.subtle;
  if (subtle?.digest) {
    try {
      const input = new ArrayBuffer(bytes.byteLength);
      new Uint8Array(input).set(bytes);
      return new Uint8Array(await subtle.digest("SHA-256", input));
    } catch {
      // WPS/dev HTTP can expose crypto.getRandomValues without usable subtle crypto.
    }
  }

  return sha256Fallback(bytes);
}

export async function generatePKCE(byteLength: number = 32): Promise<{ verifier: string; challenge: string }> {
  const verifierBytes = new Uint8Array(byteLength);
  crypto.getRandomValues(verifierBytes);
  const verifier = base64urlEncode(verifierBytes);

  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const hashBytes = await sha256ForPkce(data);
  const challenge = base64urlEncode(hashBytes);

  return { verifier, challenge };
}
