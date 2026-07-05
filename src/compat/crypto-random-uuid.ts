const UUID_BYTE_LENGTH = 16;

function byteToHex(byte: number): string {
  return byte.toString(16).padStart(2, "0");
}

function createRandomUuid(getRandomValues: Crypto["getRandomValues"]): ReturnType<Crypto["randomUUID"]> {
  const bytes = new Uint8Array(UUID_BYTE_LENGTH);
  getRandomValues.call(globalThis.crypto, bytes);

  // RFC 4122 version 4 UUID: set version and variant bits.
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const hex = Array.from(bytes, byteToHex);
  return `${hex.slice(0, 4).join("")}-${hex.slice(4, 6).join("")}-${hex
    .slice(6, 8)
    .join("")}-${hex.slice(8, 10).join("")}-${hex.slice(10).join("")}`;
}

export function installCryptoRandomUuidPatch(): void {
  const cryptoObject = globalThis.crypto;
  if (!cryptoObject || typeof cryptoObject.randomUUID === "function") return;
  if (typeof cryptoObject.getRandomValues !== "function") return;

  Object.defineProperty(cryptoObject, "randomUUID", {
    configurable: true,
    value: () => createRandomUuid(cryptoObject.getRandomValues),
  });
}
