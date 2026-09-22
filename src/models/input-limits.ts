import type { ModelInputLimits } from "@earendil-works/pi-ai";
import { Type } from "typebox";
import { Value } from "typebox/value";

const positiveInteger = Type.Integer({ minimum: 1 });

const modelInputLimitsSchema = Type.Object({
  maxRequestBytes: Type.Optional(positiveInteger),
  images: Type.Optional(Type.Object({
    resize: Type.Optional(Type.Object({
      maxWidth: Type.Optional(positiveInteger),
      maxHeight: Type.Optional(positiveInteger),
      maxBytes: Type.Optional(positiveInteger),
      jpegQuality: Type.Optional(Type.Integer({ minimum: 1, maximum: 100 })),
    })),
    maxPerMessage: Type.Optional(positiveInteger),
    maxPerRequest: Type.Optional(positiveInteger),
  })),
});

/** Decode model input metadata from extension or persisted-storage boundaries. */
export function decodeModelInputLimits(value: object): ModelInputLimits | null {
  if (!Value.Check(modelInputLimitsSchema, value)) return null;
  const { maxRequestBytes, images } = value;
  return {
    ...(maxRequestBytes !== undefined ? { maxRequestBytes } : {}),
    ...(images !== undefined
      ? {
          images: {
            ...(images.resize !== undefined ? { resize: { ...images.resize } } : {}),
            ...(images.maxPerMessage !== undefined ? { maxPerMessage: images.maxPerMessage } : {}),
            ...(images.maxPerRequest !== undefined ? { maxPerRequest: images.maxPerRequest } : {}),
          },
        }
      : {}),
  };
}
