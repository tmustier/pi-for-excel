import type {
  Credential,
  CredentialInfo,
  CredentialStore,
} from "@earendil-works/pi-ai";

export interface ProviderKeysStoreLike {
  get(provider: string): Promise<string | null>;
  set(provider: string, key: string): Promise<void>;
  delete(provider: string): Promise<void>;
  list(): Promise<string[]>;
}

/**
 * Adapts Pi for Excel's existing IndexedDB API-key store to Pi AI's runtime
 * credential contract. Browser OAuth flows continue to refresh their grants in
 * the taskpane and publish the effective access token through ProviderKeysStore.
 */
export class ProviderCredentialsStore implements CredentialStore {
  private readonly providerKeys: ProviderKeysStoreLike;
  private readonly chains = new Map<string, Promise<void>>();

  constructor(providerKeys: ProviderKeysStoreLike) {
    this.providerKeys = providerKeys;
  }

  async read(providerId: string): Promise<Credential | undefined> {
    const key = await this.providerKeys.get(providerId);
    return key ? { type: "api_key", key } : undefined;
  }

  async list(): Promise<readonly CredentialInfo[]> {
    const providers = await this.providerKeys.list();
    return providers.map((providerId) => ({ providerId, type: "api_key" }));
  }

  async modify(
    providerId: string,
    fn: (current: Credential | undefined) => Promise<Credential | undefined>,
  ): Promise<Credential | undefined> {
    let result: Credential | undefined;
    const previous = this.chains.get(providerId) ?? Promise.resolve();
    const next = previous
      .catch(() => {})
      .then(async () => {
        const current = await this.read(providerId);
        const updated = await fn(current);
        result = updated ?? current;

        if (updated?.type === "oauth") {
          throw new Error("Browser OAuth credentials must be persisted by the taskpane OAuth store.");
        }
        if (updated?.type === "api_key" && typeof updated.key === "string") {
          await this.providerKeys.set(providerId, updated.key);
        }
      });

    this.chains.set(providerId, next);
    try {
      await next;
      return result;
    } finally {
      if (this.chains.get(providerId) === next) {
        this.chains.delete(providerId);
      }
    }
  }

  async delete(providerId: string): Promise<void> {
    const previous = this.chains.get(providerId) ?? Promise.resolve();
    const next = previous
      .catch(() => {})
      .then(() => this.providerKeys.delete(providerId));

    this.chains.set(providerId, next);
    try {
      await next;
    } finally {
      if (this.chains.get(providerId) === next) {
        this.chains.delete(providerId);
      }
    }
  }
}
