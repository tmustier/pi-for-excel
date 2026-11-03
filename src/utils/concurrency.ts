export type ConcurrencyLimit = <T>(task: () => Promise<T>) => Promise<T>;

export function createConcurrencyLimit(maxConcurrent: number): ConcurrencyLimit {
  let active = 0;
  const waiting: (() => void)[] = [];

  return async (task) => {
    if (active < maxConcurrent) active += 1;
    else await new Promise<void>((resolve) => waiting.push(resolve));
    try {
      return await task();
    } finally {
      const next = waiting.shift();
      if (next) next();
      else active -= 1;
    }
  };
}

export function mapWithConcurrency<T, R>(
  items: readonly T[],
  maxConcurrent: number,
  map: (item: T) => Promise<R>,
): Promise<R[]> {
  const limit = createConcurrencyLimit(maxConcurrent);
  return Promise.all(items.map((item) => limit(() => map(item))));
}
