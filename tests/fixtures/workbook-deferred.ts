export function deferred<T>() {
  let resolve: (value: T) => void = () => { throw new Error("Deferred promise not initialized"); };
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}
