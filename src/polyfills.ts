// Polyfill Promise.withResolvers for browsers that don't have it natively (required by pdfjs-dist)
if (typeof Promise !== 'undefined' && typeof (Promise as unknown as { withResolvers?: unknown }).withResolvers === 'undefined') {
  (Promise as unknown as { withResolvers: <T>() => { promise: Promise<T>; resolve: (value: T | PromiseLike<T>) => void; reject: (reason?: unknown) => void } }).withResolvers = function <T>() {
    let resolve!: (value: T | PromiseLike<T>) => void;
    let reject!: (reason?: unknown) => void;
    const promise = new Promise<T>((res, rej) => {
      resolve = res;
      reject = rej;
    });
    return { promise, resolve, reject };
  };
}
