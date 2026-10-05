export type Admission<T> = { busy: false; value: T } | { busy: true };

export interface RunOptions {
  /** When a run is in progress, queue one follow-up instead of reporting busy. */
  coalesce?: boolean;
}

interface Queued<T> {
  task: () => Promise<T>;
  promise: Promise<Admission<T>>;
  resolve: (value: Admission<T>) => void;
  reject: (reason: unknown) => void;
}

/**
 * Runs one task at a time. A request that arrives mid-run either reports busy or, when coalescing,
 * joins the single follow-up run that starts right after the current one; any number of
 * coalescing requests share that one follow-up.
 */
export function createSingleFlight<T>() {
  let active = false;
  let queued: Queued<T> | undefined;

  async function start(task: () => Promise<T>): Promise<Admission<T>> {
    active = true;
    try {
      return { busy: false, value: await task() };
    } finally {
      active = false;
      const next = queued;
      queued = undefined;
      if (next) start(next.task).then(next.resolve, next.reject);
    }
  }

  return {
    run(task: () => Promise<T>, options: RunOptions = {}): Promise<Admission<T>> {
      if (!active) return start(task);
      if (!options.coalesce) return Promise.resolve({ busy: true });
      if (!queued) {
        let resolve!: Queued<T>['resolve'];
        let reject!: Queued<T>['reject'];
        const promise = new Promise<Admission<T>>((res, rej) => {
          resolve = res;
          reject = rej;
        });
        queued = { task, promise, resolve, reject };
      }
      return queued.promise;
    },
  };
}
