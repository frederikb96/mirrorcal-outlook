import { describe, expect, test } from 'bun:test';
import { createSingleFlight } from '../src/extension/single-flight.ts';

function gate() {
  let open!: () => void;
  const promise = new Promise<void>((resolve) => {
    open = resolve;
  });
  return { promise, open };
}

describe('single flight', () => {
  test('runs one task at a time and reports busy to requests that do not coalesce', async () => {
    const flight = createSingleFlight<string>();
    const g = gate();
    const first = flight.run(async () => {
      await g.promise;
      return 'first';
    });
    expect(await flight.run(async () => 'second')).toEqual({ busy: true });
    g.open();
    expect(await first).toEqual({ busy: false, value: 'first' });
    expect(await flight.run(async () => 'third')).toEqual({ busy: false, value: 'third' });
  });

  test('requests arriving mid-run share exactly one follow-up run', async () => {
    const flight = createSingleFlight<number>();
    let runs = 0;
    const g = gate();
    const task = async () => {
      runs += 1;
      if (runs === 1) await g.promise;
      return runs;
    };
    const first = flight.run(task, { coalesce: true });
    const second = flight.run(task, { coalesce: true });
    const third = flight.run(task, { coalesce: true });
    g.open();
    expect(await first).toEqual({ busy: false, value: 1 });
    expect(await second).toEqual({ busy: false, value: 2 });
    expect(await third).toEqual({ busy: false, value: 2 });
    expect(runs).toBe(2);
  });

  test('a request after the follow-up started gets its own run', async () => {
    const flight = createSingleFlight<number>();
    let runs = 0;
    const gates = [gate(), gate()];
    const task = async () => {
      const index = runs++;
      await gates[index]?.promise;
      return index;
    };
    const first = flight.run(task, { coalesce: true });
    const queued = flight.run(task, { coalesce: true });
    gates[0]?.open();
    await first;
    // The follow-up is running now; a new request must queue behind it, not join it.
    const later = flight.run(async () => 99, { coalesce: true });
    gates[1]?.open();
    expect(await queued).toEqual({ busy: false, value: 1 });
    expect(await later).toEqual({ busy: false, value: 99 });
  });

  test('a failing task rejects its caller and frees the flight', async () => {
    const flight = createSingleFlight<string>();
    await expect(
      flight.run(async () => {
        throw new Error('boom');
      }),
    ).rejects.toThrow('boom');
    expect(await flight.run(async () => 'ok')).toEqual({ busy: false, value: 'ok' });
  });
});
