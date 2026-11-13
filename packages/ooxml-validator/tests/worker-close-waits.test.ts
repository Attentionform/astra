import { describe, expect, it } from 'vitest';
import { EventEmitter } from 'node:events';
import { WorkerClient } from '../src/worker-client';

/** A child that was sent a signal (`killed`) but has not exited yet: close() must wait for its exit. */
function signalledButAliveChild() {
  const child = new EventEmitter() as EventEmitter & Record<string, unknown>;
  child.killed = true;
  child.exitCode = null;
  child.signalCode = null;
  child.ref = () => undefined;
  child.kill = () => true;
  child.stdin = { write: () => true };
  return child;
}

describe('WorkerClient.close', () => {
  it('waits for a child that was already signalled but has not exited', async () => {
    const client = new WorkerClient();
    const child = signalledButAliveChild();
    (client as unknown as { liveChildren: Set<unknown> }).liveChildren.add(child);
    let closed = false;
    const closing = client.close().then(() => { closed = true; });
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(closed).toBe(false);
    child.exitCode = 0;
    child.emit('exit', 0, null);
    await closing;
    expect(closed).toBe(true);
  });

  it('does not wait for a child that has already exited', async () => {
    const client = new WorkerClient();
    const child = signalledButAliveChild();
    child.exitCode = 0;
    (client as unknown as { liveChildren: Set<unknown> }).liveChildren.add(child);
    await client.close();
  });
});
