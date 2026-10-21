import { spawn, type ChildProcess } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';

const REPO = path.resolve(__dirname, '../..');

/** An unused TCP port (best effort: bound and released immediately). */
export function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.on('error', reject);
    srv.listen(0, '127.0.0.1', () => {
      const { port } = srv.address() as net.AddressInfo;
      srv.close(() => resolve(port));
    });
  });
}

/**
 * A TCP proxy in front of the shared relay: the test can cut every connection
 * (`drop()`), or refuse new ones for a while (`stop()` / `start()`), which is what a
 * network outage looks like to the browser, while the relay keeps its room state.
 * The proxy keeps its port across stop/start so a reconnecting client finds it again.
 */
export class NetworkProxy {
  private server: net.Server | null = null;
  private sockets = new Set<net.Socket>();
  port = 0;

  constructor(private readonly targetPort = 1234) {}

  get url(): string {
    return `ws://127.0.0.1:${this.port}`;
  }

  async start(): Promise<void> {
    if (this.server) return;
    if (!this.port) this.port = await freePort();
    const server = net.createServer((client) => {
      const upstream = net.connect(this.targetPort, '127.0.0.1');
      for (const s of [client, upstream]) {
        this.sockets.add(s);
        s.on('close', () => this.sockets.delete(s));
        s.on('error', () => s.destroy());
      }
      client.on('close', () => upstream.destroy());
      upstream.on('close', () => client.destroy());
      client.pipe(upstream);
      upstream.pipe(client);
    });
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(this.port, '127.0.0.1', resolve);
    });
    this.server = server;
  }

  /** Cut every open connection; new connections are still accepted. */
  drop(): void {
    for (const s of this.sockets) s.destroy();
  }

  /** Cut every connection and refuse new ones until `start()`. */
  async stop(): Promise<void> {
    const server = this.server;
    this.server = null;
    this.drop();
    if (server) await new Promise<void>((resolve) => server.close(() => resolve()));
  }
}

/**
 * The real reference relay (examples/collab-server) as a child process the test owns, so it
 * can be killed and restarted. A restart loses the room's in-memory log, like a crash would.
 */
const live = new Set<ChildProcess>();
/** Kill every relay a test started (call from afterEach so a failing test cannot leak one). */
export function killAllRelays(): void {
  for (const c of live) c.kill('SIGKILL');
  live.clear();
}
process.once('exit', killAllRelays);

export class RelayProcess {
  private child: ChildProcess | null = null;
  port = 0;

  get url(): string {
    return `ws://127.0.0.1:${this.port}`;
  }

  async start(): Promise<void> {
    if (!this.port) this.port = await freePort();
    const child = spawn(
      process.execPath,
      ['--import', 'tsx', 'examples/collab-server/src/server.ts'],
      {
        cwd: REPO,
        env: { ...process.env, PORT: String(this.port) },
        stdio: ['ignore', 'pipe', 'inherit'],
      }
    );
    this.child = child;
    live.add(child);
    child.once('exit', () => live.delete(child));
    await new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('relay did not start')), 20_000);
      child.once('exit', (code) => reject(new Error(`relay exited early (${code})`)));
      child.stdout!.on('data', (d: Buffer) => {
        if (d.toString().includes('listening')) {
          clearTimeout(timer);
          child.removeAllListeners('exit');
          resolve();
        }
      });
    });
  }

  async kill(): Promise<void> {
    const child = this.child;
    this.child = null;
    if (!child || child.exitCode !== null) return;
    await new Promise<void>((resolve) => {
      child.once('exit', () => resolve());
      child.kill('SIGKILL');
    });
  }
}
