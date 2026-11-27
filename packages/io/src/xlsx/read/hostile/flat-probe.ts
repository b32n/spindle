import { flat } from '../capped';

// Run by `hostile-memory.test.ts` in a child process with `--expose-gc`: what a string cut from a large one keeps alive, with and without
// `flat` (argv[2]: `view` or `flat`; argv[3]: `narrow` or `wide`, the text being one byte or two bytes a character). Prints the memory
// that is still held once the large string is out of scope and has been collected.
const [mode, width] = process.argv.slice(2);
const mb = (n: number): number => Math.round(n / (1024 * 1024));
const gc = (globalThis as { gc?: () => void }).gc as () => void;
const kept: string[] = [];

function cut(): void {
  const bytes = new Uint8Array(100_000_000).fill(120);
  // 0xE2 0x82 0xAC is a euro sign: past Latin-1, so the decoded text takes two bytes a character
  if (width === 'wide') bytes.set([0xe2, 0x82, 0xac], 5);
  const big = new TextDecoder().decode(bytes);
  const view = big.slice(10, 23);
  kept.push(mode === 'flat' ? flat(view) : view);
}

async function main(): Promise<void> {
  cut();
  for (let i = 0; i < 3; i++) {
    gc();
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  const m = process.memoryUsage();
  console.log(JSON.stringify({ heapMb: mb(m.heapUsed), externalMb: mb(m.external), length: kept[0].length }));
}

void main();
