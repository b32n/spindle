// Run by `child.test.ts` in a child process, to give `runChild` the endings a real probe can have (argv[2]): a result line, a line cut
// short, no output, a sleep past the time limit, a failure.
const mode = process.argv[2];
if (mode === 'json') console.log('noise\n{"a":1}');
else if (mode === 'gc') console.log(JSON.stringify({ a: typeof (globalThis as { gc?: unknown }).gc === 'function' ? 1 : 0 }));
else if (mode === 'cut') process.stdout.write('{"a":');
else if (mode === 'text') console.log('not json');
else if (mode === 'sleep') Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 10_000);
else if (mode === 'fail') {
  console.error('boom');
  process.exitCode = 3;
}
