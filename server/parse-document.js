import { execFile } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const worker = fileURLToPath(new URL('./parse-worker.js', import.meta.url));
// Process isolation bounds JavaScript heap, execution time and returned text.
export function parseDocument(path, extension) {
  return new Promise((resolve, reject) => {
    execFile(process.execPath, ['--max-old-space-size=128', worker, path, extension], {
      timeout: 10000, killSignal: 'SIGKILL', maxBuffer: 2 * 1024 * 1024,
      env: { PATH: process.env.PATH, NODE_ENV: 'production' },
    }, (error, stdout) => error ? reject(new Error('Parser limit exceeded or invalid file')) : resolve(stdout));
  });
}
