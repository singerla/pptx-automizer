/**
 * Builds the pinned renderer image (cached after the first run) and starts it
 * on a free host port. Workers inherit THUMBNAILER_URL via process.env; the
 * container id is kept for global-teardown.
 */
import { execSync } from 'child_process';
import * as net from 'net';
import * as path from 'path';

const RENDERER_IMAGE = 'pptx-renderer';
const RENDERER_CONTEXT = path.resolve(__dirname, '../../../tools/render-pptx');

const freePort = (): Promise<number> =>
  new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as net.AddressInfo;
      server.close(() => resolve(port));
    });
  });

const waitForHealthz = async (url: string, timeoutMs: number) => {
  const deadline = Date.now() + timeoutMs;
  let lastError = 'no response';
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${url}/healthz`, {
        signal: AbortSignal.timeout(2000),
      });
      if (response.ok) {
        return;
      }
      lastError = `healthz returned ${response.status}`;
    } catch (error) {
      lastError = (error as Error).message;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  throw new Error(
    `pptx-renderer container did not become healthy within ${timeoutMs}ms (${lastError})`,
  );
};

export default async function globalSetup(): Promise<void> {
  // First build downloads LibreOffice (~2 min); afterwards this is a cache hit
  execSync(`docker build -q -t ${RENDERER_IMAGE} "${RENDERER_CONTEXT}"`, {
    stdio: ['ignore', 'ignore', 'inherit'],
  });

  const port = await freePort();
  const containerId = execSync(
    `docker run -d --rm -p 127.0.0.1:${port}:3000 ${RENDERER_IMAGE}`,
    { encoding: 'utf-8' },
  ).trim();

  process.env.THUMBNAILER_URL = `http://127.0.0.1:${port}`;
  process.env.PPTX_RENDERER_CONTAINER = containerId;

  try {
    await waitForHealthz(process.env.THUMBNAILER_URL, 60000);
  } catch (error) {
    execSync(`docker rm -f ${containerId}`, { stdio: 'ignore' });
    throw error;
  }
}
