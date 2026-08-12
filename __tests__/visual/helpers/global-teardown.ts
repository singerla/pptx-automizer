import { execSync } from 'child_process';

export default async function globalTeardown(): Promise<void> {
  const containerId = process.env.PPTX_RENDERER_CONTAINER;
  if (containerId) {
    // --rm on the run makes rm -f both stop and remove
    execSync(`docker rm -f ${containerId}`, { stdio: 'ignore' });
  }
}
